import Papa from "papaparse";
import dayjs from "./dayjs";
import { compareWorklogRows } from "./aggregation";
import type { Logger, WorkAreaEntry, WorklogRow } from "./types";

function normalizeTextField(value: unknown): string {
  return String(value ?? "")
    .replace(/\r\n|\r|\n/g, " ")
    .replace(/\s+/g, " ")
    .trim();
}

function normalizeWorklogDescription(value: unknown): string {
  return normalizeTextField(value).replace(/^\s*-\s*/, "");
}

function splitWorklogParts(value: unknown): string[] {
  return String(value ?? "")
    .split(/(?:\r\n|\r|\n|[;,])+/g)
    .map((part) => normalizeWorklogDescription(part))
    .filter(Boolean);
}

function parseCsvRows(text: string): string[][] {
  const parsed = Papa.parse<string[]>(text, {
    skipEmptyLines: true,
  });
  if (parsed.errors.length > 0) {
    throw new Error(
      `CSV parse error: ${parsed.errors[0]?.message ?? "unknown"}`,
    );
  }
  return parsed.data;
}

function parseLoggedToHours(logged: unknown): number | null {
  const text = String(logged ?? "")
    .trim()
    .toLowerCase();
  if (!text) return null;
  const hourMatch = text.match(/(\d+(?:[.,]\d+)?)\s*h/);
  const minMatch = text.match(/(\d+(?:[.,]\d+)?)\s*m/);
  const hours = hourMatch?.[1] ? Number(hourMatch[1].replace(",", ".")) : 0;
  const minutes = minMatch?.[1] ? Number(minMatch[1].replace(",", ".")) : 0;
  const total = hours + minutes / 60;
  if (Number.isFinite(total) && total > 0) return Math.round(total * 100) / 100;
  const numeric = Number(text.replace(",", "."));
  if (!Number.isFinite(numeric) || numeric <= 0) return null;
  return Math.round(numeric * 100) / 100;
}

function parseWorklogDate(dateField: unknown): Date | null {
  const raw = String(dateField ?? "").trim();
  if (!raw) return null;
  const lower = raw.toLowerCase();
  if (lower.includes(" to ") || lower.includes(" bis ")) return null;

  const datePart = raw.split(/\s+(?:at|um)\s+/i)[0]?.trim() ?? "";
  const knownFormats = [
    "DD/MM/YY",
    "DD/MM/YYYY",
    "DD.MM.YY",
    "DD.MM.YYYY",
    "YYYY/MM/DD",
    "YYYY-MM-DD",
  ];
  const parsed = dayjs(datePart, knownFormats, true).hour(12);
  if (parsed.isValid()) return parsed.toDate();
  const parsedLoose = dayjs(datePart);
  if (!parsedLoose.isValid()) return null;
  return parsedLoose.hour(12).minute(0).second(0).millisecond(0).toDate();
}

export function readWorklogRowsFromCsv(
  csvContent: string,
  log: Logger,
): WorklogRow[] {
  const parsed = parseCsvRows(csvContent);
  if (parsed.length < 2) {
    throw new Error("CSV has no data rows");
  }

  const header = (parsed[0] ?? []).map((h) => String(h ?? "").trim());
  const findIndex = (name: string) =>
    header.findIndex((h) => h.toLowerCase() === name.toLowerCase());

  const idxUser = findIndex("User");
  const idxWorklog = findIndex("Worklog");
  const idxKey = findIndex("Key");
  const idxLogged = findIndex("Logged");
  const idxDate = findIndex("Date");

  const missing: string[] = [];
  if (idxUser < 0) missing.push("User");
  if (idxWorklog < 0) missing.push("Worklog");
  if (idxKey < 0) missing.push("Key");
  if (idxLogged < 0) missing.push("Logged");
  if (idxDate < 0) missing.push("Date");
  if (missing.length > 0) {
    throw new Error(`CSV header missing columns: ${missing.join(", ")}`);
  }

  const rows: WorklogRow[] = [];
  for (let recordIndex = 1; recordIndex < parsed.length; recordIndex++) {
    const record = parsed[recordIndex] ?? [];

    const user = normalizeTextField(record[idxUser]);
    const description = splitWorklogParts(record[idxWorklog]).join(", ");
    const key = normalizeTextField(record[idxKey]);
    const logged = normalizeTextField(record[idxLogged]);
    const dateField = normalizeTextField(record[idxDate]);

    const lineRef = `record ${recordIndex + 1}`;

    if (!user) {
      log(
        `Discarded CSV ${lineRef}: missing_user | ${JSON.stringify({ user, description, key, logged })}`,
      );
      continue;
    }
    if (!description) {
      log(
        `Discarded CSV ${lineRef}: missing_worklog/summary_row | ${JSON.stringify({ user, description, key, logged })}`,
      );
      continue;
    }

    const hours = parseLoggedToHours(logged);
    const date = parseWorklogDate(dateField);
    if (!hours || !date) {
      const reason =
        !hours && !date
          ? "invalid_logged_and_date"
          : !hours
            ? "invalid_logged"
            : "invalid_date";
      log(
        `Discarded CSV ${lineRef}: ${reason} | ${JSON.stringify({ user, description, key, logged, dateField })}`,
      );
      continue;
    }

    rows.push({
      dateValue: date,
      dateKey: dayjs(date).format("YYYY-MM-DD"),
      dateSort: dayjs(date).valueOf(),
      user,
      hours,
      description,
      ...(key ? { key } : null),
    });
  }

  return rows.sort(compareWorklogRows);
}

export function readWorkAreaMapFromCsv(
  csvContent: string,
  log?: Logger,
): Map<string, WorkAreaEntry> {
  const records = parseCsvRows(csvContent);
  if (records.length < 2) {
    throw new Error("work_areas.csv has no data rows");
  }

  const header = (records[0] ?? []).map((h) => String(h ?? "").trim());
  const findIndex = (name: string) =>
    header.findIndex((h) => h.toLowerCase() === name.toLowerCase());

  const idxKey = findIndex("Key");
  const idxName = findIndex("Name");
  const idxAlias = findIndex("Alias");
  const idxRate = findIndex("Rate");
  if (idxKey < 0 || idxName < 0 || idxAlias < 0 || idxRate < 0) {
    throw new Error(
      'work_areas.csv header must contain "Key", "Name", "Alias", and "Rate"',
    );
  }

  const map = new Map<string, WorkAreaEntry>();
  for (let recordIndex = 1; recordIndex < records.length; recordIndex++) {
    const record = records[recordIndex] ?? [];
    const key = normalizeTextField(record[idxKey]);
    const name = normalizeTextField(record[idxName]);
    const alias = normalizeTextField(record[idxAlias]);
    const rateRaw = normalizeTextField(record[idxRate]);
    const lineRef = `record ${recordIndex + 1}`;

    if (!key) {
      log?.(
        `Warning: discarded work_areas.csv ${lineRef}: missing Key | ${JSON.stringify({ key, name, alias, rateRaw })}`,
      );
      continue;
    }
    if (!name) {
      log?.(
        `Warning: discarded work_areas.csv ${lineRef}: missing Name | ${JSON.stringify({ key, name, alias, rateRaw })}`,
      );
      continue;
    }
    if (!alias) {
      log?.(
        `Warning: discarded work_areas.csv ${lineRef}: missing Alias | ${JSON.stringify({ key, name, alias, rateRaw })}`,
      );
      continue;
    }

    let resolvedRate: number | null = null;
    if (!rateRaw) {
      log?.(
        `Warning: work_areas.csv ${lineRef}: missing Rate | ${JSON.stringify({ key, name, alias, rateRaw })}`,
      );
    } else {
      const parsedRate = Number(rateRaw.replace(",", "."));
      if (!Number.isFinite(parsedRate) || parsedRate <= 0) {
        log?.(
          `Warning: work_areas.csv ${lineRef}: invalid Rate | ${JSON.stringify({ key, name, alias, rateRaw })}`,
        );
      } else {
        resolvedRate = parsedRate;
      }
    }

    map.set(key, { name, alias, rate: resolvedRate });
  }

  return map;
}
