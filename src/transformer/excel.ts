import * as XLSX from "xlsx-js-style";
import dayjs from "./dayjs";
import type { WorkAreaEntry, WorklogRow } from "./types";

const BLACK_BORDER_STYLE = {
  top: { style: "thin", color: { rgb: "000000" } },
  bottom: { style: "thin", color: { rgb: "000000" } },
  left: { style: "thin", color: { rgb: "000000" } },
  right: { style: "thin", color: { rgb: "000000" } },
};

function formatGermanDate(date: Date): string {
  return dayjs(date).format("DD.MM.YYYY");
}

function resolveWorkAreaValue(
  row: WorklogRow,
  workAreasByKey: Map<string, WorkAreaEntry> | null,
): string {
  if (!workAreasByKey) return "";
  const keys: string[] = [];
  if (row.key) keys.push(row.key);
  if (Array.isArray(row.keys)) keys.push(...row.keys);

  const values: string[] = [];
  const seen = new Set<string>();
  for (const key of keys) {
    const entry = workAreasByKey.get(key);
    if (!entry) continue;
    const norm = entry.alias.toLowerCase();
    if (seen.has(norm)) continue;
    seen.add(norm);
    values.push(entry.alias);
  }

  return values.join("\n");
}

function getReferencedWorkAreas(options: {
  worklogRows: WorklogRow[];
  workAreasByKey: Map<string, WorkAreaEntry>;
}): Array<WorkAreaEntry> {
  const referencedKeys = new Set<string>();
  for (const row of options.worklogRows) {
    if (row.key) referencedKeys.add(row.key);
    if (Array.isArray(row.keys)) {
      for (const key of row.keys) {
        if (key) referencedKeys.add(key);
      }
    }
  }

  const entries: Array<WorkAreaEntry> = [];
  for (const [key, entry] of options.workAreasByKey.entries()) {
    if (referencedKeys.has(key)) {
      entries.push(entry);
    }
  }
  return entries;
}

export function createXlsx(options: {
  worklogRows: WorklogRow[];
  workAreasByKey: Map<string, WorkAreaEntry> | null;
  weekly: boolean;
  includeLegend: boolean;
}): Uint8Array {
  const showAreaColumn = Boolean(
    options.workAreasByKey && options.workAreasByKey.size > 0,
  );

  const header = showAreaColumn
    ? [
        options.weekly ? "Kalenderwoche" : "Datum",
        "Mitarbeiter*in",
        "Stunden",
        "Bereich",
        "Beschreibung der Tätigkeit",
      ]
    : [
        options.weekly ? "Kalenderwoche" : "Datum",
        "Mitarbeiter*in",
        "Stunden",
        "Beschreibung der Tätigkeit",
      ];

  const rows: Array<Array<string | number>> = [header];
  for (const row of options.worklogRows) {
    const rawDateText =
      row.dateValue instanceof Date
        ? formatGermanDate(row.dateValue)
        : String(row.dateValue);
    const dateText =
      options.weekly && typeof row.dateValue === "string"
        ? rawDateText.replace(/^(\d+)\s+\(/, "$1\n(")
        : rawDateText;

    const rowValues: Array<string | number> = [dateText, row.user, row.hours];
    if (showAreaColumn) {
      rowValues.push(resolveWorkAreaValue(row, options.workAreasByKey));
    }
    rowValues.push(row.description);
    rows.push(rowValues);
  }

  const summaryRowIndex1Based = rows.length + 1;
  const lastDataRowIndex1Based = Math.max(2, rows.length);
  if (showAreaColumn) {
    rows.push(["", "Summe", "", "", ""]);
  } else {
    rows.push(["", "Summe", "", ""]);
  }

  let legendHeaderRow1Based: number | null = null;
  let legendLastRow1Based: number | null = null;

  const referencedLegendEntries =
    options.includeLegend && options.workAreasByKey
      ? getReferencedWorkAreas({
          worklogRows: options.worklogRows,
          workAreasByKey: options.workAreasByKey,
        })
      : [];
  if (referencedLegendEntries.length > 0) {
    rows.push(new Array(header.length).fill(""));
    legendHeaderRow1Based = rows.length + 1;
    rows.push([
      "Legende",
      "",
      ...new Array(Math.max(0, header.length - 2)).fill(""),
    ]);
    for (const entry of referencedLegendEntries) {
      rows.push([
        entry.alias,
        entry.name,
        ...new Array(Math.max(0, header.length - 2)).fill(""),
      ]);
    }
    legendLastRow1Based = rows.length;
  }

  const worksheet = XLSX.utils.aoa_to_sheet(rows);

  if (options.weekly) {
    const merges: Array<{
      s: { r: number; c: number };
      e: { r: number; c: number };
    }> = [];
    let start = 0;
    while (start < options.worklogRows.length) {
      const current = options.worklogRows[start];
      if (!current) break;
      const key = current.dateKey;
      let end = start;
      while (end + 1 < options.worklogRows.length) {
        const next = options.worklogRows[end + 1];
        if (!next || next.dateKey !== key) break;
        end += 1;
      }
      if (end > start) {
        merges.push({
          s: { r: start + 1, c: 0 },
          e: { r: end + 1, c: 0 },
        });
      }
      start = end + 1;
    }

    if (merges.length > 0) {
      worksheet["!merges"] = merges;
    }
  }

  const proportions = showAreaColumn ? [14, 16, 10, 13, 47] : [16, 20, 12, 52];
  const totalProportion = proportions.reduce((a, b) => a + b, 0);
  const totalChars = 120;
  worksheet["!cols"] = proportions.map((p) => ({
    wch: Math.max(8, Math.round((p / totalProportion) * totalChars)),
  }));

  worksheet[`C${summaryRowIndex1Based}`] = {
    t: "n",
    f: `SUM(C2:C${lastDataRowIndex1Based})`,
    v: 0,
    z: "0.00",
  };

  const worklogTableRowsCount = summaryRowIndex1Based;
  const tableColsCount = header.length;
  for (let r = 1; r <= worklogTableRowsCount; r++) {
    for (let c = 0; c < tableColsCount; c++) {
      const cellAddress = XLSX.utils.encode_cell({ r: r - 1, c });
      const cell = worksheet[cellAddress] ?? { t: "s", v: "" };
      cell.s = {
        ...(cell.s ?? {}),
        border: BLACK_BORDER_STYLE,
      };
      if (
        options.weekly &&
        c === 0 &&
        r >= 2 &&
        r <= 1 + options.worklogRows.length
      ) {
        cell.s = {
          ...(cell.s ?? {}),
          alignment: {
            ...(cell.s?.alignment ?? {}),
            wrapText: true,
            vertical: "center",
            horizontal: "center",
          },
        };
      }
      worksheet[cellAddress] = cell;
    }
  }

  if (legendHeaderRow1Based && legendLastRow1Based) {
    for (let r = legendHeaderRow1Based; r <= legendLastRow1Based; r++) {
      for (let c = 0; c < 2; c++) {
        const cellAddress = XLSX.utils.encode_cell({ r: r - 1, c });
        const cell = worksheet[cellAddress] ?? { t: "s", v: "" };
        cell.s = {
          ...(cell.s ?? {}),
          border: BLACK_BORDER_STYLE,
        };
        worksheet[cellAddress] = cell;
      }
    }
  }

  const workbook = XLSX.utils.book_new();
  XLSX.utils.book_append_sheet(workbook, worksheet, "Timesheet");
  const output = XLSX.write(workbook, {
    bookType: "xlsx",
    type: "array",
  }) as ArrayBuffer;

  return new Uint8Array(output);
}
