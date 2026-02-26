import dayjs from "./dayjs";
import type { WorklogRow } from "./types";

function normalizeWorklogDescription(value: unknown): string {
  return String(value ?? "")
    .replace(/\r\n|\r|\n/g, " ")
    .replace(/^\s*-\s*/, "")
    .replace(/\s+/g, " ")
    .trim();
}

function getWorkWeekStart(date: Date): dayjs.Dayjs {
  const d = dayjs(date);
  const dow = d.day();
  const diffToMonday = (dow + 6) % 7;
  return d.subtract(diffToMonday, "day");
}

function getWorkWeekRange(date: Date) {
  const start = getWorkWeekStart(date)
    .hour(12)
    .minute(0)
    .second(0)
    .millisecond(0);
  const end = start.add(4, "day");
  const weekNumber = start.isoWeek();
  const weekYear = start.isoWeekYear();
  return {
    start: start.toDate(),
    end: end.toDate(),
    weekNumber,
    weekYear,
    key: `${weekYear}-W${String(weekNumber).padStart(2, "0")}`,
  };
}

export function compareWorklogRows(a: WorklogRow, b: WorklogRow): number {
  const byDate = a.dateSort - b.dateSort;
  if (byDate !== 0) return byDate;
  const byUser = a.user.localeCompare(b.user, "de");
  if (byUser !== 0) return byUser;
  return a.description.localeCompare(b.description, "de");
}

function splitWorklogParts(value: unknown): string[] {
  return String(value ?? "")
    .split(/(?:\r\n|\r|\n|[;,])+/g)
    .map((part) => normalizeWorklogDescription(part))
    .filter(Boolean);
}

export function aggregateWeeklyRows(rows: WorklogRow[]): WorklogRow[] {
  const groups = new Map<
    string,
    {
      weekStartKey: string;
      weekStartSort: number;
      weekNumber: number;
      weekRangeShort: string;
      user: string;
      hours: number;
      descriptions: string[];
      descriptionSet: Set<string>;
      keysSet: Set<string>;
    }
  >();

  for (const row of rows) {
    if (!(row.dateValue instanceof Date)) continue;
    const range = getWorkWeekRange(row.dateValue);
    const groupKey = `${range.key}::${row.user}`;
    let group = groups.get(groupKey);
    if (!group) {
      group = {
        weekStartKey: range.key,
        weekStartSort: dayjs(range.start).valueOf(),
        weekNumber: range.weekNumber,
        weekRangeShort: `(${dayjs(range.start).format("DD.MM")} - ${dayjs(range.end).format("DD.MM")})`,
        user: row.user,
        hours: 0,
        descriptions: [],
        descriptionSet: new Set<string>(),
        keysSet: new Set<string>(),
      };
      groups.set(groupKey, group);
    }

    group.hours += row.hours;
    const parts = splitWorklogParts(row.description);
    for (const part of parts) {
      const isDaily = /^daily\b/i.test(part);
      const dedupeKey = isDaily ? "daily" : part.toLowerCase();
      if (group.descriptionSet.has(dedupeKey)) continue;
      group.descriptionSet.add(dedupeKey);
      group.descriptions.push(isDaily ? "Daily" : part);
    }

    if (row.key) group.keysSet.add(row.key);
    if (Array.isArray(row.keys)) {
      for (const key of row.keys) if (key) group.keysSet.add(key);
    }
  }

  const aggregated: WorklogRow[] = [];
  for (const group of groups.values()) {
    aggregated.push({
      dateValue: `${group.weekNumber}\n${group.weekRangeShort}`,
      dateKey: group.weekStartKey,
      dateSort: group.weekStartSort,
      user: group.user,
      hours: Math.round(group.hours * 100) / 100,
      description: group.descriptions.join(", "),
      keys: Array.from(group.keysSet),
    });
  }

  return aggregated.sort(compareWorklogRows);
}
