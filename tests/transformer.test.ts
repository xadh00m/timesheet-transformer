import { afterAll, beforeAll, describe, expect, it, vi } from "vitest";
import PizZip from "pizzip";
import { aggregateWeeklyRows } from "../src/transformer/aggregation";
import { createDocx } from "../src/transformer/docx";
import { createXlsx } from "../src/transformer/excel";
import {
  readWorkAreaMapFromCsv,
  readWorklogRowsFromCsv,
} from "../src/transformer/csv";
import * as XLSX from "xlsx";

const noopLog = () => {};

function makeTemplateArrayBuffer(): ArrayBuffer {
  const zip = new PizZip();
  zip.file(
    "word/document.xml",
    `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
  <w:body>
    <w:p><w:r><w:rPr><w:b/><w:color w:val="FF0000"/></w:rPr><w:t>&lt;Datum&gt;</w:t></w:r></w:p>
    <w:p><w:r><w:rPr><w:i/><w:u w:val="single"/></w:rPr><w:t>&lt;Monat/Jahr&gt;</w:t></w:r></w:p>
    <w:p><w:r><w:t>&lt;Tabelle&gt;</w:t></w:r></w:p>
  </w:body>
</w:document>`,
  );
  const bytes = zip.generate({ type: "uint8array" });
  return bytes.buffer.slice(
    bytes.byteOffset,
    bytes.byteOffset + bytes.byteLength,
  ) as ArrayBuffer;
}

function readDocumentXml(bytes: Uint8Array): string {
  const zip = new PizZip(bytes);
  const file = zip.file("word/document.xml");
  if (!file) throw new Error("missing word/document.xml in generated docx");
  return file.asText();
}

describe("transformer worklog normalization and splitting", () => {
  it("normalizes and splits non-weekly worklog descriptions consistently", () => {
    const csv = `User,Worklog,Key,Logged,Date
Test User A,"- task alpha review
- task beta sync",TEST-1,1h,03/02/26 at 08:00`;

    const rows = readWorklogRowsFromCsv(csv, noopLog);

    expect(rows).toHaveLength(1);
    expect(rows[0]?.description).toBe("task alpha review, task beta sync");
    expect(rows[0]?.description).not.toContain(" - ");
  });

  it("aggregates weekly descriptions without leading dashes and with clean comma-separated items", () => {
    const csv = `User,Worklog,Key,Logged,Date
Test User A,- task gamma deployment,TEST-1,1h,03/02/26 at 08:00
Test User A,task delta rollout,TEST-1,2h,04/02/26 at 08:00
Test User A,"- task alpha review
- task beta sync",TEST-1,1h,05/02/26 at 08:00
Test User A,"- task epsilon tuning
- task zeta migration",TEST-1,1h,05/02/26 at 09:00`;

    const rows = readWorklogRowsFromCsv(csv, noopLog);
    const weekly = aggregateWeeklyRows(rows);

    expect(weekly).toHaveLength(1);
    expect(weekly[0]?.description).toBe(
      "task gamma deployment, task delta rollout, task alpha review, task beta sync, task epsilon tuning, task zeta migration",
    );
    expect(weekly[0]?.description).not.toContain(" - ");
  });

  it("orders non-weekly rows by date, user, and description", () => {
    const csv = `User,Worklog,Key,Logged,Date
Test User B,task b,TEST-1,1h,05/02/26 at 08:00
Test User A,task c,TEST-1,1h,05/02/26 at 09:00
Test User A,task a,TEST-1,1h,04/02/26 at 08:00`;

    const rows = readWorklogRowsFromCsv(csv, noopLog);

    expect(rows.map((row) => row.dateKey)).toEqual([
      "2026-02-04",
      "2026-02-05",
      "2026-02-05",
    ]);
    expect(rows.map((row) => row.user)).toEqual([
      "Test User A",
      "Test User A",
      "Test User B",
    ]);
    expect(rows.map((row) => row.description)).toEqual([
      "task a",
      "task c",
      "task b",
    ]);
  });

  it("orders weekly aggregates by week and user", () => {
    const csv = `User,Worklog,Key,Logged,Date
Test User B,task week7,TEST-1,1h,10/02/26 at 08:00
Test User A,task week6 a,TEST-1,1h,03/02/26 at 08:00
Test User B,task week6 b,TEST-1,1h,04/02/26 at 08:00`;

    const rows = readWorklogRowsFromCsv(csv, noopLog);
    const weekly = aggregateWeeklyRows(rows);

    expect(weekly.map((row) => row.dateKey)).toEqual([
      "2026-W06",
      "2026-W06",
      "2026-W07",
    ]);
    expect(weekly.map((row) => row.user)).toEqual([
      "Test User A",
      "Test User B",
      "Test User B",
    ]);
  });

  it("parses dot-separated dates with decimal-comma hours", () => {
    const csv = `User,Worklog,Key,Logged,Date
Andrii Hrohul,Dashboard work,SZ-40,"8,00",25.02.2026`;

    const rows = readWorklogRowsFromCsv(csv, noopLog);

    expect(rows).toHaveLength(1);
    expect(rows[0]?.hours).toBe(8);
    expect(rows[0]?.dateKey).toBe("2026-02-25");
  });
});

describe("transformer DOCX table rendering", () => {
  beforeAll(() => {
    vi.useFakeTimers();
    vi.setSystemTime(new Date("2026-03-09T12:00:00.000Z"));
  });

  afterAll(() => {
    vi.useRealTimers();
  });

  it("hides Bereich column and does not render legend when no work areas are provided", () => {
    const rows = readWorklogRowsFromCsv(
      `User,Worklog,Key,Logged,Date
Test User A,task a,TEST-1,1h,03/02/26 at 08:00`,
      noopLog,
    );

    const docx = createDocx({
      templateArrayBuffer: makeTemplateArrayBuffer(),
      worklogRows: rows,
      workAreasByKey: null,
      weekly: false,
      includeLegend: true,
    });
    const xml = readDocumentXml(docx);

    expect(xml).toContain("09.03.2026");
    expect(xml).not.toContain("&lt;Datum&gt;");
    expect(xml).toContain("<w:b/>");
    expect(xml).toContain("<w:color w:val=\"FF0000\"/>");

    expect(xml).toContain("02/2026");
    expect(xml).not.toContain("&lt;Monat/Jahr&gt;");
    expect(xml).toContain("<w:i/>");
    expect(xml).toContain("<w:u w:val=\"single\"/>");
    expect(xml).not.toContain("Bereich");
    expect(xml).not.toContain("Legende");
  });

  it("shows Bereich column and legend when enabled with work areas", () => {
    const rows = readWorklogRowsFromCsv(
      `User,Worklog,Key,Logged,Date
Test User A,task a,TEST-1,1h,03/02/26 at 08:00`,
      noopLog,
    );
    const workAreas = readWorkAreaMapFromCsv(
      `Key,Name,Alias,Rate
    TEST-1,Team Alpha,TA,125
    TEST-2,Unused Team,UT,125`,
    );

    const docx = createDocx({
      templateArrayBuffer: makeTemplateArrayBuffer(),
      worklogRows: rows,
      workAreasByKey: workAreas,
      weekly: false,
      includeLegend: true,
    });
    const xml = readDocumentXml(docx);

    expect(xml).toContain("09.03.2026");
    expect(xml).not.toContain("&lt;Datum&gt;");
    expect(xml).toContain("<w:b/>");
    expect(xml).toContain("<w:color w:val=\"FF0000\"/>");

    expect(xml).toContain("02/2026");
    expect(xml).not.toContain("&lt;Monat/Jahr&gt;");
    expect(xml).toContain("<w:i/>");
    expect(xml).toContain("<w:u w:val=\"single\"/>");
    expect(xml).toContain("Bereich");
    expect(xml).toContain("Legende");
    expect(xml).toContain("TA");
    expect(xml).toContain("Team Alpha");
    expect(xml).not.toContain(">UT<");
    expect(xml).not.toContain(">Unused Team<");
  });
});

describe("transformer work areas CSV parsing", () => {
  it("parses Rate column when present", () => {
    const areas = readWorkAreaMapFromCsv(
      `Key,Name,Alias,Rate,
SZ-83,"UX-Design","UI/UX",106.25,
SZ-84,"Entwicklung","Dev",125,
SZ-85,"IT-Security & privacy","IT-Sec","137,5",`,
    );

    expect(areas.get("SZ-83")?.rate).toBe(106.25);
    expect(areas.get("SZ-84")?.rate).toBe(125);
    expect(areas.get("SZ-85")?.rate).toBe(137.5);
  });

  it("warns when required cells are missing or invalid", () => {
    const warnings: string[] = [];
    const areas = readWorkAreaMapFromCsv(
      `Key,Name,Alias,Rate
SZ-1,Area 1,A1,
SZ-2,Area 2,A2,not-a-number
,,A3,100`,
      (line) => warnings.push(line),
    );

    expect(areas.get("SZ-1")?.rate).toBeNull();
    expect(areas.get("SZ-2")?.rate).toBeNull();
    expect(areas.has("")).toBe(false);
    expect(warnings.join("\n")).toContain("missing Rate");
    expect(warnings.join("\n")).toContain("invalid Rate");
    expect(warnings.join("\n")).toContain("missing Key");
  });
});

describe("transformer XLSX export", () => {
  it("creates summary hours cell as formula", () => {
    const rows = readWorklogRowsFromCsv(
      `User,Worklog,Key,Logged,Date
Test User A,task a,TEST-1,1h,03/02/26 at 08:00
Test User A,task b,TEST-1,2h,03/02/26 at 09:00`,
      noopLog,
    );

    const bytes = createXlsx({
      worklogRows: rows,
      workAreasByKey: null,
      weekly: false,
      includeLegend: false,
    });

    const workbook = XLSX.read(bytes, { type: "array", cellFormula: true });
    const sheet = workbook.Sheets[workbook.SheetNames[0] ?? ""];
    expect(sheet).toBeTruthy();

    const formulae = sheet ? XLSX.utils.sheet_to_formulae(sheet) : [];
    expect(formulae).toContain("C4=SUM(C2:C3)");
    expect(sheet?.B4?.v).toBe("Summe");
  });

  it("adds legend rows below table when selected", () => {
    const rows = readWorklogRowsFromCsv(
      `User,Worklog,Key,Logged,Date
Test User A,task a,TEST-1,1h,03/02/26 at 08:00`,
      noopLog,
    );
    const workAreas = readWorkAreaMapFromCsv(
      `Key,Name,Alias,Rate
    TEST-1,Team Alpha,TA,125
    TEST-2,Unused Team,UT,125`,
    );

    const bytes = createXlsx({
      worklogRows: rows,
      workAreasByKey: workAreas,
      weekly: false,
      includeLegend: true,
    });

    const workbook = XLSX.read(bytes, { type: "array", cellFormula: true });
    const sheet = workbook.Sheets[workbook.SheetNames[0] ?? ""];
    const values = sheet ? XLSX.utils.sheet_to_json(sheet, { header: 1 }) : [];
    const flat = (values as Array<Array<unknown>>).flat().map(String);

    expect(flat).toContain("Legende");
    expect(flat).toContain("TA");
    expect(flat).toContain("Team Alpha");
    expect(flat).not.toContain("UT");
    expect(flat).not.toContain("Unused Team");
  });

  it("does not add legend rows when includeLegend is false", () => {
    const rows = readWorklogRowsFromCsv(
      `User,Worklog,Key,Logged,Date
Test User A,task a,TEST-1,1h,03/02/26 at 08:00`,
      noopLog,
    );
    const workAreas = readWorkAreaMapFromCsv(
      `Key,Name,Alias,Rate
    TEST-1,Team Alpha,TA,125`,
    );

    const bytes = createXlsx({
      worklogRows: rows,
      workAreasByKey: workAreas,
      weekly: false,
      includeLegend: false,
    });

    const workbook = XLSX.read(bytes, { type: "array", cellFormula: true });
    const sheet = workbook.Sheets[workbook.SheetNames[0] ?? ""];
    const values = sheet ? XLSX.utils.sheet_to_json(sheet, { header: 1 }) : [];
    const flat = (values as Array<Array<unknown>>).flat().map(String);

    expect(flat).not.toContain("Legende");
    expect(flat).not.toContain("Team Alpha");
  });

  it("merges first-column cells for weekly rows with the same week", () => {
    const csv = `User,Worklog,Key,Logged,Date
Test User A,task a,TEST-1,1h,03/02/26 at 08:00
Test User B,task b,TEST-1,1h,04/02/26 at 08:00
Test User A,task c,TEST-1,1h,10/02/26 at 08:00`;
    const daily = readWorklogRowsFromCsv(csv, noopLog);
    const weeklyRows = aggregateWeeklyRows(daily);

    const bytes = createXlsx({
      worklogRows: weeklyRows,
      workAreasByKey: null,
      weekly: true,
      includeLegend: false,
    });

    const workbook = XLSX.read(bytes, { type: "array", cellFormula: true });
    const sheet = workbook.Sheets[workbook.SheetNames[0] ?? ""];
    const merges = sheet?.["!merges"] ?? [];
    expect(merges).toEqual(
      expect.arrayContaining([
        expect.objectContaining({
          s: { r: 1, c: 0 },
          e: { r: 2, c: 0 },
        }),
      ]),
    );
  });

  it("renders weekly first-column text with line break", () => {
    const csv = `User,Worklog,Key,Logged,Date
Test User A,task a,TEST-1,1h,03/02/26 at 08:00`;
    const daily = readWorklogRowsFromCsv(csv, noopLog);
    const weeklyRows = aggregateWeeklyRows(daily);

    const bytes = createXlsx({
      worklogRows: weeklyRows,
      workAreasByKey: null,
      weekly: true,
      includeLegend: false,
    });

    const workbook = XLSX.read(bytes, { type: "array", cellFormula: true });
    const sheet = workbook.Sheets[workbook.SheetNames[0] ?? ""];
    expect(String(sheet?.A2?.v ?? "")).toContain("\n");
  });
});
