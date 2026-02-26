import { describe, expect, it } from "vitest";
import PizZip from "pizzip";
import { aggregateWeeklyRows } from "../src/transformer/aggregation";
import { createDocx } from "../src/transformer/docx";
import {
  readWorkAreaMapFromCsv,
  readWorklogRowsFromCsv,
} from "../src/transformer/csv";

const noopLog = () => {};

function makeTemplateArrayBuffer(): ArrayBuffer {
  const zip = new PizZip();
  zip.file(
    "word/document.xml",
    `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
  <w:body>
    <w:p><w:r><w:t>Tabelle</w:t></w:r></w:p>
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
});

describe("transformer DOCX table rendering", () => {
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
      `Key,Name,Alias
TEST-1,Team Alpha,TA
TEST-2,Unused Team,UT`,
    );

    const docx = createDocx({
      templateArrayBuffer: makeTemplateArrayBuffer(),
      worklogRows: rows,
      workAreasByKey: workAreas,
      weekly: false,
      includeLegend: true,
    });
    const xml = readDocumentXml(docx);

    expect(xml).toContain("Bereich");
    expect(xml).toContain("Legende");
    expect(xml).toContain("TA");
    expect(xml).toContain("Team Alpha");
    expect(xml).not.toContain(">UT<");
    expect(xml).not.toContain(">Unused Team<");
  });
});
