import dayjs from "./dayjs";
import type { WorkAreaEntry, WorklogRow } from "./types";
import PizZip from "pizzip";

const DOCX_TABLE_FONT_SIZE_HALFPOINTS = 16;

function escapeXml(text: unknown): string {
  return String(text ?? "")
    .replace(/&/g, "&amp;")
    .replace(/</g, "&lt;")
    .replace(/>/g, "&gt;")
    .replace(/\"/g, "&quot;")
    .replace(/'/g, "&apos;");
}

function decodeXmlEntities(text: unknown): string {
  return String(text ?? "")
    .replace(/&lt;/g, "<")
    .replace(/&gt;/g, ">")
    .replace(/&quot;/g, '"')
    .replace(/&apos;/g, "'")
    .replace(/&amp;/g, "&");
}

function extractParagraphText(wpXml: string): string {
  const parts: string[] = [];
  const regex = /<w:t(?:\s[^>]*)?>([\s\S]*?)<\/w:t>/g;
  let match: RegExpExecArray | null;
  while ((match = regex.exec(wpXml)) !== null) {
    parts.push(decodeXmlEntities(match[1]));
  }
  return parts.join("");
}

function buildDocxTextRun(
  text: unknown,
  options: { bold?: boolean } = {},
): string {
  const safe = escapeXml(text);
  const rPr =
    "<w:rPr>" +
    (options.bold ? "<w:b/><w:bCs/>" : "") +
    `<w:sz w:val="${DOCX_TABLE_FONT_SIZE_HALFPOINTS}"/>` +
    `<w:szCs w:val="${DOCX_TABLE_FONT_SIZE_HALFPOINTS}"/>` +
    "</w:rPr>";
  return `<w:r>${rPr}<w:t xml:space="preserve">${safe}</w:t></w:r>`;
}

function buildDocxParagraph(
  text: unknown,
  options: { bold?: boolean; jc?: string | null; wrap?: boolean } = {},
): string {
  const alignment = options.jc ? `<w:jc w:val="${options.jc}"/>` : "";
  const pPr = alignment || options.wrap ? `<w:pPr>${alignment}</w:pPr>` : "";
  const lines = String(text ?? "").split("\n");
  const runs: string[] = [];
  for (let index = 0; index < lines.length; index++) {
    if (index > 0) runs.push("<w:r><w:br/></w:r>");
    runs.push(buildDocxTextRun(lines[index], { bold: options.bold }));
  }
  return `<w:p>${pPr}${runs.join("")}</w:p>`;
}

function buildDocxCell(options: {
  text: unknown;
  widthDxa?: number;
  bold?: boolean;
  jc?: string | null;
  vAlign?: string | null;
  vMerge?: "restart" | "continue" | null;
}): string {
  const tcW = options.widthDxa
    ? `<w:tcW w:w="${options.widthDxa}" w:type="dxa"/>`
    : "";
  const vA = options.vAlign ? `<w:vAlign w:val="${options.vAlign}"/>` : "";
  const vm = options.vMerge
    ? options.vMerge === "restart"
      ? '<w:vMerge w:val="restart"/>'
      : "<w:vMerge/>"
    : "";
  const tcPr = tcW || vA || vm ? `<w:tcPr>${tcW}${vA}${vm}</w:tcPr>` : "";
  const p = buildDocxParagraph(options.text, {
    bold: options.bold,
    jc: options.jc,
    wrap: true,
  });
  return `<w:tc>${tcPr}${p}</w:tc>`;
}

function buildSummaryRowXml(options: {
  grid: number[];
  sumHours: number;
  showAreaColumn: boolean;
}): string {
  const cells: string[] = [
    buildDocxCell({
      text: "",
      widthDxa: options.grid[0],
      bold: true,
      vAlign: "center",
    }),
    buildDocxCell({
      text: "Summe",
      widthDxa: options.grid[1],
      bold: true,
      vAlign: "center",
    }),
    buildDocxCell({
      text: formatGermanNumber(options.sumHours, { decimals: 2 }),
      widthDxa: options.grid[2],
      bold: true,
      jc: "right",
      vAlign: "center",
    }),
  ];

  if (options.showAreaColumn) {
    cells.push(
      buildDocxCell({
        text: "",
        widthDxa: options.grid[3],
        bold: true,
        vAlign: "center",
      }),
    );
    cells.push(
      buildDocxCell({
        text: "",
        widthDxa: options.grid[4],
        bold: true,
        vAlign: "center",
      }),
    );
  } else {
    cells.push(
      buildDocxCell({
        text: "",
        widthDxa: options.grid[3],
        bold: true,
        vAlign: "center",
      }),
    );
  }

  return `<w:tr>${cells.join("")}</w:tr>`;
}

function buildWorkAreaExplanationTableXml(
  workAreasByKey: Map<string, WorkAreaEntry>,
): string {
  const fontXml = '<w:rPr><w:sz w:val="16"/><w:szCs w:val="16"/></w:rPr>';
  const rows: string[] = [];
  for (const entry of workAreasByKey.values()) {
    rows.push(
      "<w:tr>" +
        `<w:tc><w:tcPr><w:tcW w:w="2500" w:type="dxa"/></w:tcPr><w:p><w:r>${fontXml}<w:t>${escapeXml(entry.alias)}</w:t></w:r></w:p></w:tc>` +
        `<w:tc><w:tcPr><w:tcW w:w="4500" w:type="dxa"/></w:tcPr><w:p><w:r>${fontXml}<w:t>${escapeXml(entry.name)}</w:t></w:r></w:p></w:tc>` +
        "</w:tr>",
    );
  }

  return (
    "<w:tbl>" +
    "<w:tblPr>" +
    '<w:tblStyle w:val="TableGrid"/>' +
    '<w:tblW w:w="7000" w:type="dxa"/>' +
    '<w:tblLayout w:type="fixed"/>' +
    "</w:tblPr>" +
    rows.join("") +
    "</w:tbl>"
  );
}

function formatGermanDate(date: Date): string {
  return dayjs(date).format("DD.MM.YYYY");
}

function formatGermanNumber(
  value: number,
  options: { decimals?: number } = {},
): string {
  const num = Number(value);
  if (!Number.isFinite(num)) return "";
  if (typeof options.decimals === "number") {
    return num.toFixed(options.decimals).replace(".", ",");
  }
  return String(num).replace(".", ",");
}

function resolveWorkAreaValue(
  row: WorklogRow,
  workAreasByKey: Map<string, WorkAreaEntry> | null,
): string | null {
  if (!workAreasByKey) return null;
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

  return values.length > 0 ? values.join("\n") : null;
}

function buildWorklogDocxTable(options: {
  worklogRows: WorklogRow[];
  workAreasByKey: Map<string, WorkAreaEntry> | null;
  weekly: boolean;
}): string {
  const showAreaColumn = Boolean(
    options.workAreasByKey && options.workAreasByKey.size > 0,
  );
  const proportions = showAreaColumn ? [14, 16, 10, 13, 47] : [16, 20, 12, 52];
  const totalProportion = proportions.reduce((a, b) => a + b, 0);
  const grid = proportions.map((p) => Math.round((p / totalProportion) * 9000));

  const tblGrid = `<w:tblGrid>${grid.map((w) => `<w:gridCol w:w="${w}"/>`).join("")}</w:tblGrid>`;
  const tblPr =
    "<w:tblPr>" +
    '<w:tblStyle w:val="TableGrid"/>' +
    '<w:tblW w:w="5000" w:type="pct"/>' +
    '<w:tblLayout w:type="fixed"/>' +
    "</w:tblPr>";

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

  const rowsXml: string[] = [];
  const headerCells = header.map((text, index) =>
    buildDocxCell({
      text,
      widthDxa: grid[index],
      bold: true,
      jc: "center",
      vAlign: "center",
    }),
  );
  rowsXml.push(`<w:tr>${headerCells.join("")}</w:tr>`);

  const sumHours = options.worklogRows.reduce(
    (acc, row) => acc + (Number(row.hours) || 0),
    0,
  );

  const blocks: Array<{ start: number; end: number }> = [];
  let index = 0;
  while (index < options.worklogRows.length) {
    const currentRow = options.worklogRows[index];
    if (!currentRow) break;
    const start = index;
    const key = currentRow.dateKey;
    let end = index;
    while (end + 1 < options.worklogRows.length) {
      const nextRow = options.worklogRows[end + 1];
      if (!nextRow || nextRow.dateKey !== key) break;
      end++;
    }
    blocks.push({ start, end });
    index = end + 1;
  }

  const rowMergeMode: Array<"restart" | "continue" | null> = new Array(
    options.worklogRows.length,
  ).fill(null);
  for (const block of blocks) {
    if (block.end > block.start) {
      rowMergeMode[block.start] = "restart";
      for (let r = block.start + 1; r <= block.end; r++)
        rowMergeMode[r] = "continue";
    }
  }

  for (let r = 0; r < options.worklogRows.length; r++) {
    const row = options.worklogRows[r];
    if (!row) continue;
    const dateText =
      row.dateValue instanceof Date
        ? formatGermanDate(row.dateValue)
        : String(row.dateValue);
    const hoursText = formatGermanNumber(row.hours, { decimals: 2 });

    const rowCells: string[] = [
      buildDocxCell({
        text: rowMergeMode[r] === "continue" ? "" : dateText,
        widthDxa: grid[0],
        jc: options.weekly ? "center" : "left",
        vAlign: "center",
        vMerge: rowMergeMode[r],
      }),
      buildDocxCell({
        text: row.user,
        widthDxa: grid[1],
        vAlign: "top",
      }),
      buildDocxCell({
        text: hoursText,
        widthDxa: grid[2],
        jc: "right",
        vAlign: "top",
      }),
    ];

    if (showAreaColumn) {
      const workAreaText = resolveWorkAreaValue(row, options.workAreasByKey);
      rowCells.push(
        buildDocxCell({
          text: workAreaText ?? "",
          widthDxa: grid[3],
          vAlign: "top",
        }),
      );
    }

    rowCells.push(
      buildDocxCell({
        text: row.description,
        widthDxa: grid[showAreaColumn ? 4 : 3],
        vAlign: "top",
      }),
    );

    rowsXml.push(`<w:tr>${rowCells.join("")}</w:tr>`);
  }

  rowsXml.push(buildSummaryRowXml({ grid, sumHours, showAreaColumn }));

  return `<w:tbl>${tblPr}${tblGrid}${rowsXml.join("")}</w:tbl><w:p><w:r><w:t></w:t></w:r></w:p>`;
}

function replaceParagraphContainingOnlyText(options: {
  xml: string;
  needleText: string;
  replacementXml: string;
}): string {
  const paragraphs = options.xml.match(/<w:p[\s\S]*?<\/w:p>/g);
  if (!paragraphs) throw new Error("DOCX document.xml has no paragraphs");

  for (const paragraph of paragraphs) {
    const text = extractParagraphText(paragraph).replace(/\s+/g, " ").trim();
    if (text === options.needleText) {
      return options.xml.replace(paragraph, options.replacementXml);
    }
  }

  throw new Error(
    `Could not find DOCX placeholder paragraph '${options.needleText}'`,
  );
}

function getReferencedWorkAreas(options: {
  worklogRows: WorklogRow[];
  workAreasByKey: Map<string, WorkAreaEntry>;
}): Map<string, WorkAreaEntry> {
  const referencedKeys = new Set<string>();
  for (const row of options.worklogRows) {
    if (row.key) referencedKeys.add(row.key);
    if (Array.isArray(row.keys)) {
      for (const key of row.keys) {
        if (key) referencedKeys.add(key);
      }
    }
  }

  const filtered = new Map<string, WorkAreaEntry>();
  for (const [key, entry] of options.workAreasByKey.entries()) {
    if (referencedKeys.has(key)) {
      filtered.set(key, entry);
    }
  }
  return filtered;
}

export function createDocx(options: {
  templateArrayBuffer: ArrayBuffer;
  worklogRows: WorklogRow[];
  workAreasByKey: Map<string, WorkAreaEntry> | null;
  weekly: boolean;
  includeLegend: boolean;
}): Uint8Array {
  const zip = new PizZip(options.templateArrayBuffer);
  const documentPath = "word/document.xml";
  const xmlFile = zip.file(documentPath);
  if (!xmlFile) {
    throw new Error("DOCX template has no word/document.xml");
  }

  const documentXml = xmlFile.asText();
  const tableXml = buildWorklogDocxTable({
    worklogRows: options.worklogRows,
    workAreasByKey: options.workAreasByKey,
    weekly: options.weekly,
  });

  let explanationXml = "";
  if (
    options.includeLegend &&
    options.workAreasByKey &&
    options.workAreasByKey.size > 0
  ) {
    const referencedWorkAreas = getReferencedWorkAreas({
      worklogRows: options.worklogRows,
      workAreasByKey: options.workAreasByKey,
    });
    if (referencedWorkAreas.size > 0) {
      explanationXml =
        '<w:p><w:pPr><w:spacing w:after="0"/></w:pPr><w:r><w:rPr><w:b/><w:bCs/><w:sz w:val="16"/><w:szCs w:val="16"/></w:rPr><w:t>Legende</w:t></w:r></w:p>' +
        buildWorkAreaExplanationTableXml(referencedWorkAreas) +
        "<w:p><w:r><w:t></w:t></w:r></w:p>";
    }
  }

  const updatedXml = replaceParagraphContainingOnlyText({
    xml: documentXml,
    needleText: "Tabelle",
    replacementXml: tableXml + explanationXml,
  });

  zip.file(documentPath, updatedXml);
  return zip.generate({ type: "uint8array" });
}
