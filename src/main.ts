import { aggregateWeeklyRows } from "./transformer/aggregation";
import {
  readWorkAreaMapFromCsv,
  readWorklogRowsFromCsv,
} from "./transformer/csv";
import { createDocx } from "./transformer/docx";
import { createXlsx } from "./transformer/excel";
import type { WorklogRow } from "./transformer/types";
import "./style.css";

const RESULT_DOCX_FILE_NAME = "result.docx";
const RESULT_XLSX_FILE_NAME = "result.xlsx";
const DOCX_MIME_TYPE =
  "application/vnd.openxmlformats-officedocument.wordprocessingml.document";
const XLSX_MIME_TYPE =
  "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet";

type WorkAreasByKey = Map<string, { name: string; alias: string }>;
type ProcessedData = {
  worklogFileName: string;
  dailyRows: WorklogRow[];
  workAreasByKey: WorkAreasByKey | null;
};

const app = document.getElementById("app");
if (app) {
  app.innerHTML = `
    <h1>Timesheet Transformer</h1>
    <form id="uploadForm">
      <fieldset class="form-section">
        <legend>Input</legend>
        <label class="file-label">Worklog (csv):
          <input type="file" id="worklogInput" accept=".csv" required />
          <span id="worklogFileName" class="file-name"></span>
        </label>
        <label class="file-label">Work Areas (csv):
          <input type="file" id="areasInput" accept=".csv" />
          <span id="areasFileName" class="file-name"></span>
        </label>
        <button type="button" id="processButton">Process</button>
      </fieldset>

      <div id="logOutput" hidden></div>

      <fieldset id="exportSection" class="form-section" hidden>
        <legend>Export</legend>
        <label class="checkbox-label">
          <input type="checkbox" id="weeklyInput" />
          Weekly Aggregation
        </label>
        <label id="legendLabel" class="checkbox-label">
          <input type="checkbox" id="legendInput" />
          Include legend
        </label>
        <button type="button" id="downloadExcelButton" disabled>Download Excel</button>
        <button type="button" id="downloadDocxButton" disabled>Download DOCX</button>
        <input type="file" id="templatePickerInput" accept=".docx" hidden />
      </fieldset>
    </form>
  `;
}

const worklogInput = document.getElementById(
  "worklogInput",
) as HTMLInputElement | null;
const areasInput = document.getElementById(
  "areasInput",
) as HTMLInputElement | null;
const templatePickerInput = document.getElementById(
  "templatePickerInput",
) as HTMLInputElement | null;
const weeklyInput = document.getElementById(
  "weeklyInput",
) as HTMLInputElement | null;
const legendInput = document.getElementById(
  "legendInput",
) as HTMLInputElement | null;
const legendLabel = document.getElementById(
  "legendLabel",
) as HTMLLabelElement | null;
const processButton = document.getElementById(
  "processButton",
) as HTMLButtonElement | null;
const downloadExcelButton = document.getElementById(
  "downloadExcelButton",
) as HTMLButtonElement | null;
const downloadDocxButton = document.getElementById(
  "downloadDocxButton",
) as HTMLButtonElement | null;
const exportSection = document.getElementById(
  "exportSection",
) as HTMLFieldSetElement | null;
const logOutput = document.getElementById("logOutput") as HTMLDivElement | null;
const worklogFileName = document.getElementById(
  "worklogFileName",
) as HTMLSpanElement | null;
const areasFileName = document.getElementById(
  "areasFileName",
) as HTMLSpanElement | null;

let processedData: ProcessedData | null = null;
let isProcessing = false;
let isExportingDocx = false;

if (
  worklogInput &&
  areasInput &&
  templatePickerInput &&
  weeklyInput &&
  legendInput &&
  legendLabel &&
  processButton &&
  downloadExcelButton &&
  downloadDocxButton &&
  exportSection &&
  logOutput &&
  worklogFileName &&
  areasFileName
) {
  const appendLog = (
    line: string,
    level: "info" | "warn" | "error" = "info",
  ): void => {
    logOutput.hidden = false;
    const lineNode = document.createElement("div");
    const resolvedLevel =
      level === "info" && /^error:/i.test(line) ? "error" : level;
    lineNode.className =
      resolvedLevel === "warn"
        ? "log-line log-line-warn"
        : resolvedLevel === "error"
          ? "log-line log-line-error"
          : "log-line";
    lineNode.textContent = line;
    logOutput.appendChild(lineNode);
    logOutput.scrollTop = logOutput.scrollHeight;
  };

  const getResultFileName = (
    inputName: string,
    ext: "docx" | "xlsx",
  ): string => {
    const trimmedName = String(inputName ?? "").trim();
    if (!trimmedName) {
      return ext === "docx" ? RESULT_DOCX_FILE_NAME : RESULT_XLSX_FILE_NAME;
    }
    const lastDotIndex = trimmedName.lastIndexOf(".");
    const baseName =
      lastDotIndex > 0 ? trimmedName.slice(0, lastDotIndex) : trimmedName;
    const safeBaseName = baseName.trim() || "result";
    return `${safeBaseName}.${ext}`;
  };

  const warnOnMissingWorkAreaMappings = (
    rows: WorklogRow[],
    areasByKey: WorkAreasByKey,
  ): void => {
    let warningCount = 0;
    for (let index = 0; index < rows.length; index++) {
      const row = rows[index];
      if (!row) continue;
      const rawKeys = [
        row.key,
        ...(Array.isArray(row.keys) ? row.keys : []),
      ].filter((value): value is string => Boolean(value));
      const keys = Array.from(new Set(rawKeys));
      const hasMatchingKey = keys.some((key) => areasByKey.has(key));
      if (hasMatchingKey) continue;
      warningCount += 1;
      appendLog(
        `Warning: no matching work area key for worklog row ${index + 1} (user: ${row.user}, date: ${row.dateKey}, keys: ${keys.length > 0 ? keys.join(", ") : "none"}).`,
        "warn",
      );
    }

    if (warningCount > 0) {
      appendLog(
        `Warnings found: ${warningCount} worklog row(s) without matching work area keys.`,
        "warn",
      );
    }
  };

  const getRowsForExport = (): WorklogRow[] => {
    if (!processedData) return [];
    return weeklyInput.checked
      ? aggregateWeeklyRows(processedData.dailyRows)
      : processedData.dailyRows;
  };

  const triggerDownload = (
    bytes: Uint8Array,
    fileName: string,
    mimeType: string,
  ): void => {
    const resultArrayBuffer = bytes.buffer.slice(
      bytes.byteOffset,
      bytes.byteOffset + bytes.byteLength,
    );
    const blobBytes = new Uint8Array(bytes.byteLength);
    blobBytes.set(new Uint8Array(resultArrayBuffer));
    const blob = new Blob([blobBytes], { type: mimeType });
    const url = URL.createObjectURL(blob);
    const anchor = document.createElement("a");
    anchor.href = url;
    anchor.download = fileName;
    anchor.click();
    URL.revokeObjectURL(url);
  };

  const invalidateProcessedData = ({
    clearLog,
  }: {
    clearLog: boolean;
  }): void => {
    processedData = null;
    if (clearLog) {
      logOutput.textContent = "";
      logOutput.hidden = true;
    }
  };

  const syncUI = (): void => {
    const hasWorklog = Boolean(worklogInput.files?.[0]);
    const hasAreas = Boolean(areasInput.files?.[0]);
    const hasProcessed = Boolean(processedData);

    worklogFileName.textContent = worklogInput.files?.[0]?.name ?? "";
    areasFileName.textContent = areasInput.files?.[0]?.name ?? "";

    processButton.disabled = isProcessing || !hasWorklog;
    exportSection.hidden = !hasProcessed;
    exportSection.style.display = hasProcessed ? "grid" : "none";

    legendInput.disabled = !hasAreas;
    legendLabel.classList.toggle("is-disabled", !hasAreas);
    if (!hasAreas) {
      legendInput.checked = false;
    }

    downloadExcelButton.disabled = !hasProcessed;
    downloadDocxButton.disabled = isExportingDocx || !hasProcessed;
  };

  worklogInput.addEventListener("change", () => {
    invalidateProcessedData({ clearLog: true });
    syncUI();
  });
  areasInput.addEventListener("change", () => {
    invalidateProcessedData({ clearLog: true });
    syncUI();
  });
  weeklyInput.addEventListener("change", syncUI);
  legendInput.addEventListener("change", syncUI);

  processButton.addEventListener("click", async () => {
    const worklogFile = worklogInput.files?.[0] ?? null;
    if (!worklogFile) {
      appendLog("Please choose a worklog CSV file.");
      return;
    }

    isProcessing = true;
    invalidateProcessedData({ clearLog: true });
    syncUI();

    try {
      appendLog("Reading files...");
      const worklogCsv = await worklogFile.text();
      const dailyRows = readWorklogRowsFromCsv(worklogCsv, appendLog);
      if (dailyRows.length === 0) {
        throw new Error(
          "No usable worklog rows found in CSV (after filtering summary rows).",
        );
      }
      appendLog(`Usable worklog rows: ${dailyRows.length}`);

      let workAreasByKey: WorkAreasByKey | null = null;
      const areasFile = areasInput.files?.[0] ?? null;
      if (areasFile) {
        appendLog("Reading optional work areas CSV...");
        const areasCsv = await areasFile.text();
        workAreasByKey = readWorkAreaMapFromCsv(areasCsv);
        appendLog(`Loaded work areas: ${workAreasByKey.size}`);
        warnOnMissingWorkAreaMappings(dailyRows, workAreasByKey);
      }

      processedData = {
        worklogFileName: worklogFile.name,
        dailyRows,
        workAreasByKey,
      };
      appendLog("Processing finished. You can now export Excel or DOCX.");
    } catch (err) {
      const message = err instanceof Error ? err.message : String(err);
      appendLog(`Error: ${message}`, "error");
    } finally {
      isProcessing = false;
      syncUI();
    }
  });

  downloadExcelButton.addEventListener("click", () => {
    if (!processedData) return;

    const rowsForExport = getRowsForExport();
    const resultBytes = createXlsx({
      worklogRows: rowsForExport,
      workAreasByKey: processedData.workAreasByKey,
      weekly: weeklyInput.checked,
      includeLegend: Boolean(
        legendInput.checked && processedData.workAreasByKey,
      ),
    });

    triggerDownload(
      resultBytes,
      getResultFileName(processedData.worklogFileName, "xlsx"),
      XLSX_MIME_TYPE,
    );
    appendLog("XLSX downloaded.");
  });

  downloadDocxButton.addEventListener("click", () => {
    if (!processedData) return;
    templatePickerInput.value = "";
    templatePickerInput.click();
  });

  templatePickerInput.addEventListener("change", async () => {
    if (!processedData) return;
    const templateFile = templatePickerInput.files?.[0] ?? null;
    if (!templateFile) return;

    isExportingDocx = true;
    syncUI();

    try {
      const templateArrayBuffer = await templateFile.arrayBuffer();
      const rowsForExport = getRowsForExport();
      const resultBytes = createDocx({
        templateArrayBuffer,
        worklogRows: rowsForExport,
        workAreasByKey: processedData.workAreasByKey,
        weekly: weeklyInput.checked,
        includeLegend: Boolean(
          legendInput.checked && processedData.workAreasByKey,
        ),
      });

      triggerDownload(
        resultBytes,
        getResultFileName(processedData.worklogFileName, "docx"),
        DOCX_MIME_TYPE,
      );
      appendLog("DOCX downloaded.");
    } catch (err) {
      const message = err instanceof Error ? err.message : String(err);
      appendLog(`Error: ${message}`, "error");
    } finally {
      isExportingDocx = false;
      syncUI();
    }
  });

  syncUI();
}
