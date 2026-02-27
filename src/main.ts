import { aggregateWeeklyRows } from "./transformer/aggregation";
import { createDocx } from "./transformer/docx";
import { createXlsx } from "./transformer/excel";
import {
  readWorkAreaMapFromCsv,
  readWorklogRowsFromCsv,
} from "./transformer/csv";
import type { WorklogRow } from "./transformer/types";
import "./style.css";

const RESULT_DOCX_FILE_NAME = "result.docx";
const RESULT_XLSX_FILE_NAME = "result.xlsx";
const DOCX_MIME_TYPE =
  "application/vnd.openxmlformats-officedocument.wordprocessingml.document";
const XLSX_MIME_TYPE =
  "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet";

const app = document.getElementById("app");
if (app) {
  app.innerHTML = `
    <h1>Timesheet Transformer</h1>
    <form id="uploadForm">
      <fieldset class="form-section">
        <legend>1. Worklog</legend>
        <label class="file-label">Worklog (csv):
          <input type="file" id="worklogInput" accept=".csv" required />
          <span id="worklogFileName" class="file-name"></span>
        </label>
        <label class="file-label">Work Areas (csv):
          <input type="file" id="areasInput" accept=".csv" />
          <span id="areasFileName" class="file-name"></span>
        </label>
        <label class="checkbox-label">
          <input type="checkbox" id="weeklyInput" />
          Weekly Aggregation
        </label>
        <label id="legendLabel" class="checkbox-label">
          <input type="checkbox" id="legendInput" />
          Include legend
        </label>
        <button type="button" id="generateExcelButton">Generate Excel</button>
        <button type="button" id="downloadExcelButton" disabled hidden>Download Excel</button>
      </fieldset>
      <fieldset class="form-section">
        <legend>2. Timesheet</legend>
        <label class="file-label">Template (docx):
          <input type="file" id="templateInput" accept=".docx" required />
          <span id="templateFileName" class="file-name"></span>
        </label>
        <button type="submit" id="generateButton">Generate DOCX</button>
        <button type="button" id="downloadButton" disabled hidden>Download DOCX</button>
      </fieldset>
    </form>
    <textarea id="logOutput" readonly></textarea>
  `;
}

const uploadForm = document.getElementById(
  "uploadForm",
) as HTMLFormElement | null;
const templateInput = document.getElementById(
  "templateInput",
) as HTMLInputElement | null;
const worklogInput = document.getElementById(
  "worklogInput",
) as HTMLInputElement | null;
const areasInput = document.getElementById(
  "areasInput",
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
const generateButton = document.getElementById(
  "generateButton",
) as HTMLButtonElement | null;
const generateExcelButton = document.getElementById(
  "generateExcelButton",
) as HTMLButtonElement | null;
const downloadButton = document.getElementById(
  "downloadButton",
) as HTMLButtonElement | null;
const downloadExcelButton = document.getElementById(
  "downloadExcelButton",
) as HTMLButtonElement | null;
const logOutput = document.getElementById(
  "logOutput",
) as HTMLTextAreaElement | null;
const worklogFileName = document.getElementById(
  "worklogFileName",
) as HTMLSpanElement | null;
const areasFileName = document.getElementById(
  "areasFileName",
) as HTMLSpanElement | null;
const templateFileName = document.getElementById(
  "templateFileName",
) as HTMLSpanElement | null;

type WorkAreasByKey = Map<string, { name: string; alias: string }>;

let downloadUrl: string | null = null;
let downloadFileName = RESULT_DOCX_FILE_NAME;
let downloadExcelUrl: string | null = null;
let downloadExcelFileName = RESULT_XLSX_FILE_NAME;
let runVersion = 0;
let isGeneratingDocx = false;
let isGeneratingExcel = false;

if (
  uploadForm &&
  templateInput &&
  worklogInput &&
  areasInput &&
  weeklyInput &&
  legendInput &&
  legendLabel &&
  generateButton &&
  generateExcelButton &&
  downloadButton &&
  downloadExcelButton &&
  logOutput &&
  worklogFileName &&
  areasFileName &&
  templateFileName
) {
  const setDocxResultAvailableState = (hasResult: boolean): void => {
    generateButton.hidden = hasResult;
    downloadButton.hidden = !hasResult;
    if (!hasResult) {
      downloadButton.disabled = true;
    }
  };

  const setExcelResultAvailableState = (hasResult: boolean): void => {
    generateExcelButton.hidden = hasResult;
    downloadExcelButton.hidden = !hasResult;
    if (!hasResult) {
      downloadExcelButton.disabled = true;
    }
  };

  const appendLog = (line: string): void => {
    logOutput.value += `${line}\n`;
    logOutput.scrollTop = logOutput.scrollHeight;
  };

  const getResultFileNameFromWorklog = (worklogName: string): string => {
    const trimmedName = String(worklogName ?? "").trim();
    if (!trimmedName) return RESULT_DOCX_FILE_NAME;

    const lastDotIndex = trimmedName.lastIndexOf(".");
    const baseName =
      lastDotIndex > 0 ? trimmedName.slice(0, lastDotIndex) : trimmedName;
    const safeBaseName = baseName.trim() || "result";
    return `${safeBaseName}.docx`;
  };

  const getExcelFileNameFromWorklog = (worklogName: string): string => {
    const trimmedName = String(worklogName ?? "").trim();
    if (!trimmedName) return RESULT_XLSX_FILE_NAME;

    const lastDotIndex = trimmedName.lastIndexOf(".");
    const baseName =
      lastDotIndex > 0 ? trimmedName.slice(0, lastDotIndex) : trimmedName;
    const safeBaseName = baseName.trim() || "result";
    return `${safeBaseName}.xlsx`;
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
      );
    }

    if (warningCount > 0) {
      appendLog(
        `Warnings found: ${warningCount} worklog row(s) without matching work area keys.`,
      );
    }
  };

  const resetPreviousRunState = ({
    clearLog = true,
  }: { clearLog?: boolean } = {}): void => {
    runVersion += 1;
    if (downloadUrl) {
      URL.revokeObjectURL(downloadUrl);
      downloadUrl = null;
    }
    if (downloadExcelUrl) {
      URL.revokeObjectURL(downloadExcelUrl);
      downloadExcelUrl = null;
    }
    downloadFileName = RESULT_DOCX_FILE_NAME;
    downloadExcelFileName = RESULT_XLSX_FILE_NAME;
    setDocxResultAvailableState(false);
    setExcelResultAvailableState(false);
    if (clearLog) {
      logOutput.value = "";
    }
  };

  legendInput.disabled = true;

  const syncUI = (): void => {
    const hasAreasFile = Boolean(areasInput.files?.[0]);
    const hasWorklog = Boolean(worklogInput.files?.[0]);
    const hasTemplate = Boolean(templateInput.files?.[0]);

    worklogFileName.textContent = worklogInput.files?.[0]?.name ?? "";
    areasFileName.textContent = areasInput.files?.[0]?.name ?? "";
    templateFileName.textContent = templateInput.files?.[0]?.name ?? "";

    legendInput.disabled = !hasAreasFile;
    legendLabel.classList.toggle("is-disabled", !hasAreasFile);
    if (!hasAreasFile) {
      legendInput.checked = false;
    }

    generateButton.disabled = isGeneratingDocx || !(hasWorklog && hasTemplate);
    generateExcelButton.disabled = isGeneratingExcel || !hasWorklog;
  };

  const registerChangeHandler = (element: HTMLElement): void => {
    element.addEventListener("change", () => {
      resetPreviousRunState();
      syncUI();
    });
  };

  registerChangeHandler(templateInput);
  registerChangeHandler(worklogInput);
  registerChangeHandler(areasInput);
  registerChangeHandler(weeklyInput);
  registerChangeHandler(legendInput);
  syncUI();
  setDocxResultAvailableState(false);
  setExcelResultAvailableState(false);

  downloadButton.addEventListener("click", () => {
    if (!downloadUrl) return;
    const anchor = document.createElement("a");
    anchor.href = downloadUrl;
    anchor.download = downloadFileName;
    anchor.click();
  });

  downloadExcelButton.addEventListener("click", () => {
    if (!downloadExcelUrl) return;
    const anchor = document.createElement("a");
    anchor.href = downloadExcelUrl;
    anchor.download = downloadExcelFileName;
    anchor.click();
  });

  const loadWorklogData = async () => {
    const worklogFile = worklogInput.files?.[0] ?? null;
    if (!worklogFile) {
      appendLog("Please choose a worklog CSV file.");
      return null;
    }

    appendLog("Reading files...");
    const worklogCsv = await worklogFile.text();

    const dailyRows = readWorklogRowsFromCsv(worklogCsv, appendLog);
    const weekly = weeklyInput.checked;
    const worklogRows = weekly ? aggregateWeeklyRows(dailyRows) : dailyRows;

    if (worklogRows.length === 0) {
      throw new Error(
        "No usable worklog rows found in CSV (after filtering summary rows).",
      );
    }

    appendLog(`Usable worklog rows: ${worklogRows.length}`);

    let workAreasByKey: WorkAreasByKey | null = null;
    const areasFile = areasInput.files?.[0] ?? null;
    if (areasFile) {
      appendLog("Reading optional work areas CSV...");
      const areasCsv = await areasFile.text();
      workAreasByKey = readWorkAreaMapFromCsv(areasCsv);
      appendLog(`Loaded work areas: ${workAreasByKey.size}`);
      warnOnMissingWorkAreaMappings(worklogRows, workAreasByKey);
    }

    return {
      worklogFile,
      worklogRows,
      workAreasByKey,
      weekly,
    };
  };

  generateExcelButton.addEventListener("click", async () => {
    resetPreviousRunState();
    const startedRunVersion = runVersion;
    isGeneratingExcel = true;
    syncUI();

    try {
      const loaded = await loadWorklogData();
      if (!loaded) return;
      downloadExcelFileName = getExcelFileNameFromWorklog(
        loaded.worklogFile.name,
      );

      appendLog("Generating XLSX...");
      const resultBytes = createXlsx({
        worklogRows: loaded.worklogRows,
        workAreasByKey: loaded.workAreasByKey,
        weekly: loaded.weekly,
        includeLegend: Boolean(legendInput.checked && loaded.workAreasByKey),
      });

      if (startedRunVersion !== runVersion) {
        return;
      }

      const resultArrayBuffer = resultBytes.buffer.slice(
        resultBytes.byteOffset,
        resultBytes.byteOffset + resultBytes.byteLength,
      );
      const blobBytes = new Uint8Array(resultBytes.byteLength);
      blobBytes.set(new Uint8Array(resultArrayBuffer));

      const blob = new Blob([blobBytes], { type: XLSX_MIME_TYPE });
      downloadExcelUrl = URL.createObjectURL(blob);
      setExcelResultAvailableState(true);
      downloadExcelButton.disabled = false;
      appendLog("XLSX created successfully. Download Excel is now enabled.");
    } catch (err) {
      const message = err instanceof Error ? err.message : String(err);
      appendLog(`Error: ${message}`);
    } finally {
      isGeneratingExcel = false;
      syncUI();
    }
  });

  uploadForm.addEventListener("submit", async (event) => {
    event.preventDefault();
    resetPreviousRunState();
    const startedRunVersion = runVersion;

    const templateFile = templateInput.files?.[0] ?? null;
    if (!templateFile) {
      appendLog("Please choose a template DOCX file.");
      return;
    }

    const worklogFile = worklogInput.files?.[0] ?? null;
    if (!worklogFile) {
      appendLog("Please choose a worklog CSV file.");
      return;
    }
    downloadFileName = getResultFileNameFromWorklog(worklogFile.name);

    isGeneratingDocx = true;
    syncUI();

    try {
      const templateArrayBuffer = await templateFile.arrayBuffer();
      const loaded = await loadWorklogData();
      if (!loaded) return;

      appendLog("Generating DOCX...");
      const resultBytes = createDocx({
        templateArrayBuffer,
        worklogRows: loaded.worklogRows,
        workAreasByKey: loaded.workAreasByKey,
        weekly: loaded.weekly,
        includeLegend: Boolean(legendInput.checked && loaded.workAreasByKey),
      });

      if (startedRunVersion !== runVersion) {
        return;
      }

      const resultArrayBuffer = resultBytes.buffer.slice(
        resultBytes.byteOffset,
        resultBytes.byteOffset + resultBytes.byteLength,
      );
      const blobBytes = new Uint8Array(resultBytes.byteLength);
      blobBytes.set(new Uint8Array(resultArrayBuffer));

      const blob = new Blob([blobBytes], { type: DOCX_MIME_TYPE });
      downloadUrl = URL.createObjectURL(blob);
      setDocxResultAvailableState(true);
      downloadButton.disabled = false;
      appendLog("DOCX created successfully. Download result is now enabled.");
    } catch (err) {
      const message = err instanceof Error ? err.message : String(err);
      appendLog(`Error: ${message}`);
    } finally {
      isGeneratingDocx = false;
      syncUI();
    }
  });
}
