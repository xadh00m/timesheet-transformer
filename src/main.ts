import { aggregateWeeklyRows } from "./transformer/aggregation";
import { createDocx } from "./transformer/docx";
import {
  readWorkAreaMapFromCsv,
  readWorklogRowsFromCsv,
} from "./transformer/csv";
import type { WorklogRow } from "./transformer/types";
import "./style.css";

const RESULT_DOCX_FILE_NAME = "result.docx";
const DOCX_MIME_TYPE =
  "application/vnd.openxmlformats-officedocument.wordprocessingml.document";

const app = document.getElementById("app");
if (app) {
  app.innerHTML = `
    <h1>Timesheet Transformer</h1>
    <form id="uploadForm">
      <fieldset class="form-section">
        <legend>1. Worklog</legend>
        <label class="file-label">Worklog (csv):
          <input type="file" id="worklogInput" accept=".csv" required />
        </label>
        <label class="file-label">Work Areas (csv):
          <input type="file" id="areasInput" accept=".csv" />
        </label>
        <label class="checkbox-label">
          <input type="checkbox" id="weeklyInput" />
          Weekly Aggregation
        </label>
      </fieldset>
      <fieldset class="form-section">
        <legend>2. Timesheet</legend>
        <label class="file-label">Template (docx):
          <input type="file" id="templateInput" accept=".docx" required />
        </label>
        <label id="legendLabel" class="checkbox-label">
          <input type="checkbox" id="legendInput" />
          Include legend
        </label>
        <button type="submit" id="generateButton">Generate</button>
        <button type="button" id="downloadButton" disabled hidden>Download</button>
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
const downloadButton = document.getElementById(
  "downloadButton",
) as HTMLButtonElement | null;
const logOutput = document.getElementById(
  "logOutput",
) as HTMLTextAreaElement | null;

type WorkAreasByKey = Map<string, { name: string; alias: string }>;

let downloadUrl: string | null = null;
let downloadFileName = RESULT_DOCX_FILE_NAME;
let runVersion = 0;

if (
  uploadForm &&
  templateInput &&
  worklogInput &&
  areasInput &&
  weeklyInput &&
  legendInput &&
  legendLabel &&
  generateButton &&
  downloadButton &&
  logOutput
) {
  const setResultAvailableState = (hasResult: boolean): void => {
    generateButton.hidden = hasResult;
    downloadButton.hidden = !hasResult;
    if (!hasResult) {
      downloadButton.disabled = true;
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
    downloadFileName = RESULT_DOCX_FILE_NAME;
    setResultAvailableState(false);
    if (clearLog) {
      logOutput.value = "";
    }
  };

  legendInput.disabled = true;

  const syncUI = (): void => {
    const hasAreasFile = Boolean(areasInput.files?.[0]);
    const hasWorklog = Boolean(worklogInput.files?.[0]);
    const hasTemplate = Boolean(templateInput.files?.[0]);

    legendInput.disabled = !hasAreasFile;
    legendLabel.classList.toggle("is-disabled", !hasAreasFile);
    if (!hasAreasFile) {
      legendInput.checked = false;
    }

    generateButton.disabled = !(hasWorklog && hasTemplate);
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
  setResultAvailableState(false);

  downloadButton.addEventListener("click", () => {
    if (!downloadUrl) return;
    const anchor = document.createElement("a");
    anchor.href = downloadUrl;
    anchor.download = downloadFileName;
    anchor.click();
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

    generateButton.disabled = true;

    try {
      appendLog("Reading files...");
      const templateArrayBuffer = await templateFile.arrayBuffer();
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

      appendLog("Generating DOCX...");
      const resultBytes = createDocx({
        templateArrayBuffer,
        worklogRows,
        workAreasByKey,
        weekly,
        includeLegend: Boolean(legendInput.checked && workAreasByKey),
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
      setResultAvailableState(true);
      downloadButton.disabled = false;
      appendLog("DOCX created successfully. Download result is now enabled.");
    } catch (err) {
      const message = err instanceof Error ? err.message : String(err);
      appendLog(`Error: ${message}`);
    } finally {
      syncUI();
    }
  });
}
