// @vitest-environment jsdom
import { beforeEach, describe, expect, it, vi } from "vitest";

vi.mock("../src/transformer/csv", () => {
  const row = {
    dateValue: new Date("2026-02-03T12:00:00.000Z"),
    dateKey: "2026-02-03",
    dateSort: new Date("2026-02-03T12:00:00.000Z").valueOf(),
    user: "Test User",
    hours: 1,
    description: "task",
    key: "TEST-1",
  };

  return {
    readWorklogRowsFromCsv: vi.fn(() => [row]),
    readWorkAreaMapFromCsv: vi.fn(() => new Map()),
  };
});

vi.mock("../src/transformer/aggregation", () => ({
  aggregateWeeklyRows: vi.fn((rows) => rows),
}));

vi.mock("../src/transformer/docx", () => ({
  createDocx: vi.fn(() => new Uint8Array([1, 2, 3])),
}));

vi.mock("../src/transformer/excel", () => ({
  createXlsx: vi.fn(() => new Uint8Array([1, 2, 3])),
}));

function setInputFiles(input: HTMLInputElement, files: File[]): void {
  Object.defineProperty(input, "files", {
    value: files,
    configurable: true,
  });
}

function makeTemplateFileMock(): File {
  return {
    name: "Template.docx",
    arrayBuffer: async () => new ArrayBuffer(16),
  } as unknown as File;
}

function makeWorklogFileMock(): File {
  return {
    name: "worklog.csv",
    text: async () => "User,Worklog,Key,Logged,Date",
  } as unknown as File;
}

function makeAreasFileMock(): File {
  return {
    name: "work_areas.csv",
    text: async () => "Key,Name,Alias",
  } as unknown as File;
}

async function waitUntil(predicate: () => boolean): Promise<void> {
  const timeoutMs = 250;
  const startedAt = Date.now();
  while (!predicate()) {
    if (Date.now() - startedAt > timeoutMs) {
      throw new Error("Timed out waiting for UI state update");
    }
    await new Promise((resolve) => setTimeout(resolve, 5));
  }
}

describe("main process and export workflow", () => {
  beforeEach(() => {
    vi.resetModules();
    document.body.innerHTML = '<div id="app"></div>';
    Object.defineProperty(globalThis.URL, "createObjectURL", {
      value: vi.fn(() => "blob:test"),
      configurable: true,
    });
    Object.defineProperty(globalThis.URL, "revokeObjectURL", {
      value: vi.fn(),
      configurable: true,
    });
  });

  it("enables Process when worklog is selected", async () => {
    await import("../src/main");

    const worklogInput = document.getElementById(
      "worklogInput",
    ) as HTMLInputElement;
    const processButton = document.getElementById(
      "processButton",
    ) as HTMLButtonElement;

    expect(processButton.disabled).toBe(true);
    setInputFiles(worklogInput, [makeWorklogFileMock()]);
    worklogInput.dispatchEvent(new Event("change", { bubbles: true }));
    expect(processButton.disabled).toBe(false);
  });

  it("shows log and enables Download Excel after successful process", async () => {
    await import("../src/main");

    const worklogInput = document.getElementById(
      "worklogInput",
    ) as HTMLInputElement;
    const processButton = document.getElementById(
      "processButton",
    ) as HTMLButtonElement;
    const downloadExcelButton = document.getElementById(
      "downloadExcelButton",
    ) as HTMLButtonElement;
    const exportSection = document.getElementById(
      "exportSection",
    ) as HTMLFieldSetElement;
    const logOutput = document.getElementById("logOutput") as HTMLDivElement;

    expect(exportSection.hidden).toBe(true);

    setInputFiles(worklogInput, [makeWorklogFileMock()]);
    worklogInput.dispatchEvent(new Event("change", { bubbles: true }));
    processButton.click();

    await waitUntil(() => !downloadExcelButton.disabled);
    expect(exportSection.hidden).toBe(false);
    expect(logOutput.hidden).toBe(false);
    expect(logOutput.textContent).toContain("Processing finished");
  });

  it("enables Download DOCX after processing and exports after template picker change", async () => {
    await import("../src/main");

    const worklogInput = document.getElementById(
      "worklogInput",
    ) as HTMLInputElement;
    const templatePickerInput = document.getElementById(
      "templatePickerInput",
    ) as HTMLInputElement;
    const processButton = document.getElementById(
      "processButton",
    ) as HTMLButtonElement;
    const downloadDocxButton = document.getElementById(
      "downloadDocxButton",
    ) as HTMLButtonElement;
    const logOutput = document.getElementById("logOutput") as HTMLDivElement;

    expect(downloadDocxButton.disabled).toBe(true);

    setInputFiles(worklogInput, [makeWorklogFileMock()]);
    worklogInput.dispatchEvent(new Event("change", { bubbles: true }));
    processButton.click();
    await waitUntil(() => !downloadDocxButton.disabled);
    expect(downloadDocxButton.disabled).toBe(false);

    downloadDocxButton.click();
    setInputFiles(templatePickerInput, [makeTemplateFileMock()]);
    templatePickerInput.dispatchEvent(new Event("change", { bubbles: true }));
    await waitUntil(() =>
      Boolean(logOutput.textContent?.includes("DOCX downloaded.")),
    );
  });

  it("legend checkbox is enabled only when work areas file exists", async () => {
    await import("../src/main");

    const areasInput = document.getElementById(
      "areasInput",
    ) as HTMLInputElement;
    const legendInput = document.getElementById(
      "legendInput",
    ) as HTMLInputElement;

    expect(legendInput.disabled).toBe(true);

    setInputFiles(areasInput, [makeAreasFileMock()]);
    areasInput.dispatchEvent(new Event("change", { bubbles: true }));
    expect(legendInput.disabled).toBe(false);

    legendInput.checked = true;
    setInputFiles(areasInput, []);
    areasInput.dispatchEvent(new Event("change", { bubbles: true }));
    expect(legendInput.disabled).toBe(true);
    expect(legendInput.checked).toBe(false);
  });

  it("shows selected file names in custom labels", async () => {
    await import("../src/main");

    const worklogInput = document.getElementById(
      "worklogInput",
    ) as HTMLInputElement;
    const areasInput = document.getElementById(
      "areasInput",
    ) as HTMLInputElement;

    const worklogName = document.getElementById(
      "worklogFileName",
    ) as HTMLSpanElement;
    const areasName = document.getElementById(
      "areasFileName",
    ) as HTMLSpanElement;

    setInputFiles(worklogInput, [makeWorklogFileMock()]);
    worklogInput.dispatchEvent(new Event("change", { bubbles: true }));
    setInputFiles(areasInput, [makeAreasFileMock()]);
    areasInput.dispatchEvent(new Event("change", { bubbles: true }));
    expect(worklogName.textContent).toBe("worklog.csv");
    expect(areasName.textContent).toBe("work_areas.csv");
  });

  it("renders missing work-area references as warning log lines", async () => {
    await import("../src/main");

    const worklogInput = document.getElementById(
      "worklogInput",
    ) as HTMLInputElement;
    const areasInput = document.getElementById(
      "areasInput",
    ) as HTMLInputElement;
    const processButton = document.getElementById(
      "processButton",
    ) as HTMLButtonElement;
    const logOutput = document.getElementById("logOutput") as HTMLDivElement;

    setInputFiles(worklogInput, [makeWorklogFileMock()]);
    worklogInput.dispatchEvent(new Event("change", { bubbles: true }));
    setInputFiles(areasInput, [makeAreasFileMock()]);
    areasInput.dispatchEvent(new Event("change", { bubbles: true }));
    processButton.click();

    await waitUntil(() =>
      Boolean(logOutput.querySelector(".log-line-warn")?.textContent),
    );

    const warnLine = logOutput.querySelector(".log-line-warn");
    expect(warnLine?.textContent).toContain("no matching work area key");
  });
});
