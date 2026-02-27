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

describe("main UI state handling", () => {
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

  it("swaps generate/download buttons after successful generation", async () => {
    await import("../src/main");

    const templateInput = document.getElementById(
      "templateInput",
    ) as HTMLInputElement;
    const worklogInput = document.getElementById(
      "worklogInput",
    ) as HTMLInputElement;
    const uploadForm = document.getElementById("uploadForm") as HTMLFormElement;
    const generateButton = document.getElementById(
      "generateButton",
    ) as HTMLButtonElement;
    const downloadButton = document.getElementById(
      "downloadButton",
    ) as HTMLButtonElement;
    const generateExcelButton = document.getElementById(
      "generateExcelButton",
    ) as HTMLButtonElement;

    setInputFiles(templateInput, [makeTemplateFileMock()]);
    setInputFiles(worklogInput, [makeWorklogFileMock()]);

    uploadForm.dispatchEvent(
      new Event("submit", { bubbles: true, cancelable: true }),
    );
    await waitUntil(() => generateButton.hidden && !downloadButton.hidden);

    expect(generateButton.hidden).toBe(true);
    expect(downloadButton.hidden).toBe(false);
    expect(downloadButton.disabled).toBe(false);
    expect(generateExcelButton.disabled).toBe(false);
  });

  it("swaps excel generate/download buttons after successful excel generation", async () => {
    await import("../src/main");

    const worklogInput = document.getElementById(
      "worklogInput",
    ) as HTMLInputElement;
    const generateExcelButton = document.getElementById(
      "generateExcelButton",
    ) as HTMLButtonElement;
    const downloadExcelButton = document.getElementById(
      "downloadExcelButton",
    ) as HTMLButtonElement;

    setInputFiles(worklogInput, [makeWorklogFileMock()]);
    worklogInput.dispatchEvent(new Event("change", { bubbles: true }));

    expect(generateExcelButton.disabled).toBe(false);
    generateExcelButton.click();

    await waitUntil(() =>
      Boolean(generateExcelButton.hidden && !downloadExcelButton.hidden),
    );
    expect(downloadExcelButton.disabled).toBe(false);
  });

  it("resets previous result state on input/checkbox changes", async () => {
    await import("../src/main");

    const templateInput = document.getElementById(
      "templateInput",
    ) as HTMLInputElement;
    const worklogInput = document.getElementById(
      "worklogInput",
    ) as HTMLInputElement;
    const weeklyInput = document.getElementById(
      "weeklyInput",
    ) as HTMLInputElement;
    const uploadForm = document.getElementById("uploadForm") as HTMLFormElement;
    const generateButton = document.getElementById(
      "generateButton",
    ) as HTMLButtonElement;
    const downloadButton = document.getElementById(
      "downloadButton",
    ) as HTMLButtonElement;

    setInputFiles(templateInput, [makeTemplateFileMock()]);
    setInputFiles(worklogInput, [makeWorklogFileMock()]);

    uploadForm.dispatchEvent(
      new Event("submit", { bubbles: true, cancelable: true }),
    );
    await waitUntil(() => generateButton.hidden && !downloadButton.hidden);

    expect(generateButton.hidden).toBe(true);
    expect(downloadButton.hidden).toBe(false);

    weeklyInput.checked = true;
    weeklyInput.dispatchEvent(new Event("change", { bubbles: true }));

    expect(generateButton.hidden).toBe(false);
    expect(downloadButton.hidden).toBe(true);
    expect(downloadButton.disabled).toBe(true);
  });

  it("enables legend only when work areas file is selected", async () => {
    await import("../src/main");

    const areasInput = document.getElementById(
      "areasInput",
    ) as HTMLInputElement;
    const legendInput = document.getElementById(
      "legendInput",
    ) as HTMLInputElement;

    expect(legendInput.disabled).toBe(true);

    setInputFiles(areasInput, [makeWorklogFileMock()]);
    areasInput.dispatchEvent(new Event("change", { bubbles: true }));
    expect(legendInput.disabled).toBe(false);

    legendInput.checked = true;
    setInputFiles(areasInput, []);
    areasInput.dispatchEvent(new Event("change", { bubbles: true }));
    expect(legendInput.disabled).toBe(true);
    expect(legendInput.checked).toBe(false);
  });

  it("shows selected file names in custom filename labels", async () => {
    await import("../src/main");

    const templateInput = document.getElementById(
      "templateInput",
    ) as HTMLInputElement;
    const worklogInput = document.getElementById(
      "worklogInput",
    ) as HTMLInputElement;
    const worklogFileName = document.getElementById(
      "worklogFileName",
    ) as HTMLSpanElement;
    const templateFileName = document.getElementById(
      "templateFileName",
    ) as HTMLSpanElement;

    setInputFiles(worklogInput, [makeWorklogFileMock()]);
    worklogInput.dispatchEvent(new Event("change", { bubbles: true }));
    setInputFiles(templateInput, [makeTemplateFileMock()]);
    templateInput.dispatchEvent(new Event("change", { bubbles: true }));

    expect(worklogFileName.textContent).toBe("worklog.csv");
    expect(templateFileName.textContent).toBe("Template.docx");
  });

  it("keeps DOCX disabled until both worklog and template are selected", async () => {
    await import("../src/main");

    const templateInput = document.getElementById(
      "templateInput",
    ) as HTMLInputElement;
    const worklogInput = document.getElementById(
      "worklogInput",
    ) as HTMLInputElement;
    const generateButton = document.getElementById(
      "generateButton",
    ) as HTMLButtonElement;

    expect(generateButton.disabled).toBe(true);

    setInputFiles(worklogInput, [makeWorklogFileMock()]);
    worklogInput.dispatchEvent(new Event("change", { bubbles: true }));
    expect(generateButton.disabled).toBe(true);

    setInputFiles(templateInput, [makeTemplateFileMock()]);
    templateInput.dispatchEvent(new Event("change", { bubbles: true }));
    expect(generateButton.disabled).toBe(false);
  });

  it("enables Excel with worklog only, independent from template", async () => {
    await import("../src/main");

    const worklogInput = document.getElementById(
      "worklogInput",
    ) as HTMLInputElement;
    const generateExcelButton = document.getElementById(
      "generateExcelButton",
    ) as HTMLButtonElement;

    expect(generateExcelButton.disabled).toBe(true);

    setInputFiles(worklogInput, [makeWorklogFileMock()]);
    worklogInput.dispatchEvent(new Event("change", { bubbles: true }));
    expect(generateExcelButton.disabled).toBe(false);
  });

  it("resets current export result when an input changes", async () => {
    await import("../src/main");

    const templateInput = document.getElementById(
      "templateInput",
    ) as HTMLInputElement;
    const worklogInput = document.getElementById(
      "worklogInput",
    ) as HTMLInputElement;
    const generateExcelButton = document.getElementById(
      "generateExcelButton",
    ) as HTMLButtonElement;
    const downloadExcelButton = document.getElementById(
      "downloadExcelButton",
    ) as HTMLButtonElement;
    const generateButton = document.getElementById(
      "generateButton",
    ) as HTMLButtonElement;
    const downloadButton = document.getElementById(
      "downloadButton",
    ) as HTMLButtonElement;
    const uploadForm = document.getElementById("uploadForm") as HTMLFormElement;

    setInputFiles(worklogInput, [makeWorklogFileMock()]);
    worklogInput.dispatchEvent(new Event("change", { bubbles: true }));
    generateExcelButton.click();
    await waitUntil(
      () => generateExcelButton.hidden && !downloadExcelButton.hidden,
    );

    setInputFiles(templateInput, [makeTemplateFileMock()]);
    templateInput.dispatchEvent(new Event("change", { bubbles: true }));
    uploadForm.dispatchEvent(
      new Event("submit", { bubbles: true, cancelable: true }),
    );
    await waitUntil(() => generateButton.hidden && !downloadButton.hidden);

    expect(downloadButton.hidden).toBe(false);

    setInputFiles(worklogInput, [makeWorklogFileMock()]);
    worklogInput.dispatchEvent(new Event("change", { bubbles: true }));

    expect(downloadExcelButton.hidden).toBe(true);
    expect(downloadButton.hidden).toBe(true);
  });

  it("allows generating Excel first and then DOCX", async () => {
    await import("../src/main");

    const templateInput = document.getElementById(
      "templateInput",
    ) as HTMLInputElement;
    const worklogInput = document.getElementById(
      "worklogInput",
    ) as HTMLInputElement;
    const generateExcelButton = document.getElementById(
      "generateExcelButton",
    ) as HTMLButtonElement;
    const downloadExcelButton = document.getElementById(
      "downloadExcelButton",
    ) as HTMLButtonElement;
    const generateButton = document.getElementById(
      "generateButton",
    ) as HTMLButtonElement;
    const downloadButton = document.getElementById(
      "downloadButton",
    ) as HTMLButtonElement;
    const uploadForm = document.getElementById("uploadForm") as HTMLFormElement;

    setInputFiles(worklogInput, [makeWorklogFileMock()]);
    worklogInput.dispatchEvent(new Event("change", { bubbles: true }));
    generateExcelButton.click();
    await waitUntil(
      () => generateExcelButton.hidden && !downloadExcelButton.hidden,
    );

    setInputFiles(templateInput, [makeTemplateFileMock()]);
    templateInput.dispatchEvent(new Event("change", { bubbles: true }));
    uploadForm.dispatchEvent(
      new Event("submit", { bubbles: true, cancelable: true }),
    );
    await waitUntil(() => generateButton.hidden && !downloadButton.hidden);

    expect(downloadExcelButton.hidden).toBe(true);
    expect(downloadButton.disabled).toBe(false);
  });
});
