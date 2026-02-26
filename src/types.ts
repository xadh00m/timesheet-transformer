// Typen für CSV-Parsing und File-Handling
export type CsvRow = string[];
export type CsvData = CsvRow[];

export interface FileUploadResult {
  fileName: string;
  data: CsvData;
}
