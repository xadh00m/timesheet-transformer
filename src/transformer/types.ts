export type WorkAreaEntry = { name: string; alias: string };

export type WorklogRow = {
  dateValue: Date | string;
  dateKey: string;
  dateSort: number;
  user: string;
  hours: number;
  description: string;
  key?: string;
  keys?: string[];
};

export type Logger = (line: string) => void;
