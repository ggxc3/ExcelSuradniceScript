export interface ProcessRequest {
  inputPath: string;
  outputPath: string;
  sheet: string;
  columns: string;
  startRow: number;
  endRow: number;
}

export interface SheetInfo {
  name: string;
  rowCount: number;
}
