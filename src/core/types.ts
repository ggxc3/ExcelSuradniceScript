export interface ProcessRequest {
  inputPath: string;
  outputPath: string;
  sheet: string;
  columns: string;
  startRow: number;
  endRow: number;
}
