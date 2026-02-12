export interface TokenCache {
  accessToken: string;
  expiresAt: number;
  refreshToken?: string;
}

export interface SheetRange {
  address: string;
  values: (string | number | boolean | null)[][];
  columnCount: number;
  rowCount: number;
}

export interface SheetWriteRequest {
  range: string; // e.g. "A1:D10"
  values: (string | number | boolean | null)[][];
}

export interface SheetReadRequest {
  range?: string; // e.g. "A1:Z100", omit for used range
  worksheet?: string;
}

export interface OneDriveFile {
  id: string;
  name: string;
  webUrl: string;
  size: number;
  lastModifiedDateTime: string;
  lastModifiedBy?: {
    user?: {
      displayName: string;
    };
  };
}

export interface SpreadsheetRow {
  [key: string]: string | number | boolean | null;
}

export interface SheetDataResponse {
  headers: string[];
  rows: SpreadsheetRow[];
  rawValues: (string | number | boolean | null)[][];
  range: string;
  lastModified?: string;
}

export interface ApiResponse<T = unknown> {
  success: boolean;
  data?: T;
  error?: string;
}
