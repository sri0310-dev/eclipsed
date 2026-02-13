import { Client } from "@microsoft/microsoft-graph-client";
import { getMsalClient, GRAPH_SCOPES, getRedirectUri } from "./msal-config";
import { getStoredToken, storeToken } from "./token-store";
import type {
  SheetDataResponse,
  SheetRange,
  OneDriveFile,
} from "@/types/onedrive";

/**
 * Core Microsoft Graph client for OneDrive/Excel operations.
 * Supports both owned files and files shared by others.
 *
 * File ID convention:
 *   - Plain ID:       "abc123"          → /me/drive/items/abc123
 *   - Composite ID:   "driveId:itemId"  → /drives/driveId/items/itemId
 *
 * Shared files from another user's OneDrive use the composite format
 * so the Graph API routes to the correct drive.
 */

function getAuthenticatedClient(accessToken: string): Client {
  return Client.init({
    authProvider: (done) => {
      done(null, accessToken);
    },
  });
}

async function getAccessToken(): Promise<string> {
  const cached = getStoredToken();
  if (cached) {
    return cached.accessToken;
  }
  throw new Error(
    "No valid access token. User must authenticate via /api/onedrive/auth first."
  );
}

/**
 * Build the base path for a file's workbook API.
 * Handles both owned files (/me/drive/items/...) and shared files (/drives/.../items/...).
 */
function itemPath(fileId: string): string {
  if (fileId.includes(":")) {
    const [driveId, itemId] = fileId.split(":", 2);
    return `/drives/${driveId}/items/${itemId}`;
  }
  return `/me/drive/items/${fileId}`;
}

// ─── Authentication ────────────────────────────────────────────────────────

export function getAuthUrl(host?: string): string {
  const redirectUri = getRedirectUri(host);
  return `https://login.microsoftonline.com/${process.env.AZURE_TENANT_ID}/oauth2/v2.0/authorize?client_id=${process.env.AZURE_CLIENT_ID}&response_type=code&redirect_uri=${encodeURIComponent(redirectUri)}&scope=${encodeURIComponent(GRAPH_SCOPES.join(" "))}&response_mode=query`;
}

export async function exchangeCodeForToken(code: string, host?: string): Promise<void> {
  const msalClient = getMsalClient();
  const redirectUri = getRedirectUri(host);
  const tokenRequest = {
    code,
    scopes: GRAPH_SCOPES,
    redirectUri,
  };

  const response = await msalClient.acquireTokenByCode(tokenRequest);

  if (response?.accessToken) {
    storeToken({
      accessToken: response.accessToken,
      expiresAt:
        response.expiresOn?.getTime() ?? Date.now() + 3600 * 1000,
    });
  } else {
    throw new Error("Failed to acquire access token from authorization code");
  }
}

// ─── Shared With Me ────────────────────────────────────────────────────────

/**
 * List Excel files that have been shared with the authenticated user.
 * This is the OneDrive equivalent of viewing shared Google Sheets.
 * Returns files with composite IDs (driveId:itemId) for direct use.
 */
export async function listSharedWithMe(): Promise<OneDriveFile[]> {
  const token = await getAccessToken();
  const client = getAuthenticatedClient(token);

  const result = await client
    .api("/me/drive/sharedWithMe")
    .get();

  // Filter to Excel files and build composite IDs
  const files: OneDriveFile[] = [];
  for (const item of result.value || []) {
    const name: string = item.name || "";
    const isExcel =
      name.endsWith(".xlsx") ||
      name.endsWith(".xls") ||
      name.endsWith(".xlsm");
    if (!isExcel) continue;

    // Shared items have remoteItem with the actual driveId + itemId
    const remoteItem = item.remoteItem;
    const driveId = remoteItem?.parentReference?.driveId;
    const itemId = remoteItem?.id || item.id;

    files.push({
      id: item.id,
      name: item.name,
      webUrl: item.webUrl || remoteItem?.webUrl || "",
      size: remoteItem?.size || item.size || 0,
      lastModifiedDateTime:
        remoteItem?.lastModifiedDateTime || item.lastModifiedDateTime || "",
      lastModifiedBy: remoteItem?.lastModifiedBy || item.lastModifiedBy,
      compositeId: driveId ? `${driveId}:${itemId}` : item.id,
      remoteItem,
      parentReference: item.parentReference,
    });
  }

  return files;
}

// ─── File Discovery ────────────────────────────────────────────────────────

export async function searchFiles(
  filename: string
): Promise<OneDriveFile[]> {
  const token = await getAccessToken();
  const client = getAuthenticatedClient(token);

  const result = await client
    .api(`/me/drive/root/search(q='${filename}')`)
    .select("id,name,webUrl,size,lastModifiedDateTime,lastModifiedBy")
    .filter("file/mimeType eq 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet' or file/mimeType eq 'application/vnd.ms-excel'")
    .get();

  return result.value as OneDriveFile[];
}

// ─── Read Operations ───────────────────────────────────────────────────────

export async function readSheetData(
  fileId: string,
  worksheet: string,
  range?: string
): Promise<SheetDataResponse> {
  const token = await getAccessToken();
  const client = getAuthenticatedClient(token);

  const basePath = `${itemPath(fileId)}/workbook/worksheets/${worksheet}`;
  const apiPath = range
    ? `${basePath}/range(address='${range}')`
    : `${basePath}/usedRange`;

  const result: SheetRange = await client.api(apiPath).get();

  // First row is treated as headers
  const headers = (result.values[0] || []).map((h) => String(h ?? ""));
  const rows = result.values.slice(1).map((row) => {
    const obj: Record<string, string | number | boolean | null> = {};
    headers.forEach((header, i) => {
      obj[header] = row[i] ?? null;
    });
    return obj;
  });

  return {
    headers,
    rows,
    rawValues: result.values,
    range: result.address,
  };
}

export async function readCellValue(
  fileId: string,
  worksheet: string,
  cellAddress: string
): Promise<string | number | boolean | null> {
  const token = await getAccessToken();
  const client = getAuthenticatedClient(token);

  const result = await client
    .api(
      `${itemPath(fileId)}/workbook/worksheets/${worksheet}/range(address='${cellAddress}')`
    )
    .get();

  return result.values?.[0]?.[0] ?? null;
}

// ─── Write Operations ──────────────────────────────────────────────────────

export async function writeSheetData(
  fileId: string,
  worksheet: string,
  range: string,
  values: (string | number | boolean | null)[][]
): Promise<SheetRange> {
  const token = await getAccessToken();
  const client = getAuthenticatedClient(token);

  const result = await client
    .api(
      `${itemPath(fileId)}/workbook/worksheets/${worksheet}/range(address='${range}')`
    )
    .patch({ values });

  return result as SheetRange;
}

export async function appendRow(
  fileId: string,
  worksheet: string,
  tableNameOrRange: string,
  values: (string | number | boolean | null)[]
): Promise<void> {
  const token = await getAccessToken();
  const client = getAuthenticatedClient(token);

  await client
    .api(
      `${itemPath(fileId)}/workbook/worksheets/${worksheet}/tables/${tableNameOrRange}/rows`
    )
    .post({ values: [values] });
}

// ─── Worksheet Info ────────────────────────────────────────────────────────

export async function listWorksheets(
  fileId: string
): Promise<{ id: string; name: string; position: number }[]> {
  const token = await getAccessToken();
  const client = getAuthenticatedClient(token);

  const result = await client
    .api(`${itemPath(fileId)}/workbook/worksheets`)
    .get();

  return result.value.map(
    (ws: { id: string; name: string; position: number }) => ({
      id: ws.id,
      name: ws.name,
      position: ws.position,
    })
  );
}
