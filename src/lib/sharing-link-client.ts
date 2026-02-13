import ExcelJS from "exceljs";
import type { SheetDataResponse } from "@/types/onedrive";

/**
 * Direct-download client for OneDrive sharing links.
 *
 * Bypasses the Microsoft Graph API entirely — downloads the .xlsx file
 * from the public sharing URL and parses it with exceljs.
 *
 * This works even when the authenticated user's tenant has no SPO license,
 * as long as the file owner has set sharing to "Anyone with the link".
 *
 * Analogous to downloading a Google Sheet via its public export URL.
 */

/**
 * Convert a OneDrive/SharePoint sharing URL to a direct download URL.
 *
 * Supports:
 *   - Short links:  https://1drv.ms/x/c/ABC123/EfGhIjK
 *   - Live links:   https://onedrive.live.com/edit.aspx?resid=...
 *   - Embed links:  https://onedrive.live.com/embed?resid=...
 */
function toDownloadUrl(shareUrl: string): string {
  const url = new URL(shareUrl);

  // Short links: append download=1, the redirect chain handles the rest
  if (url.hostname === "1drv.ms") {
    url.searchParams.set("download", "1");
    return url.toString();
  }

  // onedrive.live.com links: swap edit for download
  if (url.hostname.includes("onedrive.live.com")) {
    url.pathname = url.pathname
      .replace(/\/edit\.aspx/i, "/download.aspx")
      .replace(/\/embed/i, "/download");
    url.searchParams.set("download", "1");
    return url.toString();
  }

  // SharePoint links
  if (url.hostname.includes("sharepoint.com")) {
    url.searchParams.set("download", "1");
    return url.toString();
  }

  // Fallback
  url.searchParams.set("download", "1");
  return url.toString();
}

/**
 * Download an Excel file from a OneDrive sharing URL and parse it.
 * No authentication needed — uses the public sharing link directly.
 */
export async function readFromShareUrl(
  shareUrl: string,
  worksheetName?: string
): Promise<SheetDataResponse> {
  const downloadUrl = toDownloadUrl(shareUrl);

  const response = await fetch(downloadUrl, {
    redirect: "follow",
    headers: {
      "User-Agent": "HectarControlTower/1.0",
    },
  });

  if (!response.ok) {
    throw new Error(
      `Failed to download file (HTTP ${response.status}). ` +
        `Ensure the sharing link uses "Anyone with the link" access.`
    );
  }

  const arrayBuffer = await response.arrayBuffer();

  const workbook = new ExcelJS.Workbook();
  await workbook.xlsx.load(arrayBuffer);

  // Find the requested worksheet or use the first one
  const worksheet = worksheetName
    ? workbook.getWorksheet(worksheetName)
    : workbook.worksheets[0];

  if (!worksheet) {
    const available = workbook.worksheets.map((ws) => ws.name).join(", ");
    throw new Error(
      `Worksheet "${worksheetName}" not found. Available: ${available}`
    );
  }

  // Convert worksheet to 2D array
  const values: (string | number | boolean | null)[][] = [];
  worksheet.eachRow({ includeEmpty: false }, (row) => {
    const rowValues: (string | number | boolean | null)[] = [];
    for (let col = 1; col <= worksheet.columnCount; col++) {
      const cell = row.getCell(col);
      const val = cell.value;

      // Handle ExcelJS value types
      if (val === null || val === undefined) {
        rowValues.push(null);
      } else if (typeof val === "object" && "result" in val) {
        // Formula cell — use the computed result
        rowValues.push(val.result as string | number | boolean | null);
      } else if (typeof val === "object" && "richText" in val) {
        // Rich text — concatenate text parts
        const rich = val as { richText: { text: string }[] };
        rowValues.push(rich.richText.map((t) => t.text).join(""));
      } else if (val instanceof Date) {
        rowValues.push(val.toISOString());
      } else if (typeof val === "object" && "hyperlink" in val) {
        // Hyperlink cell
        rowValues.push(
          (val as { text?: string; hyperlink: string }).text ||
            (val as { hyperlink: string }).hyperlink
        );
      } else {
        rowValues.push(val as string | number | boolean);
      }
    }
    values.push(rowValues);
  });

  if (values.length === 0) {
    return { headers: [], rows: [], rawValues: [], range: "A1" };
  }

  // First row = headers
  const headers = values[0].map((h) => String(h ?? ""));
  const rows = values.slice(1).map((row) => {
    const obj: Record<string, string | number | boolean | null> = {};
    headers.forEach((header, i) => {
      obj[header] = row[i] ?? null;
    });
    return obj;
  });

  // Build range string
  const lastCol = String.fromCharCode(
    64 + Math.min(values[0].length, 26)
  );
  const range = `A1:${lastCol}${values.length}`;

  return { headers, rows, rawValues: values, range };
}

/**
 * List worksheet names from a shared Excel file.
 */
export async function listWorksheetsFromShareUrl(
  shareUrl: string
): Promise<{ name: string; position: number }[]> {
  const downloadUrl = toDownloadUrl(shareUrl);

  const response = await fetch(downloadUrl, {
    redirect: "follow",
    headers: { "User-Agent": "HectarControlTower/1.0" },
  });

  if (!response.ok) {
    throw new Error(`Failed to download file (HTTP ${response.status}).`);
  }

  const arrayBuffer = await response.arrayBuffer();
  const workbook = new ExcelJS.Workbook();
  await workbook.xlsx.load(arrayBuffer);

  return workbook.worksheets.map((ws, i) => ({
    name: ws.name,
    position: i,
  }));
}
