import { NextRequest, NextResponse } from "next/server";
import {
  readSheetData,
  writeSheetData,
  listWorksheets,
  searchFiles,
} from "@/lib/graph-client";
import { isAuthenticated } from "@/lib/token-store";
import type { ApiResponse, SheetDataResponse } from "@/types/onedrive";

function unauthorized(): NextResponse<ApiResponse> {
  return NextResponse.json(
    { success: false, error: "Not authenticated. Please connect to Microsoft first." },
    { status: 401 }
  );
}

/**
 * GET /api/onedrive/sheets
 *
 * Query params:
 *   action: "read" | "worksheets" | "search" | "status"
 *   fileId: OneDrive file ID (or uses env default)
 *   worksheet: worksheet name (or uses env default)
 *   range: cell range like "A1:Z100" (optional, defaults to used range)
 *   filename: search query for file discovery
 */
export async function GET(request: NextRequest) {
  const params = request.nextUrl.searchParams;
  const action = params.get("action") || "read";

  // Status check doesn't require auth
  if (action === "status") {
    return NextResponse.json({
      success: true,
      data: {
        authenticated: isAuthenticated(),
        configuredFileId: process.env.ONEDRIVE_FILE_ID || null,
        configuredWorksheet: process.env.ONEDRIVE_WORKSHEET_NAME || null,
      },
    });
  }

  if (!isAuthenticated()) return unauthorized();

  const fileId = params.get("fileId") || process.env.ONEDRIVE_FILE_ID;
  const worksheet =
    params.get("worksheet") || process.env.ONEDRIVE_WORKSHEET_NAME || "Sheet1";

  try {
    switch (action) {
      case "read": {
        if (!fileId) {
          return NextResponse.json(
            {
              success: false,
              error:
                "No file ID provided. Set ONEDRIVE_FILE_ID env var or pass ?fileId=...",
            },
            { status: 400 }
          );
        }
        const range = params.get("range") || undefined;
        const data: SheetDataResponse = await readSheetData(
          fileId,
          worksheet,
          range
        );
        return NextResponse.json({ success: true, data });
      }

      case "worksheets": {
        if (!fileId) {
          return NextResponse.json(
            { success: false, error: "No file ID provided." },
            { status: 400 }
          );
        }
        const sheets = await listWorksheets(fileId);
        return NextResponse.json({ success: true, data: sheets });
      }

      case "search": {
        const filename = params.get("filename");
        if (!filename) {
          return NextResponse.json(
            { success: false, error: "filename query param is required for search" },
            { status: 400 }
          );
        }
        const files = await searchFiles(filename);
        return NextResponse.json({ success: true, data: files });
      }

      default:
        return NextResponse.json(
          { success: false, error: `Unknown action: ${action}` },
          { status: 400 }
        );
    }
  } catch (err) {
    const message = err instanceof Error ? err.message : "Unknown error";
    console.error("[OneDrive API] GET error:", message);
    return NextResponse.json(
      { success: false, error: message },
      { status: 500 }
    );
  }
}

/**
 * POST /api/onedrive/sheets
 *
 * Body:
 *   { action: "write", fileId?, worksheet?, range: "A1:D5", values: [[...], [...]] }
 */
export async function POST(request: NextRequest) {
  if (!isAuthenticated()) return unauthorized();

  try {
    const body = await request.json();
    const { action, range, values } = body;
    const fileId = body.fileId || process.env.ONEDRIVE_FILE_ID;
    const worksheet =
      body.worksheet || process.env.ONEDRIVE_WORKSHEET_NAME || "Sheet1";

    if (action !== "write") {
      return NextResponse.json(
        { success: false, error: `Unknown action: ${action}. Use "write".` },
        { status: 400 }
      );
    }

    if (!fileId) {
      return NextResponse.json(
        { success: false, error: "No file ID provided." },
        { status: 400 }
      );
    }

    if (!range || !values || !Array.isArray(values)) {
      return NextResponse.json(
        {
          success: false,
          error: "Both 'range' (e.g. 'A1:D5') and 'values' (2D array) are required.",
        },
        { status: 400 }
      );
    }

    const result = await writeSheetData(fileId, worksheet, range, values);
    return NextResponse.json({ success: true, data: result });
  } catch (err) {
    const message = err instanceof Error ? err.message : "Unknown error";
    console.error("[OneDrive API] POST error:", message);
    return NextResponse.json(
      { success: false, error: message },
      { status: 500 }
    );
  }
}
