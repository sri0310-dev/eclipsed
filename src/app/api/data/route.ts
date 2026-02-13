import { NextResponse } from "next/server";
import { readFromShareUrl } from "@/lib/sharing-link-client";
import { getCachedData, setCachedData, getCacheAge } from "@/lib/data-cache";

/**
 * GET /api/data
 *
 * Returns the master spreadsheet data as JSON.
 * No authentication required — uses the public OneDrive sharing URL.
 *
 * The .xlsx file is downloaded, parsed with ExcelJS, and cached for 30s.
 * Any visitor gets instant data.
 *
 * Requires: ONEDRIVE_SHARE_URL environment variable.
 * Optional: ONEDRIVE_WORKSHEET_NAME (defaults to first worksheet).
 */
export async function GET() {
  const shareUrl = process.env.ONEDRIVE_SHARE_URL;

  if (!shareUrl) {
    return NextResponse.json(
      {
        success: false,
        error: "ONEDRIVE_SHARE_URL not configured",
        setup: {
          step1: "Open your Excel file in OneDrive",
          step2: 'Click Share → "Anyone with the link" → Copy link',
          step3: "Add ONEDRIVE_SHARE_URL=<your-link> to Vercel environment variables",
        },
      },
      { status: 503 }
    );
  }

  // Serve from cache if fresh
  const cached = getCachedData();
  if (cached) {
    return NextResponse.json({
      success: true,
      data: cached,
      cached: true,
      cacheAge: getCacheAge(),
    });
  }

  // Download and parse
  try {
    const worksheet = process.env.ONEDRIVE_WORKSHEET_NAME || undefined;
    const data = await readFromShareUrl(shareUrl, worksheet);
    setCachedData(data);

    return NextResponse.json({
      success: true,
      data,
      cached: false,
    });
  } catch (err) {
    const message = err instanceof Error ? err.message : "Failed to fetch spreadsheet";
    console.error("[/api/data] Error:", message);
    return NextResponse.json(
      { success: false, error: message },
      { status: 502 }
    );
  }
}
