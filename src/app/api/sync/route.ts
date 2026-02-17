import { NextResponse } from "next/server";
import { getSupabase } from "@/lib/supabase";
import { downloadAndParse } from "@/lib/excel-parser";

/**
 * POST /api/sync
 *
 * Downloads the Excel file from OneDrive (via sharing URL),
 * parses all worksheets, and upserts everything into Supabase.
 *
 * No authentication required to trigger — the sync itself uses
 * the service role key to write to Supabase.
 *
 * Protected by a simple secret token to prevent abuse.
 * Set SYNC_SECRET env var and pass it as ?token=... or Authorization header.
 * If SYNC_SECRET is not set, sync is open (fine for dev/early stage).
 */
export async function POST(request: Request) {
  const start = Date.now();

  // Optional auth check
  const secret = process.env.SYNC_SECRET;
  if (secret) {
    const url = new URL(request.url);
    const token =
      url.searchParams.get("token") ||
      request.headers.get("authorization")?.replace("Bearer ", "");
    if (token !== secret) {
      return NextResponse.json(
        { success: false, error: "Unauthorized" },
        { status: 401 }
      );
    }
  }

  const shareUrl = process.env.ONEDRIVE_SHARE_URL;
  if (!shareUrl) {
    return NextResponse.json(
      {
        success: false,
        error: "ONEDRIVE_SHARE_URL not configured",
        setup: "Set this env var to your Excel file's sharing link",
      },
      { status: 503 }
    );
  }

  try {
    // 1. Download and parse Excel
    const parsed = await downloadAndParse(shareUrl);

    // 2. Get Supabase client
    const db = getSupabase();

    // 3. Sync each table (full replace — delete old, insert fresh)
    const stats: Record<string, number> = {};

    // Trades (main + commodity sheets)
    if (parsed.trades.length > 0) {
      await db.from("trades").delete().gte("id", 0);
      await insertBatched(db, "trades", parsed.trades);
      stats.trades = parsed.trades.length;
    }

    // NV Trades
    if (parsed.nvTrades.length > 0) {
      await db.from("nv_trades").delete().gte("id", 0);
      await insertBatched(db, "nv_trades", parsed.nvTrades);
      stats.nv_trades = parsed.nvTrades.length;
    }

    // TUT Trades
    if (parsed.tutTrades.length > 0) {
      await db.from("tut_trades").delete().gte("id", 0);
      await insertBatched(db, "tut_trades", parsed.tutTrades);
      stats.tut_trades = parsed.tutTrades.length;
    }

    // Currency Rates
    if (parsed.currencyRates.length > 0) {
      await db.from("currency_rates").delete().gte("id", 0);
      await insertBatched(db, "currency_rates", parsed.currencyRates);
      stats.currency_rates = parsed.currencyRates.length;
    }

    // Financials
    if (parsed.financials.length > 0) {
      await db.from("financials").delete().gte("id", 0);
      await insertBatched(db, "financials", parsed.financials);
      stats.financials = parsed.financials.length;
    }

    // MTM
    if (parsed.mtm.length > 0) {
      await db.from("mtm").delete().gte("id", 0);
      await insertBatched(db, "mtm", parsed.mtm);
      stats.mtm = parsed.mtm.length;
    }

    const durationMs = Date.now() - start;
    const totalRows = Object.values(stats).reduce((a, b) => a + b, 0);

    // 4. Log the sync
    await db.from("sync_log").insert({
      status: "success",
      sheets_synced: parsed.sheetsFound,
      total_rows: totalRows,
      duration_ms: durationMs,
    });

    return NextResponse.json({
      success: true,
      stats,
      totalRows,
      sheetsFound: parsed.sheetsFound,
      durationMs,
    });
  } catch (err) {
    const message = err instanceof Error ? err.message : "Sync failed";
    console.error("[/api/sync] Error:", message);

    // Try to log the error
    try {
      const db = getSupabase();
      await db.from("sync_log").insert({
        status: "error",
        error_message: message,
        duration_ms: Date.now() - start,
      });
    } catch {
      // Ignore logging errors
    }

    return NextResponse.json(
      { success: false, error: message },
      { status: 500 }
    );
  }
}

/**
 * Also support GET for Vercel Cron Jobs.
 * Vercel Cron sends GET requests to the configured endpoint.
 */
export async function GET(request: Request) {
  // For cron jobs, verify the CRON_SECRET
  const cronSecret = process.env.CRON_SECRET;
  if (cronSecret) {
    const authHeader = request.headers.get("authorization");
    if (authHeader !== `Bearer ${cronSecret}`) {
      return NextResponse.json(
        { success: false, error: "Unauthorized" },
        { status: 401 }
      );
    }
  }

  // Delegate to POST handler logic
  return POST(request);
}

// ─── Batch Insert (Supabase has row limits per request) ──────────

async function insertBatched(
  db: ReturnType<typeof getSupabase>,
  table: string,
  records: Record<string, unknown>[],
  batchSize = 500
) {
  for (let i = 0; i < records.length; i += batchSize) {
    const batch = records.slice(i, i + batchSize);
    const { error } = await db.from(table).insert(batch);
    if (error) {
      throw new Error(`Insert into ${table} failed (batch ${i}): ${error.message}`);
    }
  }
}
