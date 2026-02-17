import { NextRequest, NextResponse } from "next/server";
import { getSupabase } from "@/lib/supabase";

/**
 * GET /api/data
 *
 * Reads trade data from Supabase (fast SQL, no Excel download).
 * No authentication required — public read access.
 *
 * Query params:
 *   sheet   - Filter by source_sheet (e.g., "Main Sheet", "Sesame", "RCN")
 *   status  - Filter by position (e.g., "Active", "Inactive")
 *   table   - Which table to query: "trades" (default), "nv", "tut", "currency", "financials", "mtm", "sync"
 */
export async function GET(request: NextRequest) {
  try {
    const db = getSupabase();
    const params = request.nextUrl.searchParams;
    const table = params.get("table") || "trades";

    // ── Sync status ──
    if (table === "sync") {
      const { data, error } = await db
        .from("sync_log")
        .select("*")
        .order("synced_at", { ascending: false })
        .limit(1);

      if (error) throw new Error(error.message);

      return NextResponse.json({
        success: true,
        lastSync: data?.[0] || null,
      });
    }

    // ── Trades (main table) ──
    if (table === "trades") {
      let query = db.from("trades").select("*");

      const sheet = params.get("sheet");
      if (sheet) query = query.eq("source_sheet", sheet);

      const status = params.get("status");
      if (status) query = query.ilike("position", status);

      query = query.order("excel_row", { ascending: true });

      const { data, error } = await query;
      if (error) throw new Error(error.message);

      return NextResponse.json({
        success: true,
        data: data || [],
        count: data?.length || 0,
      });
    }

    // ── NV Trades ──
    if (table === "nv") {
      const { data, error } = await db
        .from("nv_trades")
        .select("*")
        .order("excel_row", { ascending: true });

      if (error) throw new Error(error.message);
      return NextResponse.json({ success: true, data: data || [], count: data?.length || 0 });
    }

    // ── TUT Trades ──
    if (table === "tut") {
      const { data, error } = await db
        .from("tut_trades")
        .select("*")
        .order("excel_row", { ascending: true });

      if (error) throw new Error(error.message);
      return NextResponse.json({ success: true, data: data || [], count: data?.length || 0 });
    }

    // ── Currency Rates ──
    if (table === "currency") {
      const { data, error } = await db
        .from("currency_rates")
        .select("*")
        .order("excel_row", { ascending: true });

      if (error) throw new Error(error.message);
      return NextResponse.json({ success: true, data: data || [], count: data?.length || 0 });
    }

    // ── Financials ──
    if (table === "financials") {
      const { data, error } = await db
        .from("financials")
        .select("*")
        .order("excel_row", { ascending: true });

      if (error) throw new Error(error.message);
      return NextResponse.json({ success: true, data: data || [], count: data?.length || 0 });
    }

    // ── MTM ──
    if (table === "mtm") {
      const { data, error } = await db
        .from("mtm")
        .select("*")
        .order("excel_row", { ascending: true });

      if (error) throw new Error(error.message);
      return NextResponse.json({ success: true, data: data || [], count: data?.length || 0 });
    }

    return NextResponse.json(
      { success: false, error: `Unknown table: ${table}` },
      { status: 400 }
    );
  } catch (err) {
    const message = err instanceof Error ? err.message : "Database query failed";
    console.error("[/api/data] Error:", message);
    return NextResponse.json(
      { success: false, error: message },
      { status: 500 }
    );
  }
}
