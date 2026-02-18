"use client";

import { useState, useEffect, useCallback, useMemo } from "react";

// ─── Types ────────────────────────────────────────────────────────────────────

interface Trade {
  id: number;
  source_sheet: string;
  product: string | null;
  position: string | null;
  month: string | null;
  commodity_code: string | null;
  origin: string | null;
  variety: string | null;
  packer: string | null;
  quantity_mt: number | null;
  no_of_containers: number | null;
  purchase_price_per_mt: number | null;
  sales_price_per_mt: number | null;
  purchase_value: number | null;
  sales_value: number | null;
  gross_margin: number | null;
  net_profit: number | null;
  profit_pct: number | null;
  total_expenses: number | null;
  bl_number: string | null;
  etd: string | null;
  eta: string | null;
  seller: string | null;
  buyer: string | null;
  port_of_loading: string | null;
  port_of_discharge: string | null;
  trade_no: string | null;
  advance_paid: number | null;
  advance_from_buyer: number | null;
  total_outwards: number | null;
  total_inwards: number | null;
  outward_remaining: number | null;
  inward_remaining: number | null;
  working_capital_days: number | null;
  clearance_charges: number | null;
  brokerage: number | null;
  warehouse_loss: number | null;
  claims_paid: number | null;
  claims_received: number | null;
  interest_loss: number | null;
  payment_terms: string | null;
  buyer_payment_term: string | null;
  transit_days: number | null;
  [key: string]: unknown;
}

interface SyncStatus {
  synced_at: string;
  status: string;
  total_rows: number;
  duration_ms: number;
  sheets_synced: string[];
}

// ─── Helpers ──────────────────────────────────────────────────────────────────

function fmt(n: number | null | undefined): string {
  if (n === null || n === undefined || isNaN(n)) return "—";
  return new Intl.NumberFormat("en-US", { maximumFractionDigits: 0 }).format(n);
}

function fmtCurrency(n: number | null | undefined): string {
  if (n === null || n === undefined || isNaN(n)) return "—";
  if (Math.abs(n) >= 1_000_000) {
    return "$" + (n / 1_000_000).toFixed(1) + "M";
  }
  if (Math.abs(n) >= 1_000) {
    return "$" + (n / 1_000).toFixed(0) + "K";
  }
  return new Intl.NumberFormat("en-US", {
    style: "currency",
    currency: "USD",
    maximumFractionDigits: 0,
  }).format(n);
}

function fmtCurrencyFull(n: number | null | undefined): string {
  if (n === null || n === undefined || isNaN(n)) return "—";
  return new Intl.NumberFormat("en-US", {
    style: "currency",
    currency: "USD",
    maximumFractionDigits: 0,
  }).format(n);
}

function pct(part: number, total: number): number {
  return total > 0 ? Math.round((part / total) * 100) : 0;
}

function riskColor(percentage: number): string {
  if (percentage >= 40) return "text-red-400";
  if (percentage >= 25) return "text-amber-400";
  return "text-emerald-400";
}

function riskLabel(percentage: number): string {
  if (percentage >= 40) return "Concentrated";
  if (percentage >= 25) return "Moderate";
  return "Diversified";
}

function countBy(trades: Trade[], field: keyof Trade): Map<string, number> {
  const map = new Map<string, number>();
  for (const t of trades) {
    const v = String(t[field] ?? "").trim();
    if (v && v !== "—") map.set(v, (map.get(v) || 0) + 1);
  }
  return map;
}

/** Aggregate a numeric field by a group key */
function sumByGroup(trades: Trade[], groupField: keyof Trade, sumField: keyof Trade): Map<string, number> {
  const map = new Map<string, number>();
  for (const t of trades) {
    const key = String(t[groupField] ?? "").trim();
    if (!key || key === "—") continue;
    map.set(key, (map.get(key) || 0) + (Number(t[sumField]) || 0));
  }
  return map;
}

function topN(map: Map<string, number>, n: number): [string, number][] {
  return [...map.entries()].sort((a, b) => b[1] - a[1]).slice(0, n);
}

function sumBy(trades: Trade[], field: keyof Trade): number {
  return trades.reduce((s, t) => s + (Number(t[field]) || 0), 0);
}

const MONTHS = [
  "january", "february", "march", "april", "may", "june",
  "july", "august", "september", "october", "november", "december",
];

function monthIndex(name: string): number {
  const l = name.toLowerCase().trim();
  const i = MONTHS.findIndex((m) => m.startsWith(l) || l.startsWith(m.slice(0, 3)));
  return i >= 0 ? i : 99;
}

function timeAgo(dateStr: string): string {
  const diff = Date.now() - new Date(dateStr).getTime();
  const mins = Math.floor(diff / 60000);
  if (mins < 1) return "just now";
  if (mins < 60) return `${mins}m ago`;
  const hrs = Math.floor(mins / 60);
  if (hrs < 24) return `${hrs}h ago`;
  return `${Math.floor(hrs / 24)}d ago`;
}

/**
 * Position semantics:
 * - "Inactive" = fully completed, all transactions done (gray)
 * - "Sold" = at port, financial transactions in progress (blue) — needs immediate attention
 * - Everything else = in transit / at origin / open contracts (default) — needs monitoring + MTM
 */
function positionStage(pos: string): "completed" | "at_port" | "open" {
  const p = pos.toLowerCase().trim();
  if (p === "inactive") return "completed";
  if (p === "sold") return "at_port";
  return "open";
}

function shipmentStage(t: Trade): string {
  const now = new Date();
  const etd = t.etd ? new Date(t.etd) : null;
  const eta = t.eta ? new Date(t.eta) : null;
  if (eta && now >= eta) return "Arrived";
  if (etd && eta && now >= etd && now < eta) return "In Transit";
  if (etd && now < etd) return "Awaiting Shipment";
  if (t.bl_number) return "B/L Issued";
  return "Booking";
}

// ─── Bar component ────────────────────────────────────────────────────────────

function Bar({ label, value, max, color, suffix }: { label: string; value: number; max: number; color: string; suffix?: string }) {
  const w = max > 0 ? Math.max(6, (Math.abs(value) / Math.abs(max)) * 100) : 0;
  const display = suffix === "$" ? fmtCurrency(value) : String(value);
  return (
    <div className="flex items-center gap-2 text-xs">
      <span className="w-28 truncate text-zinc-300 text-right shrink-0">{label}</span>
      <div className="flex-1 h-4 bg-zinc-800 rounded-sm overflow-hidden">
        <div className={`h-full ${color} rounded-sm`} style={{ width: `${w}%` }} />
      </div>
      <span className="w-16 text-zinc-500 text-right font-mono shrink-0 text-[11px]">{display}</span>
    </div>
  );
}

// ─── Section header ──────────────────────────────────────────────────────────

function SectionHeader({ title, subtitle }: { title: string; subtitle?: string }) {
  return (
    <div className="pt-3 pb-1">
      <h2 className="text-sm font-semibold text-zinc-200">{title}</h2>
      {subtitle && <p className="text-[11px] text-zinc-500">{subtitle}</p>}
    </div>
  );
}

// ─── Main Component ───────────────────────────────────────────────────────────

export default function ControlTower() {
  const [trades, setTrades] = useState<Trade[]>([]);
  const [loading, setLoading] = useState(true);
  const [error, setError] = useState<string | null>(null);
  const [syncStatus, setSyncStatus] = useState<SyncStatus | null>(null);
  const [syncing, setSyncing] = useState(false);
  const [positionFilter, setPositionFilter] = useState<string>("all");
  const [sheetFilter, setSheetFilter] = useState<string>("all");

  // ─── Fetch trades from Supabase ──────────────────────────────

  const fetchTrades = useCallback(async () => {
    try {
      const res = await fetch("/api/data?table=trades");
      const result = await res.json();
      if (result.success) {
        setTrades(result.data);
        setError(null);
      } else {
        setError(result.error);
      }
    } catch {
      setError("Failed to load data");
    } finally {
      setLoading(false);
    }
  }, []);

  const fetchSyncStatus = useCallback(async () => {
    try {
      const res = await fetch("/api/data?table=sync");
      const result = await res.json();
      if (result.success && result.lastSync) {
        setSyncStatus(result.lastSync);
      }
    } catch {
      // Ignore
    }
  }, []);

  // ─── Trigger sync ────────────────────────────────────────────

  const triggerSync = async () => {
    setSyncing(true);
    setError(null);
    try {
      const res = await fetch("/api/sync", { method: "POST" });
      const result = await res.json();
      if (result.success) {
        await fetchTrades();
        await fetchSyncStatus();
      } else {
        setError(result.error);
      }
    } catch {
      setError("Sync failed — check console");
    } finally {
      setSyncing(false);
    }
  };

  // ─── Auto-sync: show cached data immediately, refresh if stale ──

  useEffect(() => {
    fetchTrades();
    fetchSyncStatus();
  }, []); // eslint-disable-line react-hooks/exhaustive-deps

  useEffect(() => {
    if (loading || syncing) return;
    const shouldAutoSync =
      trades.length === 0 ||
      !syncStatus ||
      (Date.now() - new Date(syncStatus.synced_at).getTime() > 5 * 60 * 1000);
    if (shouldAutoSync) {
      triggerSync();
    }
  }, [loading, trades.length, syncStatus]); // eslint-disable-line react-hooks/exhaustive-deps

  useEffect(() => {
    if (trades.length === 0) return;
    const id = setInterval(() => {
      fetchTrades();
      fetchSyncStatus();
    }, 60_000);
    return () => clearInterval(id);
  }, [trades.length, fetchTrades, fetchSyncStatus]);

  // ─── Derived analytics ──────────────────────────────────────

  const analytics = useMemo(() => {
    if (trades.length === 0) return null;

    const filtered = sheetFilter === "all"
      ? trades
      : trades.filter((t) => t.source_sheet === sheetFilter);

    // Position groups
    const positionGroups = new Map<string, number>();
    for (const t of filtered) {
      const pos = (t.position ?? "Unknown").trim() || "Unknown";
      positionGroups.set(pos, (positionGroups.get(pos) || 0) + 1);
    }
    const positions = [...positionGroups.entries()].sort((a, b) => b[1] - a[1]);

    // Apply position filter for the "pool" (what we compute analytics on)
    const pool = positionFilter === "all"
      ? filtered
      : filtered.filter((t) => (t.position ?? "").trim() === positionFilter);

    // ── Lifecycle stages (always based on full filtered set, not position filter) ──
    const lifecycle = {
      open: filtered.filter(t => positionStage(t.position ?? "") === "open"),
      atPort: filtered.filter(t => positionStage(t.position ?? "") === "at_port"),
      completed: filtered.filter(t => positionStage(t.position ?? "") === "completed"),
    };

    // ── Deals requiring attention: "Sold" trades at port with open financials ──
    const dealsAtPort = lifecycle.atPort.map(t => ({
      ...t,
      hasOpenPayable: (Number(t.outward_remaining) || 0) > 0,
      hasOpenReceivable: (Number(t.inward_remaining) || 0) > 0,
      shipStage: shipmentStage(t),
    }));

    // ── Open/Active trades needing monitoring ──
    const openTrades = lifecycle.open.map(t => {
      const hasBuyer = !!t.buyer;
      const hasSeller = !!t.seller;
      const isClosedContract = hasBuyer && hasSeller;
      return { ...t, isClosedContract, shipStage: shipmentStage(t) };
    });

    // ── Shipment pipeline ──
    const allNonCompleted = [...lifecycle.open, ...lifecycle.atPort];
    const shipmentStages = new Map<string, number>();
    for (const t of allNonCompleted) {
      const stage = shipmentStage(t);
      shipmentStages.set(stage, (shipmentStages.get(stage) || 0) + 1);
    }
    const shipmentPipeline = ["Booking", "B/L Issued", "Awaiting Shipment", "In Transit", "Arrived"]
      .map(s => ({ stage: s, count: shipmentStages.get(s) || 0 }))
      .filter(s => s.count > 0);

    // Basic counts
    const byProduct = countBy(pool, "product");
    const byOrigin = countBy(pool, "origin");
    const byMonth = countBy(pool, "month");
    const byPacker = countBy(pool, "packer");
    const byBuyer = countBy(pool, "buyer");
    const sheets = [...new Set(trades.map((t) => t.source_sheet))].sort();
    const monthSorted: [string, number][] = [...byMonth.entries()].sort(
      (a, b) => monthIndex(a[0]) - monthIndex(b[0])
    );

    // ── Exposure by VALUE (not just count) ──
    const buyerExposureByValue = sumByGroup(pool, "buyer", "sales_value");
    const sellerExposureByValue = sumByGroup(pool, "seller", "purchase_value");
    const productExposureByValue = sumByGroup(pool, "product", "purchase_value");
    const originExposureByValue = sumByGroup(pool, "origin", "purchase_value");

    // ── P&L Heatmap: commodity × origin by net profit ──
    const topProds = topN(byProduct, 6).map(([n]) => n);
    const topOrigs = topN(byOrigin, 6).map(([n]) => n);
    const profitMatrix = topProds.map((p) =>
      topOrigs.map((o) => {
        const matching = pool.filter((t) => t.product === p && t.origin === o);
        return {
          count: matching.length,
          profit: matching.reduce((s, t) => s + (Number(t.net_profit) || 0), 0),
          revenue: matching.reduce((s, t) => s + (Number(t.sales_value) || 0), 0),
        };
      })
    );

    // Concentration risk
    const topProduct = topN(byProduct, 1)[0];
    const topOrigin = topN(byOrigin, 1)[0];
    const topPacker = topN(byPacker, 1)[0];
    const total = pool.length || 1;

    // Financial summaries
    const totalPurchaseValue = sumBy(pool, "purchase_value");
    const totalSalesValue = sumBy(pool, "sales_value");
    const totalGrossMargin = sumBy(pool, "gross_margin");
    const totalNetProfit = sumBy(pool, "net_profit");
    const totalContainers = sumBy(pool, "no_of_containers");
    const totalQuantity = sumBy(pool, "quantity_mt");

    // Cash flow
    const totalOutwards = sumBy(pool, "total_outwards");
    const totalInwards = sumBy(pool, "total_inwards");
    const outwardRemaining = sumBy(pool, "outward_remaining");
    const inwardRemaining = sumBy(pool, "inward_remaining");
    const wcDays = pool.map(t => Number(t.working_capital_days) || 0).filter(d => d > 0);
    const avgWorkingCapitalDays = wcDays.length > 0
      ? wcDays.reduce((a, b) => a + b, 0) / wcDays.length
      : 0;

    // Working capital distribution
    const wcBuckets = [
      { label: "0–30d", count: wcDays.filter(d => d <= 30).length },
      { label: "31–60d", count: wcDays.filter(d => d > 30 && d <= 60).length },
      { label: "61–90d", count: wcDays.filter(d => d > 60 && d <= 90).length },
      { label: "90d+", count: wcDays.filter(d => d > 90).length },
    ].filter(b => b.count > 0);

    // Expense breakdown
    const totalClearanceCharges = sumBy(pool, "clearance_charges");
    const totalBrokerage = sumBy(pool, "brokerage");
    const totalWarehouseLoss = sumBy(pool, "warehouse_loss");
    const totalClaimsPaid = sumBy(pool, "claims_paid");
    const totalClaimsReceived = sumBy(pool, "claims_received");
    const totalInterestLoss = sumBy(pool, "interest_loss");
    const totalExpenses = sumBy(pool, "total_expenses");

    // ── Open Receivables (buyers who still owe us) ──
    const openReceivables = pool
      .filter(t => (Number(t.inward_remaining) || 0) > 0)
      .map(t => ({
        buyer: t.buyer || "Unknown",
        product: t.product || "—",
        amount: Number(t.inward_remaining) || 0,
        trade_no: t.trade_no || "—",
        position: t.position || "—",
      }))
      .sort((a, b) => b.amount - a.amount);

    // ── Open Payables (sellers we still owe) ──
    const openPayables = pool
      .filter(t => (Number(t.outward_remaining) || 0) > 0)
      .map(t => ({
        seller: t.seller || "Unknown",
        product: t.product || "—",
        amount: Number(t.outward_remaining) || 0,
        trade_no: t.trade_no || "—",
        position: t.position || "—",
      }))
      .sort((a, b) => b.amount - a.amount);

    // ── P&L by commodity ──
    const profitByProduct = sumByGroup(pool, "product", "net_profit");

    return {
      filtered, sheets, pool, positions,
      lifecycle, dealsAtPort, openTrades,
      shipmentPipeline,
      byProduct: topN(byProduct, 10),
      byOrigin: topN(byOrigin, 10),
      byMonth: monthSorted,
      byPacker: topN(byPacker, 10),
      byBuyer: topN(byBuyer, 10),
      buyerExposureByValue: topN(buyerExposureByValue, 10),
      sellerExposureByValue: topN(sellerExposureByValue, 10),
      productExposureByValue: topN(productExposureByValue, 10),
      originExposureByValue: topN(originExposureByValue, 10),
      profitByProduct: topN(profitByProduct, 10),
      risk: {
        product: topProduct ? { name: topProduct[0], pct: pct(topProduct[1], total) } : null,
        origin: topOrigin ? { name: topOrigin[0], pct: pct(topOrigin[1], total) } : null,
        packer: topPacker ? { name: topPacker[0], pct: pct(topPacker[1], total) } : null,
      },
      profitMatrix: { products: topProds, origins: topOrigs, values: profitMatrix },
      financials: { totalPurchaseValue, totalSalesValue, totalGrossMargin, totalNetProfit, totalContainers, totalQuantity },
      cashFlow: { totalOutwards, totalInwards, outwardRemaining, inwardRemaining, avgWorkingCapitalDays },
      expenses: { totalClearanceCharges, totalBrokerage, totalWarehouseLoss, totalClaimsPaid, totalClaimsReceived, totalInterestLoss, totalExpenses },
      wcBuckets,
      openReceivables,
      openPayables,
    };
  }, [trades, sheetFilter, positionFilter]);

  const displayRows = analytics ? analytics.pool : [];

  // ─── Visible columns ──────────────────────────────────────────

  const columns: { key: keyof Trade; label: string; align?: "right" }[] = [
    { key: "product", label: "Product" },
    { key: "position", label: "Status" },
    { key: "month", label: "Month" },
    { key: "origin", label: "Origin" },
    { key: "packer", label: "Packer" },
    { key: "no_of_containers", label: "Ctrs", align: "right" },
    { key: "quantity_mt", label: "Qty (MT)", align: "right" },
    { key: "purchase_price_per_mt", label: "Buy $/MT", align: "right" },
    { key: "sales_price_per_mt", label: "Sell $/MT", align: "right" },
    { key: "gross_margin", label: "Margin", align: "right" },
    { key: "net_profit", label: "Net P&L", align: "right" },
    { key: "buyer", label: "Buyer" },
    { key: "seller", label: "Seller" },
    { key: "etd", label: "ETD" },
    { key: "eta", label: "ETA" },
    { key: "bl_number", label: "BL#" },
    { key: "trade_no", label: "Trade#" },
  ];

  // ═══════════════════════════════════════════════════════════════
  // RENDER
  // ═══════════════════════════════════════════════════════════════

  return (
    <div className="min-h-screen bg-zinc-950 text-zinc-100">
      {/* ─── Header ─────────────────────────────────────────────── */}
      <header className="border-b border-zinc-800 bg-zinc-900/50 backdrop-blur-sm sticky top-0 z-10">
        <div className="max-w-[1600px] mx-auto px-6 py-3 flex items-center justify-between">
          <div>
            <h1 className="text-lg font-bold tracking-tight">Hectar Control Tower</h1>
            <p className="text-[11px] text-zinc-500">Commodity Trading & Risk Management</p>
          </div>
          <div className="flex items-center gap-4">
            {syncStatus && (
              <span className="text-[11px] text-zinc-600">
                Synced {timeAgo(syncStatus.synced_at)} · {syncStatus.total_rows} rows
              </span>
            )}
            <button
              onClick={triggerSync}
              disabled={syncing}
              className="px-3 py-1 text-xs bg-zinc-800 hover:bg-zinc-700 text-zinc-300 rounded border border-zinc-700 transition-colors disabled:opacity-50"
            >
              {syncing ? "Syncing..." : "Refresh"}
            </button>
            {trades.length > 0 && (
              <div className="flex items-center gap-1.5">
                <div className="h-2 w-2 rounded-full bg-emerald-500 animate-pulse" />
                <span className="text-[11px] text-zinc-400">Live</span>
              </div>
            )}
          </div>
        </div>
      </header>

      <main className="max-w-[1600px] mx-auto px-6 py-5 space-y-5">
        {/* ─── Loading ────────────────────────────────────────── */}
        {loading && (
          <div className="space-y-4">
            <div className="grid grid-cols-2 md:grid-cols-4 lg:grid-cols-8 gap-3">
              {[1, 2, 3, 4, 5, 6, 7, 8].map((i) => (
                <div key={i} className="rounded-lg border border-zinc-800 bg-zinc-900/30 p-4 animate-pulse">
                  <div className="h-7 bg-zinc-800 rounded w-12 mb-1.5" />
                  <div className="h-3 bg-zinc-800/50 rounded w-16" />
                </div>
              ))}
            </div>
            <div className="text-center text-sm text-zinc-500 py-2">Loading...</div>
          </div>
        )}

        {/* ─── Empty state ──────────────────────────────────── */}
        {!loading && trades.length === 0 && !error && (
          <div className="max-w-md mx-auto py-16 text-center space-y-5">
            <div className="inline-flex h-16 w-16 items-center justify-center rounded-full bg-zinc-900 border border-zinc-800">
              <svg className="h-8 w-8 text-zinc-400" fill="none" viewBox="0 0 24 24" stroke="currentColor" strokeWidth={1.5}>
                <path strokeLinecap="round" strokeLinejoin="round" d="M3 16.5v2.25A2.25 2.25 0 005.25 21h13.5A2.25 2.25 0 0021 18.75V16.5m-13.5-9L12 3m0 0l4.5 4.5M12 3v13.5" />
              </svg>
            </div>
            <div>
              <h2 className="text-lg font-semibold text-zinc-100 mb-1">Ready to sync</h2>
              <p className="text-sm text-zinc-400">Pull the latest data from your Excel spreadsheet.</p>
            </div>
            <button
              onClick={triggerSync}
              disabled={syncing}
              className="px-6 py-2.5 bg-blue-600 hover:bg-blue-700 disabled:bg-zinc-700 text-white text-sm font-medium rounded-lg transition-colors"
            >
              {syncing ? "Syncing..." : "Run First Sync"}
            </button>
          </div>
        )}

        {/* ─── Error ────────────────────────────────────────── */}
        {error && (
          <div className="p-3 bg-red-900/20 border border-red-800 rounded-lg text-xs text-red-400">
            {error}
          </div>
        )}

        {/* ═══ DASHBOARD ═══ */}
        {analytics && (
          <>
            {/* ── Filters ── */}
            <div className="flex flex-wrap items-center gap-4">
              {analytics.sheets.length > 1 && (
                <div className="flex items-center gap-2 flex-wrap">
                  <span className="text-[11px] text-zinc-500">Sheet:</span>
                  <button onClick={() => setSheetFilter("all")}
                    className={`px-2.5 py-1 text-[11px] rounded-full border transition-colors ${sheetFilter === "all" ? "border-blue-700 bg-blue-900/30 text-blue-400" : "border-zinc-700 bg-zinc-800/50 text-zinc-400 hover:border-zinc-600"}`}>
                    All
                  </button>
                  {analytics.sheets.map((s) => (
                    <button key={s} onClick={() => setSheetFilter(s)}
                      className={`px-2.5 py-1 text-[11px] rounded-full border transition-colors ${sheetFilter === s ? "border-blue-700 bg-blue-900/30 text-blue-400" : "border-zinc-700 bg-zinc-800/50 text-zinc-400 hover:border-zinc-600"}`}>
                      {s.trim()}
                    </button>
                  ))}
                </div>
              )}
              <div className="flex items-center gap-2 flex-wrap">
                <span className="text-[11px] text-zinc-500">Position:</span>
                <button onClick={() => setPositionFilter("all")}
                  className={`px-2.5 py-1 text-[11px] rounded-full border transition-colors ${positionFilter === "all" ? "border-blue-700 bg-blue-900/30 text-blue-400" : "border-zinc-700 bg-zinc-800/50 text-zinc-400 hover:border-zinc-600"}`}>
                  All ({analytics.filtered.length})
                </button>
                {analytics.positions.map(([pos, count]) => {
                  const stage = positionStage(pos);
                  const chipColor = positionFilter === pos
                    ? "border-blue-700 bg-blue-900/30 text-blue-400"
                    : stage === "completed" ? "border-zinc-700 bg-zinc-800/50 text-zinc-500 hover:border-zinc-600"
                    : stage === "at_port" ? "border-blue-800/50 bg-blue-950/30 text-blue-400 hover:border-blue-700"
                    : "border-zinc-700 bg-zinc-800/50 text-zinc-400 hover:border-zinc-600";
                  return (
                    <button key={pos} onClick={() => setPositionFilter(pos)}
                      className={`px-2.5 py-1 text-[11px] rounded-full border transition-colors ${chipColor}`}>
                      {pos} ({count})
                    </button>
                  );
                })}
              </div>
            </div>

            {/* ═══ 1. LIFECYCLE OVERVIEW ═══ */}
            <div className="grid grid-cols-3 gap-3">
              <div className="rounded-lg border border-zinc-700 bg-zinc-900/60 px-4 py-3">
                <div className="flex items-center gap-2 mb-1">
                  <div className="h-2.5 w-2.5 rounded-full bg-amber-500" />
                  <span className="text-[11px] text-zinc-400">Open / In Transit</span>
                </div>
                <div className="text-2xl font-bold text-amber-400">{analytics.lifecycle.open.length}</div>
                <div className="text-[11px] text-zinc-600">Monitoring + MTM needed</div>
              </div>
              <div className="rounded-lg border border-blue-800/40 bg-blue-950/20 px-4 py-3">
                <div className="flex items-center gap-2 mb-1">
                  <div className="h-2.5 w-2.5 rounded-full bg-blue-500" />
                  <span className="text-[11px] text-blue-300/70">At Port / Sold</span>
                </div>
                <div className="text-2xl font-bold text-blue-400">{analytics.lifecycle.atPort.length}</div>
                <div className="text-[11px] text-blue-400/40">Financial close in progress</div>
              </div>
              <div className="rounded-lg border border-zinc-800 bg-zinc-900/40 px-4 py-3">
                <div className="flex items-center gap-2 mb-1">
                  <div className="h-2.5 w-2.5 rounded-full bg-zinc-600" />
                  <span className="text-[11px] text-zinc-500">Completed</span>
                </div>
                <div className="text-2xl font-bold text-zinc-500">{analytics.lifecycle.completed.length}</div>
                <div className="text-[11px] text-zinc-700">All transactions settled</div>
              </div>
            </div>

            {/* ═══ 2. SUMMARY KPIs ═══ */}
            <div className="grid grid-cols-2 md:grid-cols-4 lg:grid-cols-8 gap-3">
              {[
                { label: "Trades", value: String(analytics.pool.length), color: "text-blue-400" },
                { label: "Containers", value: fmt(analytics.financials.totalContainers), color: "text-cyan-400" },
                { label: "Qty (MT)", value: fmt(analytics.financials.totalQuantity), color: "text-zinc-300" },
                { label: "Purchase Value", value: fmtCurrency(analytics.financials.totalPurchaseValue), color: "text-amber-400", small: true },
                { label: "Sales Value", value: fmtCurrency(analytics.financials.totalSalesValue), color: "text-emerald-400", small: true },
                { label: "Gross Margin", value: fmtCurrency(analytics.financials.totalGrossMargin), color: "text-purple-400", small: true },
                { label: "Net Profit", value: fmtCurrency(analytics.financials.totalNetProfit), color: analytics.financials.totalNetProfit >= 0 ? "text-emerald-400" : "text-red-400", small: true },
                { label: "Margin %", value: analytics.financials.totalSalesValue > 0 ? ((analytics.financials.totalNetProfit / analytics.financials.totalSalesValue) * 100).toFixed(1) + "%" : "—", color: analytics.financials.totalNetProfit >= 0 ? "text-emerald-400" : "text-red-400", small: true },
              ].map(({ label, value, color, small }) => (
                <div key={label} className="rounded-lg border border-zinc-800 bg-zinc-900/60 px-4 py-3">
                  <div className={`${small ? "text-lg" : "text-2xl"} font-bold ${color}`}>{value}</div>
                  <div className="text-[11px] text-zinc-500">{label}</div>
                </div>
              ))}
            </div>

            {/* ═══ 3. DEALS REQUIRING ATTENTION ═══ */}
            {analytics.dealsAtPort.length > 0 && (
              <>
                <SectionHeader title="Deals at Port — Requiring Attention" subtitle="Shipment arrived, financial transactions in progress. Open action items." />
                <div className="rounded-lg border border-blue-800/30 bg-blue-950/10 overflow-hidden">
                  <div className="overflow-x-auto max-h-[30vh] overflow-y-auto">
                    <table className="w-full text-xs border-collapse">
                      <thead className="sticky top-0 z-[1]">
                        <tr className="bg-blue-950/40">
                          {["Product", "Trade#", "Origin", "Buyer", "Seller", "Qty (MT)", "Margin", "BL#", "Payable", "Receivable", "Shipment"].map(h => (
                            <th key={h} className="px-2 py-2 text-blue-300/50 border-b border-blue-800/30 font-medium text-left whitespace-nowrap">{h}</th>
                          ))}
                        </tr>
                      </thead>
                      <tbody>
                        {analytics.dealsAtPort.map((t, i) => (
                          <tr key={i} className="border-b border-blue-900/20 hover:bg-blue-900/10">
                            <td className="px-2 py-1.5 text-blue-200">{t.product || "—"}</td>
                            <td className="px-2 py-1.5 text-zinc-400 font-mono">{t.trade_no || "—"}</td>
                            <td className="px-2 py-1.5 text-zinc-300">{t.origin || "—"}</td>
                            <td className="px-2 py-1.5 text-zinc-300 max-w-[120px] truncate">{t.buyer || "—"}</td>
                            <td className="px-2 py-1.5 text-zinc-300 max-w-[120px] truncate">{t.seller || "—"}</td>
                            <td className="px-2 py-1.5 text-right font-mono text-zinc-300">{fmt(t.quantity_mt)}</td>
                            <td className={`px-2 py-1.5 text-right font-mono ${(Number(t.gross_margin) || 0) >= 0 ? "text-emerald-400" : "text-red-400"}`}>
                              {fmtCurrencyFull(Number(t.gross_margin))}
                            </td>
                            <td className="px-2 py-1.5 text-zinc-400 font-mono">{t.bl_number || "—"}</td>
                            <td className={`px-2 py-1.5 text-right font-mono ${t.hasOpenPayable ? "text-red-400" : "text-zinc-600"}`}>
                              {t.hasOpenPayable ? fmtCurrencyFull(Number(t.outward_remaining)) : "Settled"}
                            </td>
                            <td className={`px-2 py-1.5 text-right font-mono ${t.hasOpenReceivable ? "text-amber-400" : "text-zinc-600"}`}>
                              {t.hasOpenReceivable ? fmtCurrencyFull(Number(t.inward_remaining)) : "Collected"}
                            </td>
                            <td className="px-2 py-1.5">
                              <span className="px-1.5 py-0.5 rounded text-[10px] bg-blue-900/40 text-blue-300">{t.shipStage}</span>
                            </td>
                          </tr>
                        ))}
                      </tbody>
                    </table>
                  </div>
                </div>
              </>
            )}

            {/* ═══ 4. SHIPMENT PIPELINE ═══ */}
            {analytics.shipmentPipeline.length > 0 && (
              <>
                <SectionHeader title="Shipment Pipeline" subtitle="All open and at-port trades by logistics stage" />
                <div className="flex gap-2 items-end">
                  {analytics.shipmentPipeline.map(({ stage, count }) => {
                    const colors: Record<string, string> = {
                      "Booking": "bg-zinc-600",
                      "B/L Issued": "bg-amber-600",
                      "Awaiting Shipment": "bg-orange-500",
                      "In Transit": "bg-blue-500",
                      "Arrived": "bg-emerald-500",
                    };
                    const maxCount = Math.max(...analytics.shipmentPipeline.map(s => s.count), 1);
                    const h = Math.max(32, (count / maxCount) * 120);
                    return (
                      <div key={stage} className="flex-1 flex flex-col items-center gap-1.5">
                        <span className="text-lg font-bold text-zinc-200">{count}</span>
                        <div className={`w-full rounded-t-md ${colors[stage] || "bg-zinc-600"}`} style={{ height: `${h}px` }} />
                        <span className="text-[10px] text-zinc-500 text-center leading-tight">{stage}</span>
                      </div>
                    );
                  })}
                </div>
              </>
            )}

            {/* ═══ 5. CASH FLOW & EXPENSES ═══ */}
            <SectionHeader title="Cash Flow & Expenses" />
            <div className="grid grid-cols-1 md:grid-cols-2 gap-4">
              <div className="rounded-lg border border-zinc-800 bg-zinc-900/30 p-4">
                <h3 className="text-xs font-medium text-zinc-400 mb-3">Cash Flow Summary</h3>
                <div className="space-y-2.5">
                  <div className="flex justify-between text-xs">
                    <span className="text-zinc-400">Total Paid to Sellers</span>
                    <span className="text-red-400 font-mono">{fmtCurrencyFull(analytics.cashFlow.totalOutwards)}</span>
                  </div>
                  <div className="flex justify-between text-xs">
                    <span className="text-zinc-400">Total Received from Buyers</span>
                    <span className="text-emerald-400 font-mono">{fmtCurrencyFull(analytics.cashFlow.totalInwards)}</span>
                  </div>
                  <div className="border-t border-zinc-800 my-1" />
                  <div className="flex justify-between text-xs">
                    <span className="text-zinc-400">Still Owed to Sellers</span>
                    <span className="text-amber-400 font-mono">{fmtCurrencyFull(analytics.cashFlow.outwardRemaining)}</span>
                  </div>
                  <div className="flex justify-between text-xs">
                    <span className="text-zinc-400">Still Owed by Buyers</span>
                    <span className="text-amber-400 font-mono">{fmtCurrencyFull(analytics.cashFlow.inwardRemaining)}</span>
                  </div>
                  <div className="border-t border-zinc-800 my-1" />
                  <div className="flex justify-between text-xs font-medium">
                    <span className="text-zinc-300">Net Cash Position</span>
                    <span className={`font-mono ${(analytics.cashFlow.totalInwards - analytics.cashFlow.totalOutwards) >= 0 ? "text-emerald-400" : "text-red-400"}`}>
                      {fmtCurrencyFull(analytics.cashFlow.totalInwards - analytics.cashFlow.totalOutwards)}
                    </span>
                  </div>
                  {analytics.cashFlow.avgWorkingCapitalDays > 0 && (
                    <div className="flex justify-between text-xs">
                      <span className="text-zinc-400">Avg Working Capital Days</span>
                      <span className="text-zinc-300 font-mono">{Math.round(analytics.cashFlow.avgWorkingCapitalDays)}d</span>
                    </div>
                  )}
                </div>
              </div>
              <div className="rounded-lg border border-zinc-800 bg-zinc-900/30 p-4">
                <h3 className="text-xs font-medium text-zinc-400 mb-3">Expense Breakdown</h3>
                <div className="space-y-2.5">
                  {[
                    { label: "Clearance Charges", val: analytics.expenses.totalClearanceCharges },
                    { label: "Brokerage", val: analytics.expenses.totalBrokerage },
                    { label: "Warehouse Loss", val: analytics.expenses.totalWarehouseLoss },
                    { label: "Claims Paid", val: analytics.expenses.totalClaimsPaid },
                    { label: "Claims Received", val: analytics.expenses.totalClaimsReceived },
                    { label: "Interest Loss", val: analytics.expenses.totalInterestLoss },
                  ].filter(e => e.val !== 0).map(({ label, val }) => (
                    <div key={label} className="flex justify-between text-xs">
                      <span className="text-zinc-400">{label}</span>
                      <span className={`font-mono ${val > 0 ? "text-red-400" : "text-emerald-400"}`}>{fmtCurrencyFull(Math.abs(val))}</span>
                    </div>
                  ))}
                  <div className="border-t border-zinc-800 my-1" />
                  <div className="flex justify-between text-xs font-medium">
                    <span className="text-zinc-300">Total Expenses</span>
                    <span className="text-red-400 font-mono">{fmtCurrencyFull(analytics.expenses.totalExpenses)}</span>
                  </div>
                </div>
              </div>
            </div>

            {/* ═══ 6. OPEN RECEIVABLES & PAYABLES ═══ */}
            {(analytics.openReceivables.length > 0 || analytics.openPayables.length > 0) && (
              <>
                <SectionHeader title="Open Receivables & Payables" subtitle="Outstanding amounts — who owes us and who we owe" />
                <div className="grid grid-cols-1 md:grid-cols-2 gap-4">
                  {analytics.openReceivables.length > 0 && (
                    <div className="rounded-lg border border-zinc-800 bg-zinc-900/30 p-4">
                      <h3 className="text-xs font-medium text-amber-400/80 mb-3">
                        Open Receivables — {fmtCurrencyFull(analytics.openReceivables.reduce((s, r) => s + r.amount, 0))} total
                      </h3>
                      <div className="space-y-1.5 max-h-[200px] overflow-y-auto">
                        {analytics.openReceivables.slice(0, 15).map((r, i) => (
                          <div key={i} className="flex items-center justify-between text-xs">
                            <div className="flex items-center gap-2 min-w-0">
                              <span className="text-zinc-300 truncate max-w-[120px]">{r.buyer}</span>
                              <span className="text-zinc-600">·</span>
                              <span className="text-zinc-500 truncate max-w-[80px]">{r.product}</span>
                            </div>
                            <span className="text-amber-400 font-mono shrink-0">{fmtCurrencyFull(r.amount)}</span>
                          </div>
                        ))}
                      </div>
                    </div>
                  )}
                  {analytics.openPayables.length > 0 && (
                    <div className="rounded-lg border border-zinc-800 bg-zinc-900/30 p-4">
                      <h3 className="text-xs font-medium text-red-400/80 mb-3">
                        Open Payables — {fmtCurrencyFull(analytics.openPayables.reduce((s, p) => s + p.amount, 0))} total
                      </h3>
                      <div className="space-y-1.5 max-h-[200px] overflow-y-auto">
                        {analytics.openPayables.slice(0, 15).map((p, i) => (
                          <div key={i} className="flex items-center justify-between text-xs">
                            <div className="flex items-center gap-2 min-w-0">
                              <span className="text-zinc-300 truncate max-w-[120px]">{p.seller}</span>
                              <span className="text-zinc-600">·</span>
                              <span className="text-zinc-500 truncate max-w-[80px]">{p.product}</span>
                            </div>
                            <span className="text-red-400 font-mono shrink-0">{fmtCurrencyFull(p.amount)}</span>
                          </div>
                        ))}
                      </div>
                    </div>
                  )}
                </div>
              </>
            )}

            {/* ═══ 7. POSITION BOOK ═══ */}
            <SectionHeader title="Position Book" subtitle={`${displayRows.length} trades`} />
            {displayRows.length > 0 ? (
              <div className="rounded-lg border border-zinc-800 bg-zinc-900/30 overflow-hidden">
                <div className="overflow-x-auto max-h-[55vh] overflow-y-auto">
                  <table className="w-full text-xs border-collapse">
                    <thead className="sticky top-0 z-[1]">
                      <tr>
                        <th className="px-2 py-2 bg-zinc-800 text-zinc-500 border-b border-zinc-700 text-left font-medium w-8">#</th>
                        {columns.map((c) => (
                          <th key={c.key} className={`px-2 py-2 bg-zinc-800 text-zinc-500 border-b border-zinc-700 font-medium whitespace-nowrap ${c.align === "right" ? "text-right" : "text-left"}`}>
                            {c.label}
                          </th>
                        ))}
                      </tr>
                    </thead>
                    <tbody>
                      {displayRows.map((row, ri) => {
                        const stage = positionStage(row.position ?? "");
                        const rowClass = stage === "completed"
                          ? "opacity-35 hover:opacity-60"
                          : stage === "at_port"
                          ? "bg-blue-950/15 hover:bg-blue-950/25 border-b border-blue-900/20"
                          : "hover:bg-zinc-800/40 border-b border-zinc-800/30";
                        return (
                          <tr key={row.id || ri} className={`transition-opacity ${rowClass}`}>
                            <td className="px-2 py-1.5 text-zinc-600 font-mono">{ri + 1}</td>
                            {columns.map((c) => {
                              const v = row[c.key];
                              const s = v !== null && v !== undefined ? String(v) : "—";
                              const isPositionCol = c.key === "position";
                              const isProfit = (c.key === "net_profit" || c.key === "gross_margin") && typeof v === "number";
                              const posColor = isPositionCol
                                ? stage === "open" ? "text-amber-400 font-medium"
                                : stage === "at_port" ? "text-blue-400 font-medium"
                                : "text-zinc-500"
                                : "";
                              return (
                                <td key={c.key}
                                  className={`px-2 py-1.5 max-w-[140px] truncate ${c.align === "right" ? "text-right font-mono" : ""} ${
                                    isPositionCol ? posColor :
                                    isProfit && (v as number) > 0 ? "text-emerald-400" :
                                    isProfit && (v as number) < 0 ? "text-red-400" :
                                    "text-zinc-300"
                                  }`}>
                                  {c.align === "right" && typeof v === "number" ? fmt(v) : s}
                                </td>
                              );
                            })}
                          </tr>
                        );
                      })}
                    </tbody>
                  </table>
                </div>
              </div>
            ) : (
              <div className="rounded-lg border border-zinc-800 bg-zinc-900/30 p-6 text-center text-sm text-zinc-500">
                No trades found for this filter
              </div>
            )}

            {/* ═══ 8. P&L HEATMAP (Commodity × Origin by Profit) ═══ */}
            {analytics.profitMatrix.products.length > 0 && analytics.profitMatrix.origins.length > 0 && (
              <>
                <SectionHeader title="P&L Heatmap — Commodity × Origin" subtitle="Net profit by trade corridor. Green = profitable, Red = losing" />
                <div className="rounded-lg border border-zinc-800 bg-zinc-900/30 p-4 overflow-x-auto">
                  <table className="text-xs border-collapse">
                    <thead>
                      <tr>
                        <th className="px-3 py-2 text-left text-zinc-500 font-medium" />
                        {analytics.profitMatrix.origins.map((o) => (
                          <th key={o} className="px-3 py-2 text-zinc-400 font-medium whitespace-nowrap">{o}</th>
                        ))}
                      </tr>
                    </thead>
                    <tbody>
                      {analytics.profitMatrix.products.map((p, pi) => (
                        <tr key={p}>
                          <td className="px-3 py-2 text-zinc-300 font-medium whitespace-nowrap">{p}</td>
                          {analytics.profitMatrix.values[pi].map((cell, oi) => {
                            if (cell.count === 0) {
                              return <td key={oi} className="px-3 py-2 text-center text-zinc-700">·</td>;
                            }
                            const isPositive = cell.profit >= 0;
                            const maxProfit = Math.max(...analytics.profitMatrix.values.flat().map(c => Math.abs(c.profit)), 1);
                            const intensity = Math.min(Math.abs(cell.profit) / maxProfit, 1);
                            return (
                              <td key={oi} className="px-3 py-2 text-center"
                                style={{
                                  backgroundColor: isPositive
                                    ? `rgba(16, 185, 129, ${0.08 + intensity * 0.35})`
                                    : `rgba(239, 68, 68, ${0.08 + intensity * 0.35})`,
                                }}>
                                <div className={`font-mono font-medium ${isPositive ? "text-emerald-300" : "text-red-300"}`}>
                                  {fmtCurrency(cell.profit)}
                                </div>
                                <div className="text-[10px] text-zinc-500">{cell.count} trades</div>
                              </td>
                            );
                          })}
                        </tr>
                      ))}
                    </tbody>
                  </table>
                </div>
              </>
            )}

            {/* ═══ 9. P&L BY COMMODITY ═══ */}
            {analytics.profitByProduct.length > 0 && (
              <>
                <SectionHeader title="P&L by Commodity" />
                <div className="rounded-lg border border-zinc-800 bg-zinc-900/30 p-4">
                  <div className="space-y-1.5">
                    {analytics.profitByProduct.map(([name, profit]) => {
                      const maxVal = Math.max(...analytics.profitByProduct.map(([, v]) => Math.abs(v)), 1);
                      const w = Math.max(6, (Math.abs(profit) / maxVal) * 100);
                      return (
                        <div key={name} className="flex items-center gap-2 text-xs">
                          <span className="w-28 truncate text-zinc-300 text-right shrink-0">{name}</span>
                          <div className="flex-1 h-4 bg-zinc-800 rounded-sm overflow-hidden">
                            <div className={`h-full rounded-sm ${profit >= 0 ? "bg-emerald-500" : "bg-red-500"}`} style={{ width: `${w}%` }} />
                          </div>
                          <span className={`w-20 text-right font-mono shrink-0 text-[11px] ${profit >= 0 ? "text-emerald-400" : "text-red-400"}`}>
                            {fmtCurrency(profit)}
                          </span>
                        </div>
                      );
                    })}
                  </div>
                </div>
              </>
            )}

            {/* ═══ 10. EXPOSURE BY VALUE ═══ */}
            <SectionHeader title="Exposure Analysis — by Value" subtitle="Dollar-weighted exposure to counterparties, origins, and commodities" />
            <div className="grid grid-cols-1 md:grid-cols-2 gap-4">
              <div className="rounded-lg border border-zinc-800 bg-zinc-900/30 p-4">
                <h3 className="text-xs font-medium text-zinc-400 mb-3">Buyer Exposure (by Sales Value)</h3>
                <div className="space-y-1.5">
                  {analytics.buyerExposureByValue.map(([name, val]) => (
                    <Bar key={name} label={name} value={val} max={analytics.buyerExposureByValue[0]?.[1] || 1} color="bg-cyan-500" suffix="$" />
                  ))}
                </div>
              </div>
              <div className="rounded-lg border border-zinc-800 bg-zinc-900/30 p-4">
                <h3 className="text-xs font-medium text-zinc-400 mb-3">Seller Exposure (by Purchase Value)</h3>
                <div className="space-y-1.5">
                  {analytics.sellerExposureByValue.map(([name, val]) => (
                    <Bar key={name} label={name} value={val} max={analytics.sellerExposureByValue[0]?.[1] || 1} color="bg-purple-500" suffix="$" />
                  ))}
                </div>
              </div>
              <div className="rounded-lg border border-zinc-800 bg-zinc-900/30 p-4">
                <h3 className="text-xs font-medium text-zinc-400 mb-3">Commodity Exposure (by Purchase Value)</h3>
                <div className="space-y-1.5">
                  {analytics.productExposureByValue.map(([name, val]) => (
                    <Bar key={name} label={name} value={val} max={analytics.productExposureByValue[0]?.[1] || 1} color="bg-blue-500" suffix="$" />
                  ))}
                </div>
              </div>
              <div className="rounded-lg border border-zinc-800 bg-zinc-900/30 p-4">
                <h3 className="text-xs font-medium text-zinc-400 mb-3">Origin Exposure (by Purchase Value)</h3>
                <div className="space-y-1.5">
                  {analytics.originExposureByValue.map(([name, val]) => (
                    <Bar key={name} label={name} value={val} max={analytics.originExposureByValue[0]?.[1] || 1} color="bg-amber-500" suffix="$" />
                  ))}
                </div>
              </div>
            </div>

            {/* ═══ 11. DELIVERY SCHEDULE & WORKING CAPITAL ═══ */}
            <div className="grid grid-cols-1 md:grid-cols-2 gap-4">
              <div className="rounded-lg border border-zinc-800 bg-zinc-900/30 p-4">
                <h3 className="text-xs font-medium text-zinc-400 mb-3">Delivery Schedule</h3>
                <div className="space-y-1.5">
                  {analytics.byMonth.map(([name, count]) => (
                    <Bar key={name} label={name} value={count} max={Math.max(...analytics.byMonth.map(([, c]) => c), 1)} color="bg-emerald-500" />
                  ))}
                </div>
              </div>
              {analytics.wcBuckets.length > 0 && (
                <div className="rounded-lg border border-zinc-800 bg-zinc-900/30 p-4">
                  <h3 className="text-xs font-medium text-zinc-400 mb-3">
                    Working Capital Distribution
                    {analytics.cashFlow.avgWorkingCapitalDays > 0 && (
                      <span className="ml-2 text-zinc-500">(avg: {Math.round(analytics.cashFlow.avgWorkingCapitalDays)}d)</span>
                    )}
                  </h3>
                  <div className="space-y-1.5">
                    {analytics.wcBuckets.map(({ label, count }) => (
                      <Bar key={label} label={label} value={count} max={Math.max(...analytics.wcBuckets.map(b => b.count), 1)} color="bg-orange-500" />
                    ))}
                  </div>
                </div>
              )}
            </div>

            {/* ═══ 12. CONCENTRATION RISK ═══ */}
            <div className="rounded-lg border border-zinc-800 bg-zinc-900/30 p-4">
              <h3 className="text-xs font-medium text-zinc-400 mb-3">Concentration Risk</h3>
              <div className="grid grid-cols-1 md:grid-cols-3 gap-6">
                {[
                  { label: "Top Supplier", data: analytics.risk.packer },
                  { label: "Top Origin", data: analytics.risk.origin },
                  { label: "Top Commodity", data: analytics.risk.product },
                ].map(({ label, data: d }) =>
                  d && (
                    <div key={label}>
                      <div className="flex items-baseline justify-between mb-1">
                        <span className="text-xs text-zinc-300">{label}</span>
                        <span className={`text-xs font-medium ${riskColor(d.pct)}`}>
                          {d.pct}% — {riskLabel(d.pct)}
                        </span>
                      </div>
                      <div className="h-2 bg-zinc-800 rounded-full overflow-hidden">
                        <div
                          className={`h-full rounded-full ${d.pct >= 40 ? "bg-red-500" : d.pct >= 25 ? "bg-amber-500" : "bg-emerald-500"}`}
                          style={{ width: `${d.pct}%` }}
                        />
                      </div>
                      <div className="text-[11px] text-zinc-500 mt-0.5">{d.name}</div>
                    </div>
                  )
                )}
              </div>
            </div>
          </>
        )}

        <div className="text-center text-[11px] text-zinc-700 pt-2 pb-4">
          Hectar CTRM · Supabase + OneDrive · Auto-sync every 5 min
        </div>
      </main>
    </div>
  );
}
