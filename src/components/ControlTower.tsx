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

// ─── Bar component ────────────────────────────────────────────────────────────

function Bar({ label, count, max, color }: { label: string; count: number; max: number; color: string }) {
  const w = max > 0 ? Math.max(6, (count / max) * 100) : 0;
  return (
    <div className="flex items-center gap-2 text-xs">
      <span className="w-28 truncate text-zinc-300 text-right shrink-0">{label}</span>
      <div className="flex-1 h-4 bg-zinc-800 rounded-sm overflow-hidden">
        <div className={`h-full ${color} rounded-sm`} style={{ width: `${w}%` }} />
      </div>
      <span className="w-8 text-zinc-500 text-right font-mono shrink-0">{count}</span>
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
  const [showActive, setShowActive] = useState(true);
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

  // ─── Load on mount ──────────────────────────────────────────

  useEffect(() => {
    fetchTrades();
    fetchSyncStatus();
  }, []); // eslint-disable-line react-hooks/exhaustive-deps

  // Auto-refresh data every 60s
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

    const active = filtered.filter(
      (t) => (t.position ?? "").toLowerCase().trim() === "active"
    );
    const inactive = filtered.length - active.length;
    const pool = active.length > 0 ? active : filtered;

    const byProduct = countBy(pool, "product");
    const byOrigin = countBy(pool, "origin");
    const byMonth = countBy(pool, "month");
    const byPacker = countBy(pool, "packer");
    const byBuyer = countBy(pool, "buyer");

    const sheets = [...new Set(trades.map((t) => t.source_sheet))].sort();

    const monthSorted: [string, number][] = [...byMonth.entries()].sort(
      (a, b) => monthIndex(a[0]) - monthIndex(b[0])
    );

    const topProduct = topN(byProduct, 1)[0];
    const topOrigin = topN(byOrigin, 1)[0];
    const topPacker = topN(byPacker, 1)[0];
    const total = pool.length || 1;

    // Commodity × Origin matrix
    const topProds = topN(byProduct, 6).map(([n]) => n);
    const topOrigs = topN(byOrigin, 6).map(([n]) => n);
    const matrix = topProds.map((p) =>
      topOrigs.map((o) =>
        pool.filter((t) => t.product === p && t.origin === o).length
      )
    );

    // Financial summaries
    const totalPurchaseValue = sumBy(pool, "purchase_value");
    const totalSalesValue = sumBy(pool, "sales_value");
    const totalGrossMargin = sumBy(pool, "gross_margin");
    const totalNetProfit = sumBy(pool, "net_profit");
    const totalContainers = sumBy(pool, "no_of_containers");
    const totalQuantity = sumBy(pool, "quantity_mt");

    return {
      active, inactive, filtered, sheets, pool,
      byProduct: topN(byProduct, 10),
      byOrigin: topN(byOrigin, 10),
      byMonth: monthSorted,
      byPacker: topN(byPacker, 10),
      byBuyer: topN(byBuyer, 10),
      risk: {
        product: topProduct ? { name: topProduct[0], pct: pct(topProduct[1], total) } : null,
        origin: topOrigin ? { name: topOrigin[0], pct: pct(topOrigin[1], total) } : null,
        packer: topPacker ? { name: topPacker[0], pct: pct(topPacker[1], total) } : null,
      },
      matrix: { products: topProds, origins: topOrigs, values: matrix },
      financials: { totalPurchaseValue, totalSalesValue, totalGrossMargin, totalNetProfit, totalContainers, totalQuantity },
    };
  }, [trades, sheetFilter]);

  const displayRows = analytics
    ? showActive ? analytics.active : analytics.filtered
    : [];

  // ─── Visible columns (curated, not all 85) ─────────────────

  const columns: { key: keyof Trade; label: string; align?: "right" }[] = [
    { key: "product", label: "Product" },
    { key: "position", label: "Position" },
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
        <div className="max-w-[1400px] mx-auto px-6 py-3 flex items-center justify-between">
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
              {syncing ? "Syncing..." : "Sync Now"}
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

      <main className="max-w-[1400px] mx-auto px-6 py-5 space-y-5">
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

        {/* ─── Empty state: no data, need first sync ─────────── */}
        {!loading && trades.length === 0 && !error && (
          <div className="max-w-md mx-auto py-16 text-center space-y-5">
            <div className="inline-flex h-16 w-16 items-center justify-center rounded-full bg-zinc-900 border border-zinc-800">
              <svg className="h-8 w-8 text-zinc-400" fill="none" viewBox="0 0 24 24" stroke="currentColor" strokeWidth={1.5}>
                <path strokeLinecap="round" strokeLinejoin="round" d="M3 16.5v2.25A2.25 2.25 0 005.25 21h13.5A2.25 2.25 0 0021 18.75V16.5m-13.5-9L12 3m0 0l4.5 4.5M12 3v13.5" />
              </svg>
            </div>
            <div>
              <h2 className="text-lg font-semibold text-zinc-100 mb-1">Ready to sync</h2>
              <p className="text-sm text-zinc-400">
                Pull the latest data from your Excel spreadsheet into the database.
              </p>
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

        {/* ─── Error ──────────────────────────────────────────── */}
        {error && (
          <div className="p-3 bg-red-900/20 border border-red-800 rounded-lg text-xs text-red-400">
            {error}
          </div>
        )}

        {/* ═══ Dashboard (only when data exists) ═══ */}
        {analytics && (
          <>
            {/* ─── Sheet filter ──────────────────────────────── */}
            {analytics.sheets.length > 1 && (
              <div className="flex items-center gap-2">
                <span className="text-[11px] text-zinc-500">Sheet:</span>
                <button
                  onClick={() => setSheetFilter("all")}
                  className={`px-2.5 py-1 text-[11px] rounded-full border transition-colors ${
                    sheetFilter === "all"
                      ? "border-blue-700 bg-blue-900/30 text-blue-400"
                      : "border-zinc-700 bg-zinc-800/50 text-zinc-400 hover:border-zinc-600"
                  }`}
                >
                  All
                </button>
                {analytics.sheets.map((s) => (
                  <button
                    key={s}
                    onClick={() => setSheetFilter(s)}
                    className={`px-2.5 py-1 text-[11px] rounded-full border transition-colors ${
                      sheetFilter === s
                        ? "border-blue-700 bg-blue-900/30 text-blue-400"
                        : "border-zinc-700 bg-zinc-800/50 text-zinc-400 hover:border-zinc-600"
                    }`}
                  >
                    {s.trim()}
                  </button>
                ))}
              </div>
            )}

            {/* ═══ Summary Metrics ═══ */}
            <div className="grid grid-cols-2 md:grid-cols-4 lg:grid-cols-8 gap-3">
              <div className="rounded-lg border border-zinc-800 bg-zinc-900/60 px-4 py-3">
                <div className="text-2xl font-bold text-emerald-400">{analytics.active.length}</div>
                <div className="text-[11px] text-zinc-500">Active</div>
              </div>
              <div className="rounded-lg border border-zinc-800 bg-zinc-900/60 px-4 py-3">
                <div className="text-2xl font-bold text-zinc-500">{analytics.inactive}</div>
                <div className="text-[11px] text-zinc-500">Inactive</div>
              </div>
              <div className="rounded-lg border border-zinc-800 bg-zinc-900/60 px-4 py-3">
                <div className="text-2xl font-bold text-blue-400">{fmt(analytics.financials.totalContainers)}</div>
                <div className="text-[11px] text-zinc-500">Containers</div>
              </div>
              <div className="rounded-lg border border-zinc-800 bg-zinc-900/60 px-4 py-3">
                <div className="text-2xl font-bold text-cyan-400">{fmt(analytics.financials.totalQuantity)}</div>
                <div className="text-[11px] text-zinc-500">Quantity (MT)</div>
              </div>
              <div className="rounded-lg border border-zinc-800 bg-zinc-900/60 px-4 py-3">
                <div className="text-lg font-bold text-amber-400">{fmtCurrency(analytics.financials.totalPurchaseValue)}</div>
                <div className="text-[11px] text-zinc-500">Purchase Value</div>
              </div>
              <div className="rounded-lg border border-zinc-800 bg-zinc-900/60 px-4 py-3">
                <div className="text-lg font-bold text-emerald-400">{fmtCurrency(analytics.financials.totalSalesValue)}</div>
                <div className="text-[11px] text-zinc-500">Sales Value</div>
              </div>
              <div className="rounded-lg border border-zinc-800 bg-zinc-900/60 px-4 py-3">
                <div className="text-lg font-bold text-purple-400">{fmtCurrency(analytics.financials.totalGrossMargin)}</div>
                <div className="text-[11px] text-zinc-500">Gross Margin</div>
              </div>
              <div className="rounded-lg border border-zinc-800 bg-zinc-900/60 px-4 py-3">
                <div className={`text-lg font-bold ${analytics.financials.totalNetProfit >= 0 ? "text-emerald-400" : "text-red-400"}`}>
                  {fmtCurrency(analytics.financials.totalNetProfit)}
                </div>
                <div className="text-[11px] text-zinc-500">Net Profit</div>
              </div>
            </div>

            {/* ═══ Position Book ═══ */}
            <div>
              <div className="flex items-center justify-between mb-2">
                <div className="flex items-center gap-3">
                  <h2 className="text-sm font-semibold text-zinc-200">Position Book</h2>
                  <span className="text-xs text-zinc-500">{displayRows.length} rows</span>
                  <button
                    onClick={() => setShowActive(!showActive)}
                    className={`px-2 py-0.5 text-[11px] rounded-full border transition-colors ${
                      showActive
                        ? "border-emerald-800 bg-emerald-900/30 text-emerald-400"
                        : "border-zinc-700 bg-zinc-800 text-zinc-400"
                    }`}
                  >
                    {showActive ? "Active only" : "All"}
                  </button>
                </div>
              </div>
              {displayRows.length > 0 ? (
                <div className="rounded-lg border border-zinc-800 bg-zinc-900/30 overflow-hidden">
                  <div className="overflow-x-auto max-h-[55vh] overflow-y-auto">
                    <table className="w-full text-xs border-collapse">
                      <thead className="sticky top-0 z-[1]">
                        <tr>
                          <th className="px-2 py-2 bg-zinc-800 text-zinc-500 border-b border-zinc-700 text-left font-medium w-8">#</th>
                          {columns.map((c) => (
                            <th
                              key={c.key}
                              className={`px-2 py-2 bg-zinc-800 text-zinc-500 border-b border-zinc-700 font-medium whitespace-nowrap ${c.align === "right" ? "text-right" : "text-left"}`}
                            >
                              {c.label}
                            </th>
                          ))}
                        </tr>
                      </thead>
                      <tbody>
                        {displayRows.map((row, ri) => {
                          const posVal = (row.position ?? "").toLowerCase().trim();
                          return (
                            <tr key={row.id || ri} className={`border-b border-zinc-800/30 hover:bg-zinc-800/40 ${posVal === "inactive" ? "opacity-30" : ""}`}>
                              <td className="px-2 py-1.5 text-zinc-600 font-mono">{ri + 1}</td>
                              {columns.map((c) => {
                                const v = row[c.key];
                                const s = v !== null && v !== undefined ? String(v) : "—";
                                const isActive = c.key === "position" && posVal === "active";
                                const isProfit = (c.key === "net_profit" || c.key === "gross_margin") && typeof v === "number";
                                return (
                                  <td
                                    key={c.key}
                                    className={`px-2 py-1.5 max-w-[140px] truncate ${
                                      c.align === "right" ? "text-right font-mono" : ""
                                    } ${
                                      isActive ? "text-emerald-400 font-medium" :
                                      isProfit && v !== null && (v as number) > 0 ? "text-emerald-400" :
                                      isProfit && v !== null && (v as number) < 0 ? "text-red-400" :
                                      "text-zinc-300"
                                    }`}
                                  >
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
                  No {showActive ? "active positions" : "data"} found
                </div>
              )}
            </div>

            {/* ═══ Exposure Analysis ═══ */}
            <h2 className="text-sm font-semibold text-zinc-200 pt-2">Exposure Analysis</h2>
            <div className="grid grid-cols-1 md:grid-cols-3 gap-4">
              <div className="rounded-lg border border-zinc-800 bg-zinc-900/30 p-4">
                <h3 className="text-xs font-medium text-zinc-400 mb-3">By Commodity</h3>
                <div className="space-y-1.5">
                  {analytics.byProduct.map(([name, count]) => (
                    <Bar key={name} label={name} count={count} max={analytics.byProduct[0]?.[1] || 1} color="bg-blue-500" />
                  ))}
                </div>
              </div>
              <div className="rounded-lg border border-zinc-800 bg-zinc-900/30 p-4">
                <h3 className="text-xs font-medium text-zinc-400 mb-3">By Origin</h3>
                <div className="space-y-1.5">
                  {analytics.byOrigin.map(([name, count]) => (
                    <Bar key={name} label={name} count={count} max={analytics.byOrigin[0]?.[1] || 1} color="bg-amber-500" />
                  ))}
                </div>
              </div>
              <div className="rounded-lg border border-zinc-800 bg-zinc-900/30 p-4">
                <h3 className="text-xs font-medium text-zinc-400 mb-3">Delivery Schedule</h3>
                <div className="space-y-1.5">
                  {analytics.byMonth.map(([name, count]) => (
                    <Bar key={name} label={name} count={count} max={Math.max(...analytics.byMonth.map(([, c]) => c), 1)} color="bg-emerald-500" />
                  ))}
                </div>
              </div>
            </div>

            {/* ═══ Counterparty & Risk ═══ */}
            <div className="grid grid-cols-1 md:grid-cols-2 gap-4">
              <div className="rounded-lg border border-zinc-800 bg-zinc-900/30 p-4">
                <h3 className="text-xs font-medium text-zinc-400 mb-3">Counterparty Exposure (Sellers)</h3>
                <div className="space-y-1.5">
                  {analytics.byPacker.map(([name, count]) => (
                    <Bar key={name} label={name} count={count} max={analytics.byPacker[0]?.[1] || 1} color="bg-purple-500" />
                  ))}
                </div>
              </div>
              <div className="rounded-lg border border-zinc-800 bg-zinc-900/30 p-4">
                <h3 className="text-xs font-medium text-zinc-400 mb-3">Concentration Risk</h3>
                <div className="space-y-4">
                  {[
                    { label: "Top Packer", data: analytics.risk.packer },
                    { label: "Top Origin", data: analytics.risk.origin },
                    { label: "Top Commodity", data: analytics.risk.product },
                  ].map(
                    ({ label, data: d }) =>
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
            </div>

            {/* ═══ Buyer Exposure ═══ */}
            {analytics.byBuyer.length > 0 && (
              <div className="rounded-lg border border-zinc-800 bg-zinc-900/30 p-4">
                <h3 className="text-xs font-medium text-zinc-400 mb-3">Buyer Exposure</h3>
                <div className="grid grid-cols-1 md:grid-cols-2 gap-x-8 gap-y-1.5">
                  {analytics.byBuyer.map(([name, count]) => (
                    <Bar key={name} label={name} count={count} max={analytics.byBuyer[0]?.[1] || 1} color="bg-cyan-500" />
                  ))}
                </div>
              </div>
            )}

            {/* ═══ Commodity × Origin Matrix ═══ */}
            {analytics.matrix.products.length > 0 && analytics.matrix.origins.length > 0 && (
              <div className="rounded-lg border border-zinc-800 bg-zinc-900/30 p-4">
                <h3 className="text-xs font-medium text-zinc-400 mb-3">Commodity × Origin Matrix</h3>
                <div className="overflow-x-auto">
                  <table className="text-xs border-collapse">
                    <thead>
                      <tr>
                        <th className="px-2 py-1.5 text-left text-zinc-500 font-medium" />
                        {analytics.matrix.origins.map((o) => (
                          <th key={o} className="px-3 py-1.5 text-zinc-400 font-medium whitespace-nowrap">{o}</th>
                        ))}
                      </tr>
                    </thead>
                    <tbody>
                      {analytics.matrix.products.map((p, pi) => (
                        <tr key={p}>
                          <td className="px-2 py-1.5 text-zinc-300 font-medium whitespace-nowrap">{p}</td>
                          {analytics.matrix.values[pi].map((v, oi) => {
                            const maxVal = Math.max(...analytics.matrix.values.flat(), 1);
                            const intensity = v / maxVal;
                            return (
                              <td
                                key={oi}
                                className="px-3 py-1.5 text-center font-mono"
                                style={{
                                  backgroundColor: v > 0 ? `rgba(59, 130, 246, ${0.1 + intensity * 0.5})` : "transparent",
                                  color: v > 0 ? "#93c5fd" : "#3f3f46",
                                }}
                              >
                                {v || "·"}
                              </td>
                            );
                          })}
                        </tr>
                      ))}
                    </tbody>
                  </table>
                </div>
              </div>
            )}
          </>
        )}

        <div className="text-center text-[11px] text-zinc-700 pt-2 pb-4">
          Hectar CTRM · Supabase + OneDrive · Auto-refresh 60s
        </div>
      </main>
    </div>
  );
}
