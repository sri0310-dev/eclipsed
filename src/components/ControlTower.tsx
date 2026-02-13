"use client";

import { useState, useEffect, useCallback, useRef, useMemo } from "react";

// ─── Types ────────────────────────────────────────────────────────────────────

interface SheetData {
  headers: string[];
  rows: Record<string, string | number | boolean | null>[];
  rawValues: (string | number | boolean | null)[][];
  range: string;
}

interface LogEntry {
  ts: string;
  type: "load" | "add" | "delta" | "error";
  msg: string;
}

// ─── Helpers ──────────────────────────────────────────────────────────────────

function now(): string {
  return new Date().toLocaleString("en-GB", {
    day: "2-digit",
    month: "short",
    hour: "2-digit",
    minute: "2-digit",
    second: "2-digit",
  });
}

function fingerprint(data: SheetData): string {
  const tail = data.rawValues.slice(-3);
  return `${data.rawValues.length}|${JSON.stringify(tail)}`;
}

function parseRange(range: string): { lastRow: number; lastCol: string } {
  const m = range.match(/:([A-Z]+)(\d+)$/);
  return m ? { lastCol: m[1], lastRow: parseInt(m[2]) } : { lastCol: "A", lastRow: 1 };
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

function col(headers: string[], name: string): string | null {
  return headers.find((h) => h.toLowerCase().trim() === name.toLowerCase()) || null;
}

function countBy(
  rows: Record<string, string | number | boolean | null>[],
  colName: string | null
): Map<string, number> {
  const map = new Map<string, number>();
  if (!colName) return map;
  for (const r of rows) {
    const v = String(r[colName] ?? "").trim();
    if (v) map.set(v, (map.get(v) || 0) + 1);
  }
  return map;
}

function topN(map: Map<string, number>, n: number): [string, number][] {
  return [...map.entries()].sort((a, b) => b[1] - a[1]).slice(0, n);
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

// Rank columns by trading relevance
function rankColumns(headers: string[]): string[] {
  const tags: Record<string, string[]> = {
    id: ["product", "commodity code", "commodity", "code", "origin", "variety", "packer",
      "counterparty", "supplier", "buyer", "contract", "deal", "reference", "ref"],
    status: ["position", "status", "state"],
    time: ["month", "delivery", "shipment", "eta", "etd", "date"],
    qty: ["quantity", "qty", "volume", "weight", "mt", "tons", "bags", "container", "20s", "40s"],
    price: ["price", "rate", "value", "amount", "cost", "fob", "cif", "cfr", "premium",
      "margin", "pnl", "p&l", "usd", "total", "invoice"],
  };
  const order = ["id", "status", "time", "qty", "price"];
  const buckets: Map<string, string[]> = new Map(order.map((k) => [k, []]));
  buckets.set("other", []);

  for (const h of headers) {
    const l = h.toLowerCase().trim();
    let placed = false;
    for (const cat of order) {
      if (tags[cat].some((k) => l.includes(k))) {
        buckets.get(cat)!.push(h);
        placed = true;
        break;
      }
    }
    if (!placed) buckets.get("other")!.push(h);
  }
  return [...order, "other"].flatMap((k) => buckets.get(k) || []);
}

// ─── Bar component (reused in all breakdowns) ─────────────────────────────────

function Bar({
  label,
  count,
  max,
  color,
}: {
  label: string;
  count: number;
  max: number;
  color: string;
}) {
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
  const [data, setData] = useState<SheetData | null>(null);
  const [loading, setLoading] = useState(true);
  const [error, setError] = useState<string | null>(null);
  const [configured, setConfigured] = useState<boolean | null>(null);
  const [log, setLog] = useState<LogEntry[]>([]);
  const [lastRefresh, setLastRefresh] = useState<string | null>(null);
  const [showActive, setShowActive] = useState(true);
  const prevFP = useRef<string | null>(null);

  // Add row
  const [addOpen, setAddOpen] = useState(false);
  const [addValues, setAddValues] = useState<Record<string, string>>({});
  const [adding, setAdding] = useState(false);
  const [showAllFields, setShowAllFields] = useState(false);

  const emit = useCallback((type: LogEntry["type"], msg: string) => {
    setLog((prev) => [{ ts: now(), type, msg }, ...prev.slice(0, 199)]);
  }, []);

  // ─── Data fetching ────────────────────────────────────────────────────

  const fetchData = useCallback(
    async (silent = false) => {
      if (!silent) setLoading(true);
      setError(null);
      try {
        const res = await fetch("/api/onedrive/sheets?action=read");
        const result = await res.json();
        if (!result.success) {
          setError(result.error);
          if (!silent) emit("error", result.error);
          return;
        }
        const d: SheetData = result.data;
        const fp = fingerprint(d);

        if (prevFP.current && prevFP.current !== fp) {
          const prevCount = parseInt(prevFP.current.split("|")[0]);
          const newCount = d.rawValues.length;
          if (newCount !== prevCount) {
            const diff = newCount - prevCount;
            emit("delta", `Rows: ${prevCount} → ${newCount} (${diff > 0 ? "+" : ""}${diff})`);
          } else {
            emit("delta", "Cell content changed on the spreadsheet");
          }
        }
        prevFP.current = fp;
        setData(d);
        setLastRefresh(now());
        if (!silent) emit("load", `${d.rows.length} rows × ${d.headers.length} columns`);
      } catch {
        if (!silent) emit("error", "Network error");
      } finally {
        setLoading(false);
      }
    },
    [emit]
  );

  // Auto-load on mount
  useEffect(() => {
    (async () => {
      try {
        const res = await fetch("/api/onedrive/sheets?action=status");
        const { data: st } = await res.json();
        if (st.authenticated && st.configuredFileId) {
          setConfigured(true);
          fetchData();
        } else {
          setConfigured(false);
          setLoading(false);
        }
      } catch {
        setConfigured(false);
        setLoading(false);
      }
    })();
  }, []); // eslint-disable-line react-hooks/exhaustive-deps

  // Auto-refresh
  useEffect(() => {
    if (!data) return;
    const id = setInterval(() => fetchData(true), 30000);
    return () => clearInterval(id);
  }, [data, fetchData]);

  // ─── Analytics ────────────────────────────────────────────────────────

  const analytics = useMemo(() => {
    if (!data) return null;
    const posCol = col(data.headers, "position");
    const prodCol = col(data.headers, "product");
    const origCol = col(data.headers, "origin");
    const monthCol = col(data.headers, "month");
    const packerCol = col(data.headers, "packer");

    const active = posCol
      ? data.rows.filter((r) => String(r[posCol] ?? "").toLowerCase().trim() === "active")
      : data.rows;
    const inactive = data.rows.length - active.length;

    const byProduct = countBy(active, prodCol);
    const byOrigin = countBy(active, origCol);
    const byMonth = countBy(active, monthCol);
    const byPacker = countBy(active, packerCol);

    // Sort months chronologically
    const monthSorted: [string, number][] = [...byMonth.entries()].sort(
      (a, b) => monthIndex(a[0]) - monthIndex(b[0])
    );

    // Concentration risk
    const topProduct = topN(byProduct, 1)[0];
    const topOrigin = topN(byOrigin, 1)[0];
    const topPacker = topN(byPacker, 1)[0];
    const total = active.length || 1;

    // Product × Origin matrix (top 6 × top 6)
    const topProds = topN(byProduct, 6).map(([n]) => n);
    const topOrigs = topN(byOrigin, 6).map(([n]) => n);
    const matrix: number[][] = topProds.map((p) =>
      topOrigs.map((o) =>
        active.filter(
          (r) =>
            prodCol && origCol &&
            String(r[prodCol] ?? "").trim() === p &&
            String(r[origCol] ?? "").trim() === o
        ).length
      )
    );

    return {
      posCol,
      prodCol,
      origCol,
      active,
      inactive,
      byProduct: topN(byProduct, 10),
      byOrigin: topN(byOrigin, 10),
      byMonth: monthSorted,
      byPacker: topN(byPacker, 10),
      risk: {
        product: topProduct ? { name: topProduct[0], pct: pct(topProduct[1], total) } : null,
        origin: topOrigin ? { name: topOrigin[0], pct: pct(topOrigin[1], total) } : null,
        packer: topPacker ? { name: topPacker[0], pct: pct(topPacker[1], total) } : null,
      },
      matrix: { products: topProds, origins: topOrigs, values: matrix },
    };
  }, [data]);

  const displayRows = analytics ? (showActive ? analytics.active : data?.rows ?? []) : [];
  const ranked = data ? rankColumns(data.headers) : [];

  // ─── Add row handler ──────────────────────────────────────────────────

  const handleAdd = async () => {
    if (!data) return;
    setAdding(true);
    setError(null);
    const { lastRow, lastCol } = parseRange(data.range);
    const next = lastRow + 1;
    const range = `A${next}:${lastCol}${next}`;
    const values = [data.headers.map((h) => addValues[h] || "")];

    try {
      const res = await fetch("/api/onedrive/sheets", {
        method: "POST",
        headers: { "Content-Type": "application/json" },
        body: JSON.stringify({ action: "write", range, values }),
      });
      const result = await res.json();
      if (!result.success) {
        emit("error", `Write failed: ${result.error}`);
      } else {
        const summary = Object.entries(addValues)
          .filter(([, v]) => v)
          .slice(0, 4)
          .map(([k, v]) => `${k}=${v}`)
          .join(", ");
        emit("add", `Row ${next}: ${summary}`);
        setAddValues({});
        fetchData(true);
      }
    } catch {
      emit("error", "Network error writing row");
    } finally {
      setAdding(false);
    }
  };

  // ─── Render: Loading ──────────────────────────────────────────────────

  if (loading && !data) {
    return (
      <div className="min-h-screen bg-zinc-950 text-zinc-100">
        <header className="border-b border-zinc-800 bg-zinc-900/50 backdrop-blur-sm sticky top-0 z-10">
          <div className="max-w-[1400px] mx-auto px-6 py-4">
            <h1 className="text-lg font-bold tracking-tight">Hectar Control Tower</h1>
            <p className="text-xs text-zinc-500">Commodity Trading & Risk Management</p>
          </div>
        </header>
        <main className="max-w-[1400px] mx-auto px-6 py-6 space-y-4">
          <div className="grid grid-cols-5 gap-3">
            {[1, 2, 3, 4, 5].map((i) => (
              <div key={i} className="rounded-lg border border-zinc-800 bg-zinc-900/30 p-4 animate-pulse">
                <div className="h-7 bg-zinc-800 rounded w-12 mb-1.5" />
                <div className="h-3 bg-zinc-800/50 rounded w-16" />
              </div>
            ))}
          </div>
          <div className="rounded-lg border border-zinc-800 bg-zinc-900/30 p-4 animate-pulse">
            <div className="h-4 bg-zinc-800/50 rounded w-32 mb-4" />
            {[1, 2, 3, 4, 5, 6, 7, 8].map((i) => (
              <div key={i} className="h-7 bg-zinc-800/30 rounded mb-1" />
            ))}
          </div>
        </main>
      </div>
    );
  }

  // ─── Render: Setup needed ─────────────────────────────────────────────

  if (configured === false) {
    return (
      <div className="min-h-screen bg-zinc-950 text-zinc-100 flex items-center justify-center">
        <div className="max-w-md rounded-lg border border-zinc-800 bg-zinc-900/50 p-8">
          <h1 className="text-lg font-bold mb-4">Hectar Control Tower</h1>
          <p className="text-sm text-zinc-400 mb-4">
            Set these Vercel environment variables, then redeploy:
          </p>
          <div className="space-y-2 font-mono text-xs">
            <div className="bg-zinc-900 border border-zinc-700 rounded px-3 py-2 text-zinc-300">
              ONEDRIVE_FILE_ID=<span className="text-amber-400">your_composite_id</span>
            </div>
            <div className="bg-zinc-900 border border-zinc-700 rounded px-3 py-2 text-zinc-300">
              ONEDRIVE_WORKSHEET_NAME=<span className="text-amber-400">Main Sheet</span>
            </div>
          </div>
          <p className="text-xs text-zinc-500 mt-4">
            The composite ID is found via Shared With Me in the Graph API.
            You already have it from the Feeder tab.
          </p>
        </div>
      </div>
    );
  }

  // ─── Render: Main CTRM ────────────────────────────────────────────────

  const keyFields = data ? data.headers.slice(0, 10) : [];
  const extraFields = data ? data.headers.slice(10) : [];

  return (
    <div className="min-h-screen bg-zinc-950 text-zinc-100">
      {/* Header */}
      <header className="border-b border-zinc-800 bg-zinc-900/50 backdrop-blur-sm sticky top-0 z-10">
        <div className="max-w-[1400px] mx-auto px-6 py-3 flex items-center justify-between">
          <div>
            <h1 className="text-lg font-bold tracking-tight">Hectar Control Tower</h1>
            <p className="text-[11px] text-zinc-500">Commodity Trading & Risk Management</p>
          </div>
          <div className="flex items-center gap-4">
            {lastRefresh && (
              <span className="text-[11px] text-zinc-600">{lastRefresh}</span>
            )}
            <button
              onClick={() => fetchData()}
              disabled={loading}
              className="px-3 py-1 text-xs bg-zinc-800 hover:bg-zinc-700 text-zinc-300 rounded border border-zinc-700 transition-colors disabled:opacity-50"
            >
              {loading ? "..." : "Refresh"}
            </button>
            <div className="flex items-center gap-1.5">
              <div className="h-2 w-2 rounded-full bg-emerald-500" />
              <span className="text-[11px] text-zinc-400">Live</span>
            </div>
          </div>
        </div>
      </header>

      <main className="max-w-[1400px] mx-auto px-6 py-5 space-y-5">
        {error && (
          <div className="p-2 bg-red-900/20 border border-red-800 rounded text-xs text-red-400">
            {error}
          </div>
        )}

        {/* ═══ Summary Metrics ═══ */}
        {analytics && (
          <div className="grid grid-cols-2 md:grid-cols-5 gap-3">
            <div className="rounded-lg border border-zinc-800 bg-zinc-900/60 px-4 py-3">
              <div className="text-2xl font-bold text-emerald-400">
                {analytics.active.length}
              </div>
              <div className="text-[11px] text-zinc-500">Active Positions</div>
            </div>
            <div className="rounded-lg border border-zinc-800 bg-zinc-900/60 px-4 py-3">
              <div className="text-2xl font-bold text-zinc-500">
                {analytics.inactive}
              </div>
              <div className="text-[11px] text-zinc-500">Inactive</div>
            </div>
            <div className="rounded-lg border border-zinc-800 bg-zinc-900/60 px-4 py-3">
              <div className="text-2xl font-bold text-blue-400">
                {analytics.byProduct.length}
              </div>
              <div className="text-[11px] text-zinc-500">Commodities</div>
            </div>
            <div className="rounded-lg border border-zinc-800 bg-zinc-900/60 px-4 py-3">
              <div className="text-2xl font-bold text-amber-400">
                {analytics.byOrigin.length}
              </div>
              <div className="text-[11px] text-zinc-500">Origins</div>
            </div>
            <div className="rounded-lg border border-zinc-800 bg-zinc-900/60 px-4 py-3">
              <div className="text-2xl font-bold text-purple-400">
                {analytics.byPacker.length}
              </div>
              <div className="text-[11px] text-zinc-500">Counterparties</div>
            </div>
          </div>
        )}

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
                      <th className="px-2 py-2 bg-zinc-800 text-zinc-500 border-b border-zinc-700 text-left font-medium w-8">
                        #
                      </th>
                      {ranked.slice(0, 20).map((h) => (
                        <th
                          key={h}
                          className="px-2 py-2 bg-zinc-800 text-zinc-500 border-b border-zinc-700 text-left font-medium whitespace-nowrap"
                        >
                          {h}
                        </th>
                      ))}
                    </tr>
                  </thead>
                  <tbody>
                    {displayRows.map((row, ri) => {
                      const posVal = analytics?.posCol
                        ? String(row[analytics.posCol] ?? "").toLowerCase().trim()
                        : "";
                      return (
                        <tr
                          key={ri}
                          className={`border-b border-zinc-800/30 hover:bg-zinc-800/40 ${
                            posVal === "inactive" ? "opacity-30" : ""
                          }`}
                        >
                          <td className="px-2 py-1.5 text-zinc-600 font-mono">{ri + 1}</td>
                          {ranked.slice(0, 20).map((h) => {
                            const v = row[h];
                            const s = v !== null && v !== undefined ? String(v) : "";
                            const isActive = h === analytics?.posCol && posVal === "active";
                            return (
                              <td
                                key={h}
                                className={`px-2 py-1.5 max-w-[160px] truncate ${
                                  isActive ? "text-emerald-400 font-medium" : "text-zinc-300"
                                }`}
                              >
                                {s}
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
        {analytics && (
          <>
            <h2 className="text-sm font-semibold text-zinc-200 pt-2">Exposure Analysis</h2>
            <div className="grid grid-cols-1 md:grid-cols-3 gap-4">
              {/* By Commodity */}
              <div className="rounded-lg border border-zinc-800 bg-zinc-900/30 p-4">
                <h3 className="text-xs font-medium text-zinc-400 mb-3">By Commodity</h3>
                <div className="space-y-1.5">
                  {analytics.byProduct.map(([name, count]) => (
                    <Bar key={name} label={name} count={count} max={analytics.byProduct[0]?.[1] || 1} color="bg-blue-500" />
                  ))}
                </div>
              </div>

              {/* By Origin */}
              <div className="rounded-lg border border-zinc-800 bg-zinc-900/30 p-4">
                <h3 className="text-xs font-medium text-zinc-400 mb-3">By Origin</h3>
                <div className="space-y-1.5">
                  {analytics.byOrigin.map(([name, count]) => (
                    <Bar key={name} label={name} count={count} max={analytics.byOrigin[0]?.[1] || 1} color="bg-amber-500" />
                  ))}
                </div>
              </div>

              {/* Delivery Schedule */}
              <div className="rounded-lg border border-zinc-800 bg-zinc-900/30 p-4">
                <h3 className="text-xs font-medium text-zinc-400 mb-3">Delivery Schedule</h3>
                <div className="space-y-1.5">
                  {analytics.byMonth.map(([name, count]) => (
                    <Bar key={name} label={name} count={count} max={Math.max(...analytics.byMonth.map(([, c]) => c), 1)} color="bg-emerald-500" />
                  ))}
                </div>
              </div>
            </div>
          </>
        )}

        {/* ═══ Counterparty & Risk ═══ */}
        {analytics && (
          <div className="grid grid-cols-1 md:grid-cols-2 gap-4">
            {/* Counterparty */}
            <div className="rounded-lg border border-zinc-800 bg-zinc-900/30 p-4">
              <h3 className="text-xs font-medium text-zinc-400 mb-3">Counterparty Exposure</h3>
              <div className="space-y-1.5">
                {analytics.byPacker.map(([name, count]) => (
                  <Bar key={name} label={name} count={count} max={analytics.byPacker[0]?.[1] || 1} color="bg-purple-500" />
                ))}
              </div>
            </div>

            {/* Concentration Risk */}
            <div className="rounded-lg border border-zinc-800 bg-zinc-900/30 p-4">
              <h3 className="text-xs font-medium text-zinc-400 mb-3">Concentration Risk</h3>
              <div className="space-y-4">
                {analytics.risk.packer && (
                  <div>
                    <div className="flex items-baseline justify-between mb-1">
                      <span className="text-xs text-zinc-300">Top Counterparty</span>
                      <span className={`text-xs font-medium ${riskColor(analytics.risk.packer.pct)}`}>
                        {analytics.risk.packer.pct}% — {riskLabel(analytics.risk.packer.pct)}
                      </span>
                    </div>
                    <div className="h-2 bg-zinc-800 rounded-full overflow-hidden">
                      <div
                        className={`h-full rounded-full ${analytics.risk.packer.pct >= 40 ? "bg-red-500" : analytics.risk.packer.pct >= 25 ? "bg-amber-500" : "bg-emerald-500"}`}
                        style={{ width: `${analytics.risk.packer.pct}%` }}
                      />
                    </div>
                    <div className="text-[11px] text-zinc-500 mt-0.5">{analytics.risk.packer.name}</div>
                  </div>
                )}
                {analytics.risk.origin && (
                  <div>
                    <div className="flex items-baseline justify-between mb-1">
                      <span className="text-xs text-zinc-300">Top Origin</span>
                      <span className={`text-xs font-medium ${riskColor(analytics.risk.origin.pct)}`}>
                        {analytics.risk.origin.pct}% — {riskLabel(analytics.risk.origin.pct)}
                      </span>
                    </div>
                    <div className="h-2 bg-zinc-800 rounded-full overflow-hidden">
                      <div
                        className={`h-full rounded-full ${analytics.risk.origin.pct >= 40 ? "bg-red-500" : analytics.risk.origin.pct >= 25 ? "bg-amber-500" : "bg-emerald-500"}`}
                        style={{ width: `${analytics.risk.origin.pct}%` }}
                      />
                    </div>
                    <div className="text-[11px] text-zinc-500 mt-0.5">{analytics.risk.origin.name}</div>
                  </div>
                )}
                {analytics.risk.product && (
                  <div>
                    <div className="flex items-baseline justify-between mb-1">
                      <span className="text-xs text-zinc-300">Top Commodity</span>
                      <span className={`text-xs font-medium ${riskColor(analytics.risk.product.pct)}`}>
                        {analytics.risk.product.pct}% — {riskLabel(analytics.risk.product.pct)}
                      </span>
                    </div>
                    <div className="h-2 bg-zinc-800 rounded-full overflow-hidden">
                      <div
                        className={`h-full rounded-full ${analytics.risk.product.pct >= 40 ? "bg-red-500" : analytics.risk.product.pct >= 25 ? "bg-amber-500" : "bg-emerald-500"}`}
                        style={{ width: `${analytics.risk.product.pct}%` }}
                      />
                    </div>
                    <div className="text-[11px] text-zinc-500 mt-0.5">{analytics.risk.product.name}</div>
                  </div>
                )}
              </div>
            </div>
          </div>
        )}

        {/* ═══ Exposure Matrix (Product × Origin) ═══ */}
        {analytics && analytics.matrix.products.length > 0 && analytics.matrix.origins.length > 0 && (
          <div className="rounded-lg border border-zinc-800 bg-zinc-900/30 p-4">
            <h3 className="text-xs font-medium text-zinc-400 mb-3">
              Commodity × Origin Matrix
            </h3>
            <div className="overflow-x-auto">
              <table className="text-xs border-collapse">
                <thead>
                  <tr>
                    <th className="px-2 py-1.5 text-left text-zinc-500 font-medium" />
                    {analytics.matrix.origins.map((o) => (
                      <th key={o} className="px-3 py-1.5 text-zinc-400 font-medium whitespace-nowrap">
                        {o}
                      </th>
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
                              backgroundColor:
                                v > 0 ? `rgba(59, 130, 246, ${0.1 + intensity * 0.5})` : "transparent",
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

        {/* ═══ Add Trade ═══ */}
        {data && (
          <div className="rounded-lg border border-zinc-800 bg-zinc-900/30">
            <button
              onClick={() => setAddOpen(!addOpen)}
              className="w-full flex items-center justify-between px-4 py-3 text-sm text-zinc-300 hover:bg-zinc-800/30 transition-colors"
            >
              <span className="font-medium">Add Trade</span>
              <span className="text-zinc-600 text-xs">
                {addOpen ? "collapse" : `row ${parseRange(data.range).lastRow + 1}`}
              </span>
            </button>
            {addOpen && (
              <div className="px-4 pb-4 border-t border-zinc-800">
                <div className="grid grid-cols-2 md:grid-cols-3 lg:grid-cols-5 gap-2 mt-3 mb-3">
                  {keyFields.map((h) => (
                    <div key={h}>
                      <label className="block text-[11px] text-zinc-500 mb-0.5 truncate">{h}</label>
                      <input
                        type="text"
                        value={addValues[h] || ""}
                        onChange={(e) => setAddValues((p) => ({ ...p, [h]: e.target.value }))}
                        className="w-full px-2 py-1.5 bg-zinc-900 border border-zinc-700 rounded text-xs text-zinc-200 focus:outline-none focus:border-blue-500"
                        placeholder="—"
                      />
                    </div>
                  ))}
                </div>
                {extraFields.length > 0 && (
                  <>
                    <button
                      onClick={() => setShowAllFields(!showAllFields)}
                      className="text-xs text-blue-400 hover:text-blue-300 mb-2"
                    >
                      {showAllFields ? "Hide" : `+${extraFields.length} fields`}
                    </button>
                    {showAllFields && (
                      <div className="grid grid-cols-3 md:grid-cols-5 lg:grid-cols-7 gap-2 mb-3 max-h-48 overflow-y-auto">
                        {extraFields.map((h) => (
                          <div key={h}>
                            <label className="block text-[11px] text-zinc-500 mb-0.5 truncate">{h}</label>
                            <input
                              type="text"
                              value={addValues[h] || ""}
                              onChange={(e) => setAddValues((p) => ({ ...p, [h]: e.target.value }))}
                              className="w-full px-2 py-1.5 bg-zinc-900 border border-zinc-700 rounded text-xs text-zinc-200 focus:outline-none focus:border-blue-500"
                              placeholder="—"
                            />
                          </div>
                        ))}
                      </div>
                    )}
                  </>
                )}
                <button
                  onClick={handleAdd}
                  disabled={adding}
                  className="px-4 py-2 bg-emerald-600 hover:bg-emerald-700 disabled:bg-zinc-700 text-white text-sm font-medium rounded-lg transition-colors"
                >
                  {adding ? "Writing..." : "Add Trade"}
                </button>
              </div>
            )}
          </div>
        )}

        {/* ═══ Activity Log ═══ */}
        {log.length > 0 && (
          <div className="rounded-lg border border-zinc-800 bg-zinc-900/30 p-4">
            <div className="flex items-center justify-between mb-2">
              <h3 className="text-xs font-medium text-zinc-500">Activity Log</h3>
              <button onClick={() => setLog([])} className="text-[10px] text-zinc-600 hover:text-zinc-400">
                clear
              </button>
            </div>
            <div className="space-y-0.5 max-h-32 overflow-y-auto">
              {log.map((e, i) => (
                <div key={i} className="flex items-center gap-2 text-[11px]">
                  <span className="text-zinc-600 font-mono w-[120px] shrink-0">{e.ts}</span>
                  <span
                    className={`px-1.5 py-0.5 rounded text-[10px] font-medium w-12 text-center shrink-0 ${
                      e.type === "load"
                        ? "bg-blue-900/40 text-blue-400"
                        : e.type === "add"
                          ? "bg-emerald-900/40 text-emerald-400"
                          : e.type === "delta"
                            ? "bg-purple-900/40 text-purple-400"
                            : "bg-red-900/40 text-red-400"
                    }`}
                  >
                    {e.type === "delta" ? "Δ" : e.type.toUpperCase()}
                  </span>
                  <span className="text-zinc-400 truncate">{e.msg}</span>
                </div>
              ))}
            </div>
          </div>
        )}

        {/* Footer */}
        <div className="text-center text-[11px] text-zinc-700 pt-2 pb-4">
          Hectar CTRM v0.2 · Auto-refresh 30s · Microsoft Graph API
        </div>
      </main>
    </div>
  );
}
