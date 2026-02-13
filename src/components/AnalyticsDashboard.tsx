"use client";

import { useState, useEffect, useCallback, useRef } from "react";

interface SheetData {
  headers: string[];
  rows: Record<string, string | number | boolean | null>[];
  rawValues: (string | number | boolean | null)[][];
  range: string;
}

interface ChangeLogEntry {
  timestamp: string;
  action: "load" | "write" | "add_row" | "change_detected" | "error";
  details: string;
}

interface StatusData {
  authenticated: boolean;
  configuredFileId: string | null;
  configuredWorksheet: string | null;
  configuredShareUrl: string | null;
}

function parseLastRow(range: string): { lastRow: number; lastCol: string } {
  const match = range.match(/:([A-Z]+)(\d+)$/);
  if (match) return { lastCol: match[1], lastRow: parseInt(match[2]) };
  return { lastCol: "A", lastRow: 1 };
}

function ts(): string {
  return new Date().toLocaleString("en-GB", {
    day: "2-digit",
    month: "short",
    hour: "2-digit",
    minute: "2-digit",
    second: "2-digit",
  });
}

// Identify the most useful columns for a commodity trading positions view.
// Returns ordered list: identity cols → status cols → timing → quantities → rest
function rankColumns(headers: string[]): string[] {
  const lower = headers.map((h) => h.toLowerCase().trim());

  const identity = [
    "product",
    "commodity code",
    "commodity",
    "code",
    "origin",
    "variety",
    "packer",
    "counterparty",
    "supplier",
    "buyer",
    "contract",
    "contract no",
    "deal",
    "reference",
    "ref",
  ];
  const status = ["position", "status", "state", "active"];
  const timing = [
    "month",
    "delivery month",
    "shipment",
    "delivery",
    "eta",
    "etd",
    "date",
    "shipment date",
  ];
  const quantity = [
    "quantity",
    "qty",
    "volume",
    "weight",
    "mt",
    "tons",
    "bags",
    "containers",
    "container",
    "20s",
    "40s",
  ];
  const price = [
    "price",
    "rate",
    "value",
    "amount",
    "cost",
    "fob",
    "cif",
    "cfr",
    "premium",
    "margin",
    "pnl",
    "p&l",
    "usd",
    "total",
    "invoice",
  ];

  const buckets: string[][] = [[], [], [], [], [], []];

  headers.forEach((h) => {
    const l = h.toLowerCase().trim();
    if (identity.some((k) => l.includes(k))) buckets[0].push(h);
    else if (status.some((k) => l.includes(k))) buckets[1].push(h);
    else if (timing.some((k) => l.includes(k))) buckets[2].push(h);
    else if (quantity.some((k) => l.includes(k))) buckets[3].push(h);
    else if (price.some((k) => l.includes(k))) buckets[4].push(h);
    else buckets[5].push(h);
  });

  return buckets.flat();
}

export default function AnalyticsDashboard() {
  // ─── State ───
  const [status, setStatus] = useState<StatusData | null>(null);
  const [data, setData] = useState<SheetData | null>(null);
  const [loading, setLoading] = useState(true);
  const [error, setError] = useState<string | null>(null);
  const [changeLog, setChangeLog] = useState<ChangeLogEntry[]>([]);
  const [lastRefresh, setLastRefresh] = useState<string | null>(null);
  const prevRowCount = useRef<number | null>(null);

  // Add row
  const [addRowOpen, setAddRowOpen] = useState(false);
  const [addRowValues, setAddRowValues] = useState<Record<string, string>>({});
  const [addingRow, setAddingRow] = useState(false);
  const [showAllFields, setShowAllFields] = useState(false);

  // Filter
  const [showActive, setShowActive] = useState(true);

  const log = useCallback(
    (action: ChangeLogEntry["action"], details: string) => {
      setChangeLog((prev) => [{ timestamp: ts(), action, details }, ...prev.slice(0, 99)]);
    },
    []
  );

  // ─── Fetch data (uses env vars on server, no manual input needed) ───
  const fetchData = useCallback(
    async (silent = false) => {
      if (!silent) setLoading(true);
      setError(null);
      try {
        const res = await fetch("/api/onedrive/sheets?action=read");
        const result = await res.json();

        if (!result.success) {
          setError(result.error);
          if (!silent) log("error", result.error);
          return;
        }

        const newData: SheetData = result.data;
        const newRowCount = newData.rows.length;

        if (prevRowCount.current !== null && prevRowCount.current !== newRowCount) {
          const diff = newRowCount - prevRowCount.current;
          log(
            "change_detected",
            `Row count: ${prevRowCount.current} → ${newRowCount} (${diff > 0 ? "+" : ""}${diff})`
          );
        }

        prevRowCount.current = newRowCount;
        setData(newData);
        setLastRefresh(ts());
        if (!silent)
          log("load", `${newRowCount} rows, ${newData.headers.length} columns`);
      } catch {
        const msg = "Network error fetching data";
        setError(msg);
        if (!silent) log("error", msg);
      } finally {
        setLoading(false);
      }
    },
    [log]
  );

  // ─── On mount: check status then auto-load ───
  useEffect(() => {
    (async () => {
      try {
        const res = await fetch("/api/onedrive/sheets?action=status");
        const result = await res.json();
        const st: StatusData = result.data;
        setStatus(st);

        if (st.authenticated && st.configuredFileId) {
          fetchData();
        } else {
          setLoading(false);
        }
      } catch {
        setLoading(false);
      }
    })();
  }, []); // eslint-disable-line react-hooks/exhaustive-deps

  // Auto-refresh every 30s once data is loaded
  useEffect(() => {
    if (!data) return;
    const interval = setInterval(() => fetchData(true), 30000);
    return () => clearInterval(interval);
  }, [data, fetchData]);

  // ─── Analytics ───
  const positionCol = data?.headers.find(
    (h) => h.toLowerCase().trim() === "position"
  );

  const activeRows = data
    ? positionCol
      ? data.rows.filter(
          (r) =>
            String(r[positionCol] ?? "")
              .toLowerCase()
              .trim() === "active"
        )
      : data.rows
    : [];

  const inactiveCount = data
    ? positionCol
      ? data.rows.filter(
          (r) =>
            String(r[positionCol] ?? "")
              .toLowerCase()
              .trim() !== "active"
        ).length
      : 0
    : 0;

  const displayRows = showActive ? activeRows : data?.rows ?? [];

  const colIdx = (name: string) =>
    data?.headers.find((h) => h.toLowerCase().trim() === name.toLowerCase()) ||
    null;

  const uniqueCount = (colName: string): number => {
    const col = colIdx(colName);
    if (!col || !data) return 0;
    const set = new Set(
      data.rows.map((r) => String(r[col] ?? "").trim()).filter(Boolean)
    );
    return set.size;
  };

  // Ranked columns for the table
  const rankedCols = data ? rankColumns(data.headers) : [];

  // ─── Add row ───
  const handleAddRow = async () => {
    if (!data) return;
    setAddingRow(true);
    setError(null);

    const { lastRow, lastCol } = parseLastRow(data.range);
    const nextRow = lastRow + 1;
    const newRange = `A${nextRow}:${lastCol}${nextRow}`;
    const values = [data.headers.map((h) => addRowValues[h] || "")];

    try {
      const res = await fetch("/api/onedrive/sheets", {
        method: "POST",
        headers: { "Content-Type": "application/json" },
        body: JSON.stringify({ action: "write", range: newRange, values }),
      });
      const result = await res.json();

      if (!result.success) {
        setError(result.error);
        log("error", `Add row failed: ${result.error}`);
      } else {
        const filled = Object.entries(addRowValues)
          .filter(([, v]) => v)
          .slice(0, 5)
          .map(([k, v]) => `${k}=${v}`)
          .join(", ");
        log("add_row", `Row ${nextRow}: ${filled}`);
        setAddRowValues({});
        fetchData(true);
      }
    } catch {
      log("error", "Network error adding row");
    } finally {
      setAddingRow(false);
    }
  };

  const keyFields = data ? data.headers.slice(0, 10) : [];
  const extraFields = data ? data.headers.slice(10) : [];

  // ─── Render ───

  // Loading skeleton
  if (loading && !data) {
    return (
      <div className="space-y-4">
        <div className="grid grid-cols-4 gap-4">
          {[1, 2, 3, 4].map((i) => (
            <div
              key={i}
              className="rounded-lg border border-zinc-800 bg-zinc-800/30 p-5 animate-pulse"
            >
              <div className="h-8 bg-zinc-700/50 rounded w-16 mb-2" />
              <div className="h-3 bg-zinc-700/30 rounded w-24" />
            </div>
          ))}
        </div>
        <div className="rounded-lg border border-zinc-800 bg-zinc-800/30 p-6 animate-pulse">
          <div className="h-4 bg-zinc-700/30 rounded w-48 mb-4" />
          <div className="space-y-2">
            {[1, 2, 3, 4, 5].map((i) => (
              <div key={i} className="h-8 bg-zinc-700/20 rounded" />
            ))}
          </div>
        </div>
      </div>
    );
  }

  // Not configured — show one-line setup
  if (!loading && (!status?.authenticated || !status?.configuredFileId)) {
    return (
      <div className="rounded-lg border border-amber-800/50 bg-amber-900/10 p-6">
        <h2 className="text-lg font-semibold text-zinc-100 mb-2">
          One-Time Setup
        </h2>
        {!status?.authenticated ? (
          <p className="text-sm text-zinc-400">
            Switch to the{" "}
            <span className="text-blue-400 font-medium">OneDrive Feeder</span>{" "}
            tab and click{" "}
            <span className="text-blue-400">
              &quot;Connect to Microsoft OneDrive&quot;
            </span>{" "}
            to authenticate.
          </p>
        ) : (
          <div className="text-sm text-zinc-400 space-y-2">
            <p>
              Add your composite File ID as a Vercel environment variable:
            </p>
            <code className="block bg-zinc-900 text-zinc-300 px-3 py-2 rounded text-xs font-mono">
              ONEDRIVE_FILE_ID=your_composite_id_here
            </code>
            <p className="text-xs text-zinc-500">
              Find it in the OneDrive Feeder tab → click &quot;Shared With Me&quot; →
              copy the composite ID. Then add it in Vercel &gt; Settings &gt;
              Environment Variables and redeploy.
            </p>
          </div>
        )}
      </div>
    );
  }

  return (
    <div className="space-y-4">
      {/* ─── Summary Strip ─── */}
      <div className="grid grid-cols-2 md:grid-cols-4 gap-3">
        <div className="rounded-lg border border-zinc-800 bg-zinc-900/60 px-4 py-3">
          <div className="text-2xl font-bold text-emerald-400">
            {activeRows.length}
          </div>
          <div className="text-[11px] text-zinc-500">Active Positions</div>
        </div>
        <div className="rounded-lg border border-zinc-800 bg-zinc-900/60 px-4 py-3">
          <div className="text-2xl font-bold text-zinc-400">
            {inactiveCount}
          </div>
          <div className="text-[11px] text-zinc-500">Inactive</div>
        </div>
        <div className="rounded-lg border border-zinc-800 bg-zinc-900/60 px-4 py-3">
          <div className="text-2xl font-bold text-blue-400">
            {uniqueCount("product")}
          </div>
          <div className="text-[11px] text-zinc-500">Products</div>
        </div>
        <div className="rounded-lg border border-zinc-800 bg-zinc-900/60 px-4 py-3">
          <div className="text-2xl font-bold text-amber-400">
            {uniqueCount("origin")}
          </div>
          <div className="text-[11px] text-zinc-500">Origins</div>
        </div>
      </div>

      {/* ─── Toolbar ─── */}
      <div className="flex items-center justify-between">
        <div className="flex items-center gap-3">
          <h2 className="text-sm font-semibold text-zinc-200">
            {showActive ? "Active Positions" : "All Positions"}
          </h2>
          <span className="text-xs text-zinc-500">
            {displayRows.length} rows
          </span>
          <button
            onClick={() => setShowActive(!showActive)}
            className={`px-2.5 py-1 text-[11px] rounded-full border transition-colors ${
              showActive
                ? "border-emerald-700 bg-emerald-900/30 text-emerald-300"
                : "border-zinc-700 bg-zinc-800 text-zinc-400"
            }`}
          >
            {showActive ? "Active only" : "Show all"}
          </button>
        </div>
        <div className="flex items-center gap-3">
          {lastRefresh && (
            <span className="text-[11px] text-zinc-600">
              {lastRefresh} · auto-refresh 30s
            </span>
          )}
          <button
            onClick={() => fetchData()}
            disabled={loading}
            className="px-3 py-1 text-xs bg-zinc-800 hover:bg-zinc-700 text-zinc-300 rounded-lg border border-zinc-700 transition-colors disabled:opacity-50"
          >
            {loading ? "..." : "Refresh"}
          </button>
        </div>
      </div>

      {error && (
        <div className="p-2 bg-red-900/20 border border-red-800 rounded text-xs text-red-400">
          {error}
        </div>
      )}

      {/* ─── Positions Table ─── */}
      {data && displayRows.length > 0 && (
        <div className="rounded-lg border border-zinc-800 bg-zinc-900/40 overflow-hidden">
          <div className="overflow-x-auto max-h-[65vh]">
            <table className="w-full text-xs border-collapse">
              <thead className="sticky top-0 z-[1]">
                <tr>
                  <th className="px-2 py-2 bg-zinc-800 text-zinc-400 border-b border-zinc-700 text-left font-medium w-10">
                    #
                  </th>
                  {rankedCols.slice(0, 20).map((h) => (
                    <th
                      key={h}
                      className="px-2 py-2 bg-zinc-800 text-zinc-400 border-b border-zinc-700 text-left font-medium whitespace-nowrap"
                    >
                      {h}
                    </th>
                  ))}
                  {rankedCols.length > 20 && (
                    <th className="px-2 py-2 bg-zinc-800 text-zinc-500 border-b border-zinc-700 text-left font-medium">
                      +{rankedCols.length - 20}
                    </th>
                  )}
                </tr>
              </thead>
              <tbody>
                {displayRows.map((row, ri) => {
                  const pos = positionCol
                    ? String(row[positionCol] ?? "")
                        .toLowerCase()
                        .trim()
                    : "";
                  return (
                    <tr
                      key={ri}
                      className={`border-b border-zinc-800/50 hover:bg-zinc-800/40 ${
                        pos === "inactive" ? "opacity-40" : ""
                      }`}
                    >
                      <td className="px-2 py-1.5 text-zinc-600 font-mono">
                        {ri + 1}
                      </td>
                      {rankedCols.slice(0, 20).map((h) => {
                        const val = row[h];
                        const display =
                          val !== null && val !== undefined ? String(val) : "";
                        const isPos =
                          h === positionCol && pos === "active";
                        return (
                          <td
                            key={h}
                            className={`px-2 py-1.5 max-w-[180px] truncate ${
                              isPos
                                ? "text-emerald-400 font-medium"
                                : "text-zinc-300"
                            }`}
                          >
                            {display}
                          </td>
                        );
                      })}
                      {rankedCols.length > 20 && (
                        <td className="px-2 py-1.5 text-zinc-600">...</td>
                      )}
                    </tr>
                  );
                })}
              </tbody>
            </table>
          </div>
        </div>
      )}

      {data && displayRows.length === 0 && (
        <div className="rounded-lg border border-zinc-800 bg-zinc-900/40 p-8 text-center text-sm text-zinc-500">
          No {showActive ? "active positions" : "data"} found.
        </div>
      )}

      {/* ─── Add Row (collapsible) ─── */}
      {data && (
        <div className="rounded-lg border border-zinc-800 bg-zinc-900/40">
          <button
            onClick={() => setAddRowOpen(!addRowOpen)}
            className="w-full flex items-center justify-between px-4 py-3 text-sm text-zinc-300 hover:bg-zinc-800/30 transition-colors"
          >
            <span className="font-medium">Add New Row</span>
            <span className="text-zinc-500 text-xs">
              {addRowOpen ? "collapse" : `→ row ${parseLastRow(data.range).lastRow + 1}`}
            </span>
          </button>
          {addRowOpen && (
            <div className="px-4 pb-4 border-t border-zinc-800">
              <div className="grid grid-cols-2 md:grid-cols-3 lg:grid-cols-5 gap-2 mt-3 mb-3">
                {keyFields.map((header) => (
                  <div key={header}>
                    <label className="block text-[11px] text-zinc-500 mb-0.5 truncate">
                      {header}
                    </label>
                    <input
                      type="text"
                      value={addRowValues[header] || ""}
                      onChange={(e) =>
                        setAddRowValues((prev) => ({
                          ...prev,
                          [header]: e.target.value,
                        }))
                      }
                      className="w-full px-2 py-1.5 bg-zinc-900 border border-zinc-700 rounded text-xs text-zinc-200 placeholder-zinc-600 focus:outline-none focus:border-blue-500"
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
                    {showAllFields
                      ? "Hide extra fields"
                      : `Show all ${extraFields.length} additional fields`}
                  </button>
                  {showAllFields && (
                    <div className="grid grid-cols-2 md:grid-cols-4 lg:grid-cols-6 gap-2 mb-3 max-h-48 overflow-y-auto">
                      {extraFields.map((header) => (
                        <div key={header}>
                          <label className="block text-[11px] text-zinc-500 mb-0.5 truncate">
                            {header}
                          </label>
                          <input
                            type="text"
                            value={addRowValues[header] || ""}
                            onChange={(e) =>
                              setAddRowValues((prev) => ({
                                ...prev,
                                [header]: e.target.value,
                              }))
                            }
                            className="w-full px-2 py-1.5 bg-zinc-900 border border-zinc-700 rounded text-xs text-zinc-200 placeholder-zinc-600 focus:outline-none focus:border-blue-500"
                            placeholder="—"
                          />
                        </div>
                      ))}
                    </div>
                  )}
                </>
              )}
              <button
                onClick={handleAddRow}
                disabled={addingRow}
                className="px-4 py-2 bg-emerald-600 hover:bg-emerald-700 disabled:bg-zinc-700 text-white text-sm font-medium rounded-lg transition-colors"
              >
                {addingRow ? "Adding..." : "Add Row"}
              </button>
            </div>
          )}
        </div>
      )}

      {/* ─── Change Log ─── */}
      {changeLog.length > 0 && (
        <div className="rounded-lg border border-zinc-800 bg-zinc-900/40 p-4">
          <div className="flex items-center justify-between mb-2">
            <h3 className="text-xs font-medium text-zinc-400">Change Log</h3>
            <button
              onClick={() => setChangeLog([])}
              className="text-[10px] text-zinc-600 hover:text-zinc-400"
            >
              clear
            </button>
          </div>
          <div className="space-y-0.5 max-h-40 overflow-y-auto">
            {changeLog.map((entry, i) => (
              <div
                key={i}
                className="flex items-center gap-2 text-[11px] py-0.5"
              >
                <span className="text-zinc-600 font-mono w-[130px] shrink-0">
                  {entry.timestamp}
                </span>
                <span
                  className={`px-1.5 py-0.5 rounded text-[10px] font-medium w-14 text-center shrink-0 ${
                    entry.action === "load"
                      ? "bg-blue-900/40 text-blue-400"
                      : entry.action === "add_row"
                        ? "bg-emerald-900/40 text-emerald-400"
                        : entry.action === "change_detected"
                          ? "bg-purple-900/40 text-purple-400"
                          : entry.action === "write"
                            ? "bg-amber-900/40 text-amber-400"
                            : "bg-red-900/40 text-red-400"
                  }`}
                >
                  {entry.action === "add_row"
                    ? "ADD"
                    : entry.action === "change_detected"
                      ? "DELTA"
                      : entry.action === "load"
                        ? "LOAD"
                        : entry.action.toUpperCase()}
                </span>
                <span className="text-zinc-400 truncate">{entry.details}</span>
              </div>
            ))}
          </div>
        </div>
      )}
    </div>
  );
}
