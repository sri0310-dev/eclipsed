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
  action: "read" | "write" | "add_row" | "external_change" | "error";
  details: string;
}

function parseLastRow(range: string): { lastRow: number; lastCol: string } {
  // Range like "'Main Sheet'!A1:CG768" or "A1:CG768"
  const match = range.match(/:([A-Z]+)(\d+)$/);
  if (match) {
    return { lastCol: match[1], lastRow: parseInt(match[2]) };
  }
  return { lastCol: "A", lastRow: 1 };
}

function formatTimestamp(): string {
  return new Date().toLocaleString("en-US", {
    month: "short",
    day: "numeric",
    hour: "2-digit",
    minute: "2-digit",
    second: "2-digit",
  });
}

export default function AnalyticsDashboard() {
  const [data, setData] = useState<SheetData | null>(null);
  const [loading, setLoading] = useState(false);
  const [error, setError] = useState<string | null>(null);
  const [fileId, setFileId] = useState("");
  const [worksheet, setWorksheet] = useState("");
  const [changeLog, setChangeLog] = useState<ChangeLogEntry[]>([]);
  const [autoRefresh, setAutoRefresh] = useState(false);
  const [lastRefresh, setLastRefresh] = useState<string | null>(null);
  const prevRowCount = useRef<number | null>(null);

  // ─── Add Row state ───
  const [addRowValues, setAddRowValues] = useState<Record<string, string>>({});
  const [addingRow, setAddingRow] = useState(false);
  const [showAllFields, setShowAllFields] = useState(false);

  const log = useCallback((action: ChangeLogEntry["action"], details: string) => {
    setChangeLog((prev) => [
      { timestamp: formatTimestamp(), action, details },
      ...prev,
    ]);
  }, []);

  const fetchData = useCallback(async (silent = false) => {
    if (!silent) setLoading(true);
    setError(null);
    try {
      const params = new URLSearchParams({ action: "read" });
      if (fileId) params.set("fileId", fileId);
      if (worksheet) params.set("worksheet", worksheet);

      const res = await fetch(`/api/onedrive/sheets?${params}`);
      const result = await res.json();

      if (!result.success) {
        setError(result.error);
        if (!silent) log("error", result.error);
        return;
      }

      const newData: SheetData = result.data;
      const newRowCount = newData.rows.length;

      // Detect external changes
      if (prevRowCount.current !== null && prevRowCount.current !== newRowCount) {
        const diff = newRowCount - prevRowCount.current;
        log(
          "external_change",
          `Row count changed: ${prevRowCount.current} → ${newRowCount} (${diff > 0 ? "+" : ""}${diff})`
        );
      }

      prevRowCount.current = newRowCount;
      setData(newData);
      setLastRefresh(formatTimestamp());
      if (!silent) log("read", `Loaded ${newRowCount} rows, ${newData.headers.length} columns`);
    } catch {
      setError("Failed to fetch data");
      if (!silent) log("error", "Failed to fetch data from API");
    } finally {
      setLoading(false);
    }
  }, [fileId, worksheet, log]);

  // Auto-refresh interval
  useEffect(() => {
    if (!autoRefresh || !data) return;
    const interval = setInterval(() => fetchData(true), 30000);
    return () => clearInterval(interval);
  }, [autoRefresh, data, fetchData]);

  // ─── Analytics computations ───
  const analytics = data
    ? (() => {
        const colIndex = (name: string) =>
          data.headers.findIndex(
            (h) => h.toLowerCase().trim() === name.toLowerCase()
          );

        const positionIdx = colIndex("position");
        const productIdx = colIndex("product");
        const originIdx = colIndex("origin");
        const monthIdx = colIndex("month");

        const count = (
          idx: number,
          filter?: string
        ): number | Map<string, number> => {
          if (idx < 0) return filter ? 0 : new Map();
          if (filter) {
            return data.rows.filter(
              (r) =>
                String(r[data.headers[idx]] ?? "")
                  .toLowerCase()
                  .trim() === filter.toLowerCase()
            ).length;
          }
          const map = new Map<string, number>();
          data.rows.forEach((r) => {
            const val = String(r[data.headers[idx]] ?? "").trim();
            if (val) map.set(val, (map.get(val) || 0) + 1);
          });
          return map;
        };

        const activeCount =
          positionIdx >= 0 ? (count(positionIdx, "active") as number) : 0;
        const inactiveCount =
          positionIdx >= 0 ? (count(positionIdx, "inactive") as number) : 0;
        const productMap =
          productIdx >= 0
            ? (count(productIdx) as Map<string, number>)
            : new Map<string, number>();
        const originMap =
          originIdx >= 0
            ? (count(originIdx) as Map<string, number>)
            : new Map<string, number>();
        const monthMap =
          monthIdx >= 0
            ? (count(monthIdx) as Map<string, number>)
            : new Map<string, number>();

        const topN = (map: Map<string, number>, n: number) =>
          [...map.entries()]
            .sort((a, b) => b[1] - a[1])
            .slice(0, n);

        return {
          totalRows: data.rows.length,
          totalCols: data.headers.length,
          activeCount,
          inactiveCount,
          uniqueProducts: productMap.size,
          uniqueOrigins: originMap.size,
          topProducts: topN(productMap, 6),
          topOrigins: topN(originMap, 6),
          monthBreakdown: topN(monthMap, 12),
        };
      })()
    : null;

  // ─── Add Row handler ───
  const handleAddRow = async () => {
    if (!data) return;
    setAddingRow(true);
    setError(null);

    const { lastRow, lastCol } = parseLastRow(data.range);
    const nextRow = lastRow + 1;
    const newRange = `A${nextRow}:${lastCol}${nextRow}`;

    // Build values array matching column order
    const values = [
      data.headers.map((h) => addRowValues[h] || ""),
    ];

    try {
      const body: Record<string, unknown> = {
        action: "write",
        range: newRange,
        values,
      };
      if (fileId) body.fileId = fileId;
      if (worksheet) body.worksheet = worksheet;

      const res = await fetch("/api/onedrive/sheets", {
        method: "POST",
        headers: { "Content-Type": "application/json" },
        body: JSON.stringify(body),
      });
      const result = await res.json();

      if (!result.success) {
        setError(result.error);
        log("error", `Add row failed: ${result.error}`);
      } else {
        log(
          "add_row",
          `Added row ${nextRow}: ${Object.entries(addRowValues)
            .filter(([, v]) => v)
            .map(([k, v]) => `${k}=${v}`)
            .slice(0, 4)
            .join(", ")}${Object.keys(addRowValues).length > 4 ? "..." : ""}`
        );
        setAddRowValues({});
        // Refresh data to see the new row
        fetchData(true);
      }
    } catch {
      setError("Failed to add row");
      log("error", "Network error adding row");
    } finally {
      setAddingRow(false);
    }
  };

  // ─── Key fields for the compact form ───
  const keyFields = data
    ? data.headers.slice(0, 10)
    : [];
  const remainingFields = data
    ? data.headers.slice(10)
    : [];

  return (
    <div className="space-y-6">
      {/* Connection bar */}
      <div className="rounded-lg border border-zinc-700 bg-zinc-800/50 p-4">
        <div className="flex flex-wrap items-end gap-3">
          <div className="flex-1 min-w-[200px]">
            <label className="block text-xs text-zinc-400 mb-1">
              File ID (composite ID from Shared With Me)
            </label>
            <input
              type="text"
              value={fileId}
              onChange={(e) => setFileId(e.target.value)}
              placeholder="Uses ONEDRIVE_FILE_ID env var if empty"
              className="w-full px-3 py-2 bg-zinc-900 border border-zinc-600 rounded-lg text-sm text-zinc-200 placeholder-zinc-500 focus:outline-none focus:border-blue-500"
            />
          </div>
          <div className="w-40">
            <label className="block text-xs text-zinc-400 mb-1">
              Worksheet
            </label>
            <input
              type="text"
              value={worksheet}
              onChange={(e) => setWorksheet(e.target.value)}
              placeholder="Uses env default"
              className="w-full px-3 py-2 bg-zinc-900 border border-zinc-600 rounded-lg text-sm text-zinc-200 placeholder-zinc-500 focus:outline-none focus:border-blue-500"
            />
          </div>
          <button
            onClick={() => fetchData()}
            disabled={loading}
            className="px-5 py-2 bg-blue-600 hover:bg-blue-700 disabled:bg-zinc-600 text-white text-sm font-medium rounded-lg transition-colors"
          >
            {loading ? "Loading..." : "Load Data"}
          </button>
          <label className="flex items-center gap-2 text-xs text-zinc-400 cursor-pointer select-none">
            <input
              type="checkbox"
              checked={autoRefresh}
              onChange={(e) => setAutoRefresh(e.target.checked)}
              className="rounded border-zinc-600"
            />
            Auto-refresh (30s)
          </label>
          {lastRefresh && (
            <span className="text-[11px] text-zinc-500">
              Last: {lastRefresh}
            </span>
          )}
        </div>
      </div>

      {error && (
        <div className="p-3 bg-red-900/30 border border-red-700 rounded-lg text-sm text-red-300">
          {error}
        </div>
      )}

      {/* ─── Summary Cards ─── */}
      {analytics && (
        <div className="grid grid-cols-2 md:grid-cols-4 gap-4">
          <div className="rounded-lg border border-zinc-700 bg-zinc-800/50 p-4">
            <div className="text-3xl font-bold text-zinc-100">
              {analytics.totalRows.toLocaleString()}
            </div>
            <div className="text-xs text-zinc-400 mt-1">
              Total Rows ({analytics.totalCols} cols)
            </div>
          </div>
          <div className="rounded-lg border border-zinc-700 bg-zinc-800/50 p-4">
            <div className="text-3xl font-bold text-emerald-400">
              {analytics.activeCount.toLocaleString()}
            </div>
            <div className="text-xs text-zinc-400 mt-1">
              Active Positions
            </div>
            {analytics.inactiveCount > 0 && (
              <div className="text-xs text-zinc-500 mt-0.5">
                {analytics.inactiveCount} inactive
              </div>
            )}
          </div>
          <div className="rounded-lg border border-zinc-700 bg-zinc-800/50 p-4">
            <div className="text-3xl font-bold text-blue-400">
              {analytics.uniqueProducts}
            </div>
            <div className="text-xs text-zinc-400 mt-1">
              Unique Products
            </div>
          </div>
          <div className="rounded-lg border border-zinc-700 bg-zinc-800/50 p-4">
            <div className="text-3xl font-bold text-amber-400">
              {analytics.uniqueOrigins}
            </div>
            <div className="text-xs text-zinc-400 mt-1">
              Origins / Countries
            </div>
          </div>
        </div>
      )}

      {/* ─── Breakdown Cards ─── */}
      {analytics && (
        <div className="grid grid-cols-1 md:grid-cols-3 gap-4">
          {/* Products */}
          <div className="rounded-lg border border-zinc-700 bg-zinc-800/50 p-4">
            <h3 className="text-sm font-medium text-zinc-200 mb-3">
              Top Products
            </h3>
            <div className="space-y-2">
              {analytics.topProducts.map(([name, count]) => (
                <div key={name} className="flex items-center gap-2">
                  <div
                    className="h-2 rounded-full bg-blue-500"
                    style={{
                      width: `${Math.max(8, (count / analytics.totalRows) * 100)}%`,
                    }}
                  />
                  <span className="text-xs text-zinc-300 whitespace-nowrap">
                    {name}
                  </span>
                  <span className="text-xs text-zinc-500 ml-auto">{count}</span>
                </div>
              ))}
            </div>
          </div>

          {/* Origins */}
          <div className="rounded-lg border border-zinc-700 bg-zinc-800/50 p-4">
            <h3 className="text-sm font-medium text-zinc-200 mb-3">
              Top Origins
            </h3>
            <div className="space-y-2">
              {analytics.topOrigins.map(([name, count]) => (
                <div key={name} className="flex items-center gap-2">
                  <div
                    className="h-2 rounded-full bg-amber-500"
                    style={{
                      width: `${Math.max(8, (count / analytics.totalRows) * 100)}%`,
                    }}
                  />
                  <span className="text-xs text-zinc-300 whitespace-nowrap">
                    {name}
                  </span>
                  <span className="text-xs text-zinc-500 ml-auto">{count}</span>
                </div>
              ))}
            </div>
          </div>

          {/* Months */}
          <div className="rounded-lg border border-zinc-700 bg-zinc-800/50 p-4">
            <h3 className="text-sm font-medium text-zinc-200 mb-3">
              Monthly Distribution
            </h3>
            <div className="space-y-2">
              {analytics.monthBreakdown.map(([name, count]) => (
                <div key={name} className="flex items-center gap-2">
                  <div
                    className="h-2 rounded-full bg-emerald-500"
                    style={{
                      width: `${Math.max(8, (count / analytics.totalRows) * 100)}%`,
                    }}
                  />
                  <span className="text-xs text-zinc-300 whitespace-nowrap">
                    {name}
                  </span>
                  <span className="text-xs text-zinc-500 ml-auto">{count}</span>
                </div>
              ))}
            </div>
          </div>
        </div>
      )}

      {/* ─── Recent Rows ─── */}
      {data && data.rows.length > 0 && (
        <div className="rounded-lg border border-zinc-700 bg-zinc-800/50 p-4">
          <h3 className="text-sm font-medium text-zinc-200 mb-3">
            Latest 5 Rows
          </h3>
          <div className="overflow-auto">
            <table className="w-full text-xs border-collapse">
              <thead>
                <tr>
                  <th className="text-left px-2 py-1.5 bg-zinc-700 text-zinc-300 border border-zinc-600 font-medium">
                    #
                  </th>
                  {data.headers.slice(0, 12).map((h, i) => (
                    <th
                      key={i}
                      className="text-left px-2 py-1.5 bg-zinc-700 text-zinc-300 border border-zinc-600 font-medium whitespace-nowrap"
                    >
                      {h || `Col ${i + 1}`}
                    </th>
                  ))}
                  {data.headers.length > 12 && (
                    <th className="text-left px-2 py-1.5 bg-zinc-700 text-zinc-500 border border-zinc-600">
                      +{data.headers.length - 12} more
                    </th>
                  )}
                </tr>
              </thead>
              <tbody>
                {data.rows.slice(-5).map((row, ri) => (
                  <tr key={ri} className="hover:bg-zinc-700/50">
                    <td className="px-2 py-1 border border-zinc-700 text-zinc-500 font-mono">
                      {data.rows.length - 4 + ri}
                    </td>
                    {data.headers.slice(0, 12).map((h, ci) => (
                      <td
                        key={ci}
                        className="px-2 py-1 border border-zinc-700 text-zinc-300 max-w-[150px] truncate"
                      >
                        {row[h] !== null && row[h] !== undefined
                          ? String(row[h])
                          : ""}
                      </td>
                    ))}
                    {data.headers.length > 12 && (
                      <td className="px-2 py-1 border border-zinc-700 text-zinc-500">
                        ...
                      </td>
                    )}
                  </tr>
                ))}
              </tbody>
            </table>
          </div>
        </div>
      )}

      {/* ─── Add Row ─── */}
      {data && (
        <div className="rounded-lg border border-zinc-700 bg-zinc-800/50 p-4">
          <div className="flex items-center justify-between mb-3">
            <h3 className="text-sm font-medium text-zinc-200">
              Add New Row
            </h3>
            <span className="text-[11px] text-zinc-500">
              Will write to row {parseLastRow(data.range).lastRow + 1}
            </span>
          </div>

          <div className="grid grid-cols-2 md:grid-cols-3 lg:grid-cols-5 gap-2 mb-3">
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
                  className="w-full px-2 py-1.5 bg-zinc-900 border border-zinc-600 rounded text-xs text-zinc-200 placeholder-zinc-600 focus:outline-none focus:border-blue-500"
                  placeholder="—"
                />
              </div>
            ))}
          </div>

          {remainingFields.length > 0 && (
            <>
              <button
                onClick={() => setShowAllFields(!showAllFields)}
                className="text-xs text-blue-400 hover:text-blue-300 mb-2"
              >
                {showAllFields
                  ? "Hide additional fields"
                  : `Show all ${remainingFields.length} additional fields`}
              </button>
              {showAllFields && (
                <div className="grid grid-cols-2 md:grid-cols-4 lg:grid-cols-6 gap-2 mb-3 max-h-60 overflow-y-auto pr-1">
                  {remainingFields.map((header) => (
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
                        className="w-full px-2 py-1.5 bg-zinc-900 border border-zinc-600 rounded text-xs text-zinc-200 placeholder-zinc-600 focus:outline-none focus:border-blue-500"
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
            className="px-4 py-2 bg-emerald-600 hover:bg-emerald-700 disabled:bg-zinc-600 text-white text-sm font-medium rounded-lg transition-colors"
          >
            {addingRow ? "Adding..." : "Add Row"}
          </button>
        </div>
      )}

      {/* ─── Change Log ─── */}
      <div className="rounded-lg border border-zinc-700 bg-zinc-800/50 p-4">
        <div className="flex items-center justify-between mb-3">
          <h3 className="text-sm font-medium text-zinc-200">Change Log</h3>
          {changeLog.length > 0 && (
            <button
              onClick={() => setChangeLog([])}
              className="text-[11px] text-zinc-500 hover:text-zinc-300"
            >
              Clear
            </button>
          )}
        </div>

        {changeLog.length === 0 ? (
          <p className="text-xs text-zinc-500">
            No activity yet. Load data to start tracking changes.
          </p>
        ) : (
          <div className="space-y-1 max-h-60 overflow-y-auto">
            {changeLog.map((entry, i) => (
              <div
                key={i}
                className="flex items-start gap-2 text-xs py-1 border-b border-zinc-800 last:border-0"
              >
                <span className="text-zinc-500 font-mono whitespace-nowrap min-w-[140px]">
                  {entry.timestamp}
                </span>
                <span
                  className={`px-1.5 py-0.5 rounded text-[10px] font-medium whitespace-nowrap ${
                    entry.action === "read"
                      ? "bg-blue-900/50 text-blue-300"
                      : entry.action === "write"
                        ? "bg-amber-900/50 text-amber-300"
                        : entry.action === "add_row"
                          ? "bg-emerald-900/50 text-emerald-300"
                          : entry.action === "external_change"
                            ? "bg-purple-900/50 text-purple-300"
                            : "bg-red-900/50 text-red-300"
                  }`}
                >
                  {entry.action === "add_row"
                    ? "ADD"
                    : entry.action === "external_change"
                      ? "CHANGE"
                      : entry.action.toUpperCase()}
                </span>
                <span className="text-zinc-300">{entry.details}</span>
              </div>
            ))}
          </div>
        )}
      </div>
    </div>
  );
}
