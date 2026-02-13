"use client";

import { useState } from "react";
import type { SheetDataResponse } from "@/types/onedrive";

export default function SpreadsheetViewer() {
  const [data, setData] = useState<SheetDataResponse | null>(null);
  const [loading, setLoading] = useState(false);
  const [error, setError] = useState<string | null>(null);
  const [shareUrl, setShareUrl] = useState("");
  const [worksheet, setWorksheet] = useState("");
  const [mode, setMode] = useState<"share-link" | "graph-api">("graph-api");

  // Graph API mode fields
  const [fileId, setFileId] = useState("");
  const [range, setRange] = useState("");

  const readData = async () => {
    setLoading(true);
    setError(null);
    try {
      const params = new URLSearchParams();

      if (mode === "share-link") {
        params.set("action", "share-read");
        if (shareUrl) params.set("url", shareUrl);
        if (worksheet) params.set("worksheet", worksheet);
      } else {
        params.set("action", "read");
        if (fileId) params.set("fileId", fileId);
        if (worksheet) params.set("worksheet", worksheet);
        if (range) params.set("range", range);
      }

      const res = await fetch(`/api/onedrive/sheets?${params}`);
      const result = await res.json();

      if (!result.success) {
        setError(result.error);
      } else {
        setData(result.data);
      }
    } catch {
      setError("Failed to fetch data from the API");
    } finally {
      setLoading(false);
    }
  };

  return (
    <div className="rounded-lg border border-zinc-700 bg-zinc-800/50 p-6">
      <div className="flex items-center justify-between mb-4">
        <h2 className="text-lg font-semibold text-zinc-100">
          Read from Spreadsheet
        </h2>
        <div className="flex bg-zinc-900 rounded-lg p-0.5">
          <button
            onClick={() => setMode("share-link")}
            className={`px-3 py-1 text-xs rounded-md transition-colors ${
              mode === "share-link"
                ? "bg-amber-600 text-white"
                : "text-zinc-400 hover:text-zinc-200"
            }`}
          >
            Sharing Link
          </button>
          <button
            onClick={() => setMode("graph-api")}
            className={`px-3 py-1 text-xs rounded-md transition-colors ${
              mode === "graph-api"
                ? "bg-blue-600 text-white"
                : "text-zinc-400 hover:text-zinc-200"
            }`}
          >
            Graph API
          </button>
        </div>
      </div>

      {mode === "share-link" ? (
        <div className="grid grid-cols-1 md:grid-cols-2 gap-3 mb-4">
          <div className="md:col-span-2">
            <label className="block text-xs text-zinc-400 mb-1">
              OneDrive Sharing URL
            </label>
            <input
              type="text"
              value={shareUrl}
              onChange={(e) => setShareUrl(e.target.value)}
              placeholder="Paste sharing link (uses ONEDRIVE_SHARE_URL env var if empty)"
              className="w-full px-3 py-2 bg-zinc-900 border border-zinc-600 rounded-lg text-sm text-zinc-200 placeholder-zinc-500 focus:outline-none focus:border-amber-500"
            />
          </div>
          <div>
            <label className="block text-xs text-zinc-400 mb-1">
              Worksheet (optional)
            </label>
            <input
              type="text"
              value={worksheet}
              onChange={(e) => setWorksheet(e.target.value)}
              placeholder="First worksheet"
              className="w-full px-3 py-2 bg-zinc-900 border border-zinc-600 rounded-lg text-sm text-zinc-200 placeholder-zinc-500 focus:outline-none focus:border-amber-500"
            />
          </div>
          <div className="flex items-end">
            <p className="text-[11px] text-zinc-500 leading-tight">
              Downloads .xlsx directly from the sharing URL. No SPO license
              needed. File owner must set &quot;Anyone with the link&quot;
              access.
            </p>
          </div>
        </div>
      ) : (
        <div className="grid grid-cols-1 md:grid-cols-3 gap-3 mb-4">
          <div>
            <label className="block text-xs text-zinc-400 mb-1">
              File ID (optional override)
            </label>
            <input
              type="text"
              value={fileId}
              onChange={(e) => setFileId(e.target.value)}
              placeholder="Uses .env default"
              className="w-full px-3 py-2 bg-zinc-900 border border-zinc-600 rounded-lg text-sm text-zinc-200 placeholder-zinc-500 focus:outline-none focus:border-blue-500"
            />
          </div>
          <div>
            <label className="block text-xs text-zinc-400 mb-1">
              Worksheet (optional override)
            </label>
            <input
              type="text"
              value={worksheet}
              onChange={(e) => setWorksheet(e.target.value)}
              placeholder="Uses .env default"
              className="w-full px-3 py-2 bg-zinc-900 border border-zinc-600 rounded-lg text-sm text-zinc-200 placeholder-zinc-500 focus:outline-none focus:border-blue-500"
            />
          </div>
          <div>
            <label className="block text-xs text-zinc-400 mb-1">
              Range (optional, e.g. A1:Z50)
            </label>
            <input
              type="text"
              value={range}
              onChange={(e) => setRange(e.target.value)}
              placeholder="Auto-detect used range"
              className="w-full px-3 py-2 bg-zinc-900 border border-zinc-600 rounded-lg text-sm text-zinc-200 placeholder-zinc-500 focus:outline-none focus:border-blue-500"
            />
          </div>
        </div>
      )}

      <button
        onClick={readData}
        disabled={loading}
        className={`px-4 py-2 text-white text-sm font-medium rounded-lg transition-colors disabled:bg-zinc-600 ${
          mode === "share-link"
            ? "bg-amber-600 hover:bg-amber-700"
            : "bg-emerald-600 hover:bg-emerald-700"
        }`}
      >
        {loading ? "Reading..." : "Read Data"}
      </button>

      {error && (
        <div className="mt-4 p-3 bg-red-900/30 border border-red-700 rounded-lg text-sm text-red-300">
          {error}
        </div>
      )}

      {data && (
        <div className="mt-4 overflow-auto max-h-96">
          <div className="text-xs text-zinc-400 mb-2">
            Range: {data.range} | {data.rows.length} rows | {data.headers.length}{" "}
            columns
          </div>
          <table className="w-full text-sm border-collapse">
            <thead>
              <tr>
                {data.headers.map((h, i) => (
                  <th
                    key={i}
                    className="text-left px-3 py-2 bg-zinc-700 text-zinc-200 border border-zinc-600 font-medium text-xs"
                  >
                    {h || `Col ${i + 1}`}
                  </th>
                ))}
              </tr>
            </thead>
            <tbody>
              {data.rows.map((row, ri) => (
                <tr key={ri} className="hover:bg-zinc-700/50">
                  {data.headers.map((h, ci) => (
                    <td
                      key={ci}
                      className="px-3 py-1.5 border border-zinc-700 text-zinc-300 text-xs"
                    >
                      {row[h] !== null && row[h] !== undefined
                        ? String(row[h])
                        : ""}
                    </td>
                  ))}
                </tr>
              ))}
            </tbody>
          </table>
        </div>
      )}
    </div>
  );
}
