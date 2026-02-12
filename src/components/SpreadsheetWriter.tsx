"use client";

import { useState } from "react";

export default function SpreadsheetWriter() {
  const [range, setRange] = useState("A1:C2");
  const [valuesText, setValuesText] = useState(
    JSON.stringify(
      [
        ["Header1", "Header2", "Header3"],
        ["Value1", "Value2", "Value3"],
      ],
      null,
      2
    )
  );
  const [fileId, setFileId] = useState("");
  const [worksheet, setWorksheet] = useState("");
  const [loading, setLoading] = useState(false);
  const [result, setResult] = useState<string | null>(null);
  const [error, setError] = useState<string | null>(null);

  const writeData = async () => {
    setLoading(true);
    setError(null);
    setResult(null);

    let values: (string | number | boolean | null)[][];
    try {
      values = JSON.parse(valuesText);
      if (!Array.isArray(values) || !Array.isArray(values[0])) {
        throw new Error("Values must be a 2D array");
      }
    } catch (e) {
      setError(
        `Invalid JSON: ${e instanceof Error ? e.message : "Parse error"}. Values must be a 2D array like [["a","b"],["c","d"]]`
      );
      setLoading(false);
      return;
    }

    try {
      const body: Record<string, unknown> = {
        action: "write",
        range,
        values,
      };
      if (fileId) body.fileId = fileId;
      if (worksheet) body.worksheet = worksheet;

      const res = await fetch("/api/onedrive/sheets", {
        method: "POST",
        headers: { "Content-Type": "application/json" },
        body: JSON.stringify(body),
      });
      const data = await res.json();

      if (!data.success) {
        setError(data.error);
      } else {
        setResult(
          `Successfully wrote to range ${data.data?.address || range}`
        );
      }
    } catch {
      setError("Failed to write data");
    } finally {
      setLoading(false);
    }
  };

  return (
    <div className="rounded-lg border border-zinc-700 bg-zinc-800/50 p-6">
      <h2 className="text-lg font-semibold text-zinc-100 mb-4">
        Write to Spreadsheet
      </h2>

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
            Target Range (e.g. A1:C2)
          </label>
          <input
            type="text"
            value={range}
            onChange={(e) => setRange(e.target.value)}
            className="w-full px-3 py-2 bg-zinc-900 border border-zinc-600 rounded-lg text-sm text-zinc-200 placeholder-zinc-500 focus:outline-none focus:border-blue-500"
          />
        </div>
      </div>

      <div className="mb-4">
        <label className="block text-xs text-zinc-400 mb-1">
          Values (JSON 2D array)
        </label>
        <textarea
          value={valuesText}
          onChange={(e) => setValuesText(e.target.value)}
          rows={6}
          className="w-full px-3 py-2 bg-zinc-900 border border-zinc-600 rounded-lg text-sm text-zinc-200 font-mono placeholder-zinc-500 focus:outline-none focus:border-blue-500"
        />
      </div>

      <button
        onClick={writeData}
        disabled={loading}
        className="px-4 py-2 bg-amber-600 hover:bg-amber-700 disabled:bg-zinc-600 text-white text-sm font-medium rounded-lg transition-colors"
      >
        {loading ? "Writing..." : "Write Data"}
      </button>

      {error && (
        <div className="mt-4 p-3 bg-red-900/30 border border-red-700 rounded-lg text-sm text-red-300">
          {error}
        </div>
      )}

      {result && (
        <div className="mt-4 p-3 bg-emerald-900/30 border border-emerald-700 rounded-lg text-sm text-emerald-300">
          {result}
        </div>
      )}
    </div>
  );
}
