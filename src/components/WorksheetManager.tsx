"use client";

import { useState } from "react";

interface Worksheet {
  id: string;
  name: string;
  position: number;
}

export default function WorksheetManager() {
  const [worksheets, setWorksheets] = useState<Worksheet[]>([]);
  const [loading, setLoading] = useState(false);
  const [error, setError] = useState<string | null>(null);
  const [newName, setNewName] = useState("");
  const [creating, setCreating] = useState(false);
  const [shareUrl, setShareUrl] = useState("");
  const [message, setMessage] = useState<string | null>(null);

  const fetchWorksheets = async () => {
    setLoading(true);
    setError(null);
    try {
      const params = new URLSearchParams({ action: "share-graph-worksheets" });
      if (shareUrl) params.set("url", shareUrl);

      const res = await fetch(`/api/onedrive/sheets?${params}`);
      const data = await res.json();

      if (!data.success) {
        setError(data.error);
      } else {
        setWorksheets(data.data);
      }
    } catch {
      setError("Failed to fetch worksheets");
    } finally {
      setLoading(false);
    }
  };

  const createWorksheet = async () => {
    if (!newName.trim()) return;
    setCreating(true);
    setError(null);
    setMessage(null);
    try {
      const body: Record<string, string> = {
        action: "share-graph-create-worksheet",
        name: newName.trim(),
      };
      if (shareUrl) body.url = shareUrl;

      const res = await fetch("/api/onedrive/sheets", {
        method: "POST",
        headers: { "Content-Type": "application/json" },
        body: JSON.stringify(body),
      });
      const data = await res.json();

      if (!data.success) {
        setError(data.error);
      } else {
        setMessage(`Created worksheet "${data.data.name}"`);
        setNewName("");
        // Refresh list
        fetchWorksheets();
      }
    } catch {
      setError("Failed to create worksheet");
    } finally {
      setCreating(false);
    }
  };

  return (
    <div className="rounded-lg border border-zinc-700 bg-zinc-800/50 p-6">
      <h2 className="text-lg font-semibold text-zinc-100 mb-4">
        Worksheet Manager
      </h2>

      <div className="mb-4">
        <label className="block text-xs text-zinc-400 mb-1">
          OneDrive Sharing URL
        </label>
        <input
          type="text"
          value={shareUrl}
          onChange={(e) => setShareUrl(e.target.value)}
          placeholder="Uses ONEDRIVE_SHARE_URL env var if empty"
          className="w-full px-3 py-2 bg-zinc-900 border border-zinc-600 rounded-lg text-sm text-zinc-200 placeholder-zinc-500 focus:outline-none focus:border-blue-500"
        />
      </div>

      <div className="flex gap-2 mb-4">
        <button
          onClick={fetchWorksheets}
          disabled={loading}
          className="px-4 py-2 bg-blue-600 hover:bg-blue-700 disabled:bg-zinc-600 text-white text-sm font-medium rounded-lg transition-colors"
        >
          {loading ? "Loading..." : "List Worksheets"}
        </button>
      </div>

      {worksheets.length > 0 && (
        <div className="mb-4">
          <div className="text-xs text-zinc-400 mb-2">
            {worksheets.length} worksheet{worksheets.length !== 1 ? "s" : ""} found
          </div>
          <div className="flex flex-wrap gap-2">
            {worksheets.map((ws) => (
              <div
                key={ws.id}
                className="px-3 py-1.5 bg-zinc-900 border border-zinc-600 rounded-lg text-sm text-zinc-200 flex items-center gap-2"
              >
                <span className="text-zinc-500 text-xs font-mono">
                  #{ws.position}
                </span>
                <span>{ws.name}</span>
              </div>
            ))}
          </div>
        </div>
      )}

      <div className="border-t border-zinc-700 pt-4 mt-4">
        <div className="text-xs text-zinc-400 mb-2">Create New Worksheet</div>
        <div className="flex gap-2">
          <input
            type="text"
            value={newName}
            onChange={(e) => setNewName(e.target.value)}
            placeholder="Worksheet name"
            className="flex-1 px-3 py-2 bg-zinc-900 border border-zinc-600 rounded-lg text-sm text-zinc-200 placeholder-zinc-500 focus:outline-none focus:border-blue-500"
            onKeyDown={(e) => {
              if (e.key === "Enter") createWorksheet();
            }}
          />
          <button
            onClick={createWorksheet}
            disabled={creating || !newName.trim()}
            className="px-4 py-2 bg-emerald-600 hover:bg-emerald-700 disabled:bg-zinc-600 text-white text-sm font-medium rounded-lg transition-colors"
          >
            {creating ? "Creating..." : "Create"}
          </button>
        </div>
      </div>

      {error && (
        <div className="mt-4 p-3 bg-red-900/30 border border-red-700 rounded-lg text-sm text-red-300">
          {error}
        </div>
      )}

      {message && (
        <div className="mt-4 p-3 bg-emerald-900/30 border border-emerald-700 rounded-lg text-sm text-emerald-300">
          {message}
        </div>
      )}
    </div>
  );
}
