"use client";

import { useState } from "react";
import type { OneDriveFile } from "@/types/onedrive";

export default function FileSearch() {
  const [query, setQuery] = useState("");
  const [files, setFiles] = useState<OneDriveFile[]>([]);
  const [loading, setLoading] = useState(false);
  const [error, setError] = useState<string | null>(null);
  const [searched, setSearched] = useState(false);

  const search = async () => {
    if (!query.trim()) return;
    setLoading(true);
    setError(null);
    setSearched(true);

    try {
      const res = await fetch(
        `/api/onedrive/sheets?action=search&filename=${encodeURIComponent(query)}`
      );
      const data = await res.json();
      if (!data.success) {
        setError(data.error);
      } else {
        setFiles(data.data);
      }
    } catch {
      setError("Failed to search files");
    } finally {
      setLoading(false);
    }
  };

  const loadShared = async () => {
    setLoading(true);
    setError(null);
    setSearched(true);
    setQuery("");

    try {
      const res = await fetch("/api/onedrive/sheets?action=shared");
      const data = await res.json();
      if (!data.success) {
        setError(data.error);
      } else {
        setFiles(data.data);
      }
    } catch {
      setError("Failed to load shared files");
    } finally {
      setLoading(false);
    }
  };

  return (
    <div className="rounded-lg border border-zinc-700 bg-zinc-800/50 p-6">
      <h2 className="text-lg font-semibold text-zinc-100 mb-4">
        Find Excel Files
      </h2>

      <div className="flex gap-3 mb-4">
        <input
          type="text"
          value={query}
          onChange={(e) => setQuery(e.target.value)}
          onKeyDown={(e) => e.key === "Enter" && search()}
          placeholder="Search your OneDrive by file name..."
          className="flex-1 px-3 py-2 bg-zinc-900 border border-zinc-600 rounded-lg text-sm text-zinc-200 placeholder-zinc-500 focus:outline-none focus:border-blue-500"
        />
        <button
          onClick={search}
          disabled={loading}
          className="px-4 py-2 bg-blue-600 hover:bg-blue-700 disabled:bg-zinc-600 text-white text-sm font-medium rounded-lg transition-colors"
        >
          {loading ? "Loading..." : "Search"}
        </button>
        <button
          onClick={loadShared}
          disabled={loading}
          className="px-4 py-2 bg-amber-600 hover:bg-amber-700 disabled:bg-zinc-600 text-white text-sm font-medium rounded-lg transition-colors whitespace-nowrap"
        >
          Shared With Me
        </button>
      </div>

      {error && (
        <div className="p-3 bg-red-900/30 border border-red-700 rounded-lg text-sm text-red-300">
          {error}
        </div>
      )}

      {searched && files.length === 0 && !loading && !error && (
        <div className="text-sm text-zinc-400">
          No Excel files found. Ask your colleague to share the file with your
          Microsoft account, then click &quot;Shared With Me&quot;.
        </div>
      )}

      {files.length > 0 && (
        <div className="space-y-2">
          {files.map((file, idx) => {
            const displayId = file.compositeId || file.id;
            return (
              <div
                key={file.id || idx}
                className="p-3 bg-zinc-900 rounded-lg border border-zinc-700"
              >
                <div className="flex justify-between items-start">
                  <div className="min-w-0 flex-1">
                    <div className="flex items-center gap-2">
                      <span className="text-sm font-medium text-zinc-200">
                        {file.name}
                      </span>
                      {file.compositeId && (
                        <span className="text-[10px] bg-amber-900/50 text-amber-300 px-1.5 py-0.5 rounded font-medium">
                          SHARED
                        </span>
                      )}
                    </div>
                    <div className="text-xs text-zinc-400 mt-1 font-mono truncate">
                      ID: {displayId}
                    </div>
                    <div className="text-xs text-zinc-500 mt-1">
                      {file.size > 0
                        ? `${(file.size / 1024).toFixed(1)} KB | `
                        : ""}
                      {file.lastModifiedDateTime
                        ? `Modified: ${new Date(file.lastModifiedDateTime).toLocaleString()}`
                        : ""}
                      {file.lastModifiedBy?.user?.displayName &&
                        ` by ${file.lastModifiedBy.user.displayName}`}
                    </div>
                  </div>
                  <button
                    onClick={() => {
                      navigator.clipboard.writeText(displayId);
                    }}
                    className="text-xs text-blue-400 hover:text-blue-300 ml-3 whitespace-nowrap"
                  >
                    Copy ID
                  </button>
                </div>
              </div>
            );
          })}
        </div>
      )}
    </div>
  );
}
