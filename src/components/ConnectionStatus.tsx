"use client";

import { useState, useEffect } from "react";

interface StatusData {
  authenticated: boolean;
  configuredFileId: string | null;
  configuredWorksheet: string | null;
  configuredShareUrl: string | null;
}

export default function ConnectionStatus() {
  const [status, setStatus] = useState<StatusData | null>(null);
  const [loading, setLoading] = useState(true);
  const [authUrl, setAuthUrl] = useState<string | null>(null);

  const checkStatus = async () => {
    setLoading(true);
    try {
      const res = await fetch("/api/onedrive/sheets?action=status");
      const data = await res.json();
      setStatus(data.data);

      if (!data.data.authenticated) {
        const authRes = await fetch("/api/onedrive/auth");
        const authData = await authRes.json();
        if (authData.authUrl) setAuthUrl(authData.authUrl);
      }
    } catch {
      console.error("Failed to check connection status");
    } finally {
      setLoading(false);
    }
  };

  useEffect(() => {
    checkStatus();

    // Check URL params for auth callback result
    const params = new URLSearchParams(window.location.search);
    if (params.get("auth_success")) {
      checkStatus();
      window.history.replaceState({}, "", "/");
    }
    if (params.get("auth_error")) {
      alert("Auth error: " + params.get("auth_error"));
      window.history.replaceState({}, "", "/");
    }
  }, []);

  if (loading) {
    return (
      <div className="rounded-lg border border-zinc-700 bg-zinc-800/50 p-6">
        <div className="animate-pulse flex items-center gap-3">
          <div className="h-3 w-3 rounded-full bg-zinc-600" />
          <span className="text-zinc-400">Checking connection...</span>
        </div>
      </div>
    );
  }

  const hasShareUrl = !!status?.configuredShareUrl;

  return (
    <div className="rounded-lg border border-zinc-700 bg-zinc-800/50 p-6">
      <h2 className="text-lg font-semibold text-zinc-100 mb-4">
        OneDrive Connection
      </h2>

      <div className="space-y-3">
        {/* Sharing link status (primary method) */}
        <div className="flex items-center gap-3">
          <div
            className={`h-3 w-3 rounded-full ${
              hasShareUrl ? "bg-emerald-500" : "bg-zinc-600"
            }`}
          />
          <span className="text-zinc-300 text-sm">
            {hasShareUrl
              ? "Sharing URL configured"
              : "No ONEDRIVE_SHARE_URL set"}
          </span>
          {hasShareUrl && (
            <span className="text-[10px] bg-amber-900/50 text-amber-300 px-1.5 py-0.5 rounded font-medium">
              NO SPO NEEDED
            </span>
          )}
        </div>

        {/* Graph API auth status */}
        <div className="flex items-center gap-3">
          <div
            className={`h-3 w-3 rounded-full ${
              status?.authenticated ? "bg-emerald-500" : "bg-zinc-600"
            }`}
          />
          <span className="text-zinc-300 text-sm">
            {status?.authenticated
              ? "Graph API authenticated"
              : "Graph API not connected (optional)"}
          </span>
        </div>

        {status?.configuredWorksheet && (
          <div className="flex items-center gap-3">
            <div className="h-3 w-3 rounded-full bg-emerald-500" />
            <span className="text-zinc-300 text-sm">
              Worksheet: {status.configuredWorksheet}
            </span>
          </div>
        )}

        {status?.configuredFileId && (
          <div className="flex items-center gap-3">
            <div className="h-3 w-3 rounded-full bg-emerald-500" />
            <span className="text-zinc-300 text-sm font-mono">
              File ID: {status.configuredFileId.slice(0, 20)}...
            </span>
          </div>
        )}

        {!hasShareUrl && !status?.authenticated && (
          <div className="mt-3 p-3 bg-amber-900/20 border border-amber-800/50 rounded-lg">
            <p className="text-xs text-amber-300/80 leading-relaxed">
              Recommended: Set{" "}
              <code className="bg-zinc-900 px-1 rounded">ONEDRIVE_SHARE_URL</code>{" "}
              with your colleague&apos;s sharing link. This bypasses the SPO license
              requirement entirely.
            </p>
          </div>
        )}

        {!status?.authenticated && authUrl && (
          <a
            href={authUrl}
            className="inline-block mt-3 px-4 py-2 bg-zinc-700 hover:bg-zinc-600 text-white text-sm font-medium rounded-lg transition-colors"
          >
            Connect via Graph API (optional)
          </a>
        )}

        <button
          onClick={checkStatus}
          className="mt-2 text-sm text-zinc-400 hover:text-zinc-200 underline"
        >
          Refresh status
        </button>
      </div>
    </div>
  );
}
