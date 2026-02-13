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
  const [authError, setAuthError] = useState<string | null>(null);

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
      setAuthError(params.get("auth_error"));
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

  return (
    <div className="rounded-lg border border-zinc-700 bg-zinc-800/50 p-6">
      <h2 className="text-lg font-semibold text-zinc-100 mb-4">
        OneDrive Connection
      </h2>

      <div className="space-y-3">
        {/* Graph API auth status */}
        <div className="flex items-center gap-3">
          <div
            className={`h-3 w-3 rounded-full ${
              status?.authenticated ? "bg-emerald-500" : "bg-red-500"
            }`}
          />
          <span className="text-zinc-300">
            {status?.authenticated
              ? "Connected to Microsoft Graph"
              : "Not connected"}
          </span>
        </div>

        {status?.configuredShareUrl && (
          <div className="flex items-center gap-3">
            <div className="h-3 w-3 rounded-full bg-emerald-500" />
            <span className="text-zinc-300 text-sm">
              Share URL configured
            </span>
          </div>
        )}

        {status?.configuredWorksheet && (
          <div className="flex items-center gap-3">
            <div className="h-3 w-3 rounded-full bg-emerald-500" />
            <span className="text-zinc-300 text-sm">
              Default worksheet: {status.configuredWorksheet}
            </span>
          </div>
        )}

        {!status?.authenticated && authUrl && (
          <a
            href={authUrl}
            className="inline-block mt-3 px-4 py-2 bg-blue-600 hover:bg-blue-700 text-white text-sm font-medium rounded-lg transition-colors"
          >
            Connect to Microsoft OneDrive
          </a>
        )}

        {authError && (
          <div className="mt-3 p-3 bg-red-900/30 border border-red-700 rounded-lg text-sm text-red-300 space-y-2">
            <div className="font-medium">Connection failed</div>
            <div className="text-xs text-red-400">{authError}</div>
            {authError.includes("70000") && (
              <div className="text-xs text-amber-400 mt-2 p-2 bg-amber-900/20 border border-amber-800 rounded">
                <strong>Fix:</strong> In Azure Portal &gt; App Registration
                &gt; Manifest, change{" "}
                <code className="bg-zinc-900 px-1 rounded">
                  &quot;signInAudience&quot;
                </code>{" "}
                from{" "}
                <code className="bg-zinc-900 px-1 rounded text-red-300">
                  &quot;AzureADMyOrg&quot;
                </code>{" "}
                to{" "}
                <code className="bg-zinc-900 px-1 rounded text-emerald-300">
                  &quot;PersonalMicrosoftAccount&quot;
                </code>
                , then save and redeploy.
              </div>
            )}
          </div>
        )}

        {status?.authenticated && (
          <button
            onClick={checkStatus}
            className="mt-2 text-sm text-zinc-400 hover:text-zinc-200 underline"
          >
            Refresh status
          </button>
        )}
      </div>
    </div>
  );
}
