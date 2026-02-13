"use client";

import { useState } from "react";
import ConnectionStatus from "./ConnectionStatus";
import SpreadsheetViewer from "./SpreadsheetViewer";
import SpreadsheetWriter from "./SpreadsheetWriter";
import WorksheetManager from "./WorksheetManager";
import FileSearch from "./FileSearch";
import AnalyticsDashboard from "./AnalyticsDashboard";

type Tab = "feeder" | "analytics";

export default function AppShell() {
  const [activeTab, setActiveTab] = useState<Tab>("analytics");

  return (
    <div className="min-h-screen bg-zinc-950 text-zinc-100">
      {/* Header */}
      <header className="border-b border-zinc-800 bg-zinc-900/50 backdrop-blur-sm sticky top-0 z-10">
        <div className="max-w-7xl mx-auto px-6 py-4 flex items-center justify-between">
          <div>
            <h1 className="text-xl font-bold tracking-tight">
              Hectar Control Tower
            </h1>
            <p className="text-xs text-zinc-500 mt-0.5">
              Commodity Trading & Risk Management Suite
            </p>
          </div>
          <div className="flex items-center gap-3">
            <div className="flex bg-zinc-800 rounded-lg p-0.5">
              <button
                onClick={() => setActiveTab("analytics")}
                className={`px-4 py-1.5 text-sm rounded-md transition-colors ${
                  activeTab === "analytics"
                    ? "bg-blue-600 text-white font-medium"
                    : "text-zinc-400 hover:text-zinc-200"
                }`}
              >
                Analytics
              </button>
              <button
                onClick={() => setActiveTab("feeder")}
                className={`px-4 py-1.5 text-sm rounded-md transition-colors ${
                  activeTab === "feeder"
                    ? "bg-blue-600 text-white font-medium"
                    : "text-zinc-400 hover:text-zinc-200"
                }`}
              >
                OneDrive Feeder
              </button>
            </div>
          </div>
        </div>
      </header>

      {/* Main Content */}
      <main className="max-w-7xl mx-auto px-6 py-8 space-y-6">
        {activeTab === "analytics" && <AnalyticsDashboard />}

        {activeTab === "feeder" && (
          <>
            {/* Status Banner */}
            <div className="grid grid-cols-1 lg:grid-cols-3 gap-6">
              <div className="lg:col-span-1">
                <ConnectionStatus />
              </div>
              <div className="lg:col-span-2">
                <div className="rounded-lg border border-zinc-700 bg-zinc-800/50 p-6 h-full">
                  <h2 className="text-lg font-semibold text-zinc-100 mb-3">
                    Setup Guide
                  </h2>
                  <ol className="space-y-2 text-sm text-zinc-400">
                    <li className="flex gap-2">
                      <span className="text-zinc-500 font-mono">1.</span>
                      <span>
                        In{" "}
                        <span className="text-blue-400">
                          Azure App Registration
                        </span>{" "}
                        &gt; Manifest, confirm{" "}
                        <code className="text-xs bg-zinc-900 px-1.5 py-0.5 rounded text-amber-400">
                          &quot;signInAudience&quot;:
                          &quot;PersonalMicrosoftAccount&quot;
                        </code>
                      </span>
                    </li>
                    <li className="flex gap-2">
                      <span className="text-zinc-500 font-mono">2.</span>
                      <span>
                        Click{" "}
                        <span className="text-blue-400">
                          &quot;Connect to Microsoft OneDrive&quot;
                        </span>{" "}
                        and sign in with your personal Microsoft account
                      </span>
                    </li>
                    <li className="flex gap-2">
                      <span className="text-zinc-500 font-mono">3.</span>
                      <span>
                        Use{" "}
                        <span className="text-amber-400">
                          &quot;Shared With Me&quot;
                        </span>{" "}
                        to discover files and get the composite File ID
                      </span>
                    </li>
                    <li className="flex gap-2">
                      <span className="text-zinc-500 font-mono">4.</span>
                      <span>
                        Set{" "}
                        <code className="text-xs bg-zinc-900 px-1.5 py-0.5 rounded text-zinc-300">
                          ONEDRIVE_FILE_ID
                        </code>{" "}
                        in Vercel env — then all panels work automatically
                      </span>
                    </li>
                  </ol>
                </div>
              </div>
            </div>

            {/* Worksheet Manager */}
            <WorksheetManager />

            {/* File Search */}
            <FileSearch />

            {/* Read/Write Test Panels */}
            <div className="grid grid-cols-1 xl:grid-cols-2 gap-6">
              <SpreadsheetViewer />
              <SpreadsheetWriter />
            </div>
          </>
        )}

        {/* Footer info */}
        <div className="text-center text-xs text-zinc-600 py-4 border-t border-zinc-800">
          Hectar Operations Control Tower v0.1 | OneDrive Feeder Module |
          Powered by Microsoft Graph API
        </div>
      </main>
    </div>
  );
}
