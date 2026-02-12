import ConnectionStatus from "@/components/ConnectionStatus";
import SpreadsheetViewer from "@/components/SpreadsheetViewer";
import SpreadsheetWriter from "@/components/SpreadsheetWriter";
import FileSearch from "@/components/FileSearch";

export default function Home() {
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
          <div className="flex items-center gap-4">
            <span className="text-xs text-zinc-500 bg-zinc-800 px-3 py-1 rounded-full">
              Module 1: OneDrive Feeder
            </span>
          </div>
        </div>
      </header>

      {/* Main Content */}
      <main className="max-w-7xl mx-auto px-6 py-8 space-y-6">
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
                    Register an app in{" "}
                    <span className="text-blue-400">Azure Portal</span> &gt; App
                    Registrations. Set redirect URI to{" "}
                    <code className="text-xs bg-zinc-900 px-1.5 py-0.5 rounded text-zinc-300">
                      http://localhost:3000/api/onedrive/auth/callback
                    </code>
                  </span>
                </li>
                <li className="flex gap-2">
                  <span className="text-zinc-500 font-mono">2.</span>
                  <span>
                    Copy Client ID, Tenant ID, and create a Client Secret.
                    Add them to{" "}
                    <code className="text-xs bg-zinc-900 px-1.5 py-0.5 rounded text-zinc-300">
                      .env.local
                    </code>
                  </span>
                </li>
                <li className="flex gap-2">
                  <span className="text-zinc-500 font-mono">3.</span>
                  <span>
                    Grant API permissions:{" "}
                    <code className="text-xs bg-zinc-900 px-1.5 py-0.5 rounded text-zinc-300">
                      Files.ReadWrite
                    </code>{" "}
                    and{" "}
                    <code className="text-xs bg-zinc-900 px-1.5 py-0.5 rounded text-zinc-300">
                      User.Read
                    </code>{" "}
                    (Delegated)
                  </span>
                </li>
                <li className="flex gap-2">
                  <span className="text-zinc-500 font-mono">4.</span>
                  <span>
                    Click &quot;Connect to Microsoft OneDrive&quot; below, sign in, then
                    search for your master Excel file and copy its ID to{" "}
                    <code className="text-xs bg-zinc-900 px-1.5 py-0.5 rounded text-zinc-300">
                      .env.local
                    </code>
                  </span>
                </li>
              </ol>
            </div>
          </div>
        </div>

        {/* File Search */}
        <FileSearch />

        {/* Read/Write Test Panels */}
        <div className="grid grid-cols-1 xl:grid-cols-2 gap-6">
          <SpreadsheetViewer />
          <SpreadsheetWriter />
        </div>

        {/* Footer info */}
        <div className="text-center text-xs text-zinc-600 py-4 border-t border-zinc-800">
          Hectar Operations Control Tower v0.1 | OneDrive Feeder Module |
          Powered by Microsoft Graph API
        </div>
      </main>
    </div>
  );
}
