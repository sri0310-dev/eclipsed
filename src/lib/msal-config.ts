import { ConfidentialClientApplication, Configuration, LogLevel } from "@azure/msal-node";

const msalConfig: Configuration = {
  auth: {
    clientId: process.env.AZURE_CLIENT_ID!,
    authority: `https://login.microsoftonline.com/${process.env.AZURE_TENANT_ID || "common"}`,
    clientSecret: process.env.AZURE_CLIENT_SECRET!,
  },
  system: {
    loggerOptions: {
      logLevel: LogLevel.Warning,
      piiLoggingEnabled: false,
      loggerCallback: (_level, message) => {
        console.log("[MSAL]", message);
      },
    },
  },
};

// Singleton MSAL client instance
let msalInstance: ConfidentialClientApplication | null = null;

export function getMsalClient(): ConfidentialClientApplication {
  if (!msalInstance) {
    msalInstance = new ConfidentialClientApplication(msalConfig);
  }
  return msalInstance;
}

// Scopes needed for OneDrive/Excel operations
// Files.ReadWrite.All is required to access files shared by others
export const GRAPH_SCOPES = [
  "https://graph.microsoft.com/Files.ReadWrite.All",
  "https://graph.microsoft.com/User.Read",
];

const CALLBACK_PATH = "/api/onedrive/auth/callback";

/**
 * Build the redirect URI dynamically from the incoming request host.
 * This way it works on both localhost:3000 and eclipsed-eight.vercel.app
 * without needing to change env vars per environment.
 */
export function getRedirectUri(host?: string): string {
  if (host) {
    const protocol = host.startsWith("localhost") ? "http" : "https";
    return `${protocol}://${host}${CALLBACK_PATH}`;
  }
  // Fallback to env var
  return process.env.AZURE_REDIRECT_URI || `http://localhost:3000${CALLBACK_PATH}`;
}
