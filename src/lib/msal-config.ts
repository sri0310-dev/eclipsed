import { ConfidentialClientApplication, Configuration, LogLevel } from "@azure/msal-node";

const msalConfig: Configuration = {
  auth: {
    clientId: process.env.AZURE_CLIENT_ID!,
    authority: `https://login.microsoftonline.com/${process.env.AZURE_TENANT_ID}`,
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
export const GRAPH_SCOPES = [
  "https://graph.microsoft.com/Files.ReadWrite",
  "https://graph.microsoft.com/User.Read",
];

export const REDIRECT_URI = process.env.AZURE_REDIRECT_URI!;
