// lib/msal.ts
import { PublicClientApplication, Configuration, LogLevel } from "@azure/msal-browser";

// ─── MSAL Configuration ──────────────────────────────────────────────
// Authority "/organizations" → the Microsoft login page lets the user
// pick ANY work/school tenant they belong to (Enfrasys, client guest, etc.)
// This is how the user "switches" into the client tenant (Scenario 2 / 3).
const msalConfig: Configuration = {
    auth: {
        clientId: process.env.NEXT_PUBLIC_AZURE_CLIENT_ID || "",
        authority: "https://login.microsoftonline.com/organizations",
        redirectUri:
            typeof window !== "undefined"
                ? `${window.location.origin}/blank.html`   // ← lightweight redirect target
                : "/blank.html",
    },
    cache: {
        cacheLocation: "sessionStorage",
    },
    system: {
        loggerOptions: {
            logLevel: LogLevel.Error,
            loggerCallback: () => { },
            piiLoggingEnabled: false,
        },
    },
};

// Single shared MSAL instance
export const msalInstance = new PublicClientApplication(msalConfig);

// The ONLY scope we need: read Azure Resource Manager on the user's behalf.
// Low-risk delegated permission → user can self-consent (no client admin)
// in tenants that allow user consent (the default).
export const ARM_SCOPES = ["https://management.azure.com/user_impersonation"];

// Ensure MSAL is initialized once (MSAL v3+ requires explicit init).
let initialized = false;
export async function ensureMsalInitialized() {
    if (!initialized) {
        await msalInstance.initialize();
        initialized = true;
    }
}
