import fs from "node:fs";
import os from "node:os";
import path from "node:path";
import { PublicClientApplication } from "@azure/msal-node";

const SCOPES = [
  "openid",
  "profile",
  "offline_access",
  "User.Read",
  "Mail.ReadWrite",
  "Mail.Send",
  "Calendars.Read",
];

function requiredEnv(name) {
  const value = process.env[name]?.trim();
  if (!value) {
    throw new Error(
      `Missing ${name}. Set it in your Cursor mcp.json env block. See README.md.`
    );
  }
  return value;
}

function cachePath() {
  const dir = path.join(os.homedir(), ".outlook-mcp");
  fs.mkdirSync(dir, { recursive: true });
  return path.join(dir, "msal-cache.json");
}

function loadCachePlugin() {
  const file = cachePath();
  return {
    beforeCacheAccess: async (context) => {
      if (fs.existsSync(file)) {
        context.tokenCache.deserialize(fs.readFileSync(file, "utf8"));
      }
    },
    afterCacheAccess: async (context) => {
      if (context.cacheHasChanged) {
        fs.writeFileSync(file, context.tokenCache.serialize(), { mode: 0o600 });
      }
    },
  };
}

let pca;
let account;

function getApp() {
  if (pca) return pca;

  const clientId = requiredEnv("MS_CLIENT_ID");
  const tenantId = process.env.MS_TENANT_ID?.trim() || "common";

  // Device-code flow uses a public client. A client secret is not used here.
  // If MS_CLIENT_SECRET is set, it is ignored for auth (kept for config compatibility).
  pca = new PublicClientApplication({
    auth: {
      clientId,
      authority: `https://login.microsoftonline.com/${tenantId}`,
    },
    cache: {
      cachePlugin: loadCachePlugin(),
    },
  });

  return pca;
}

async function pickAccount() {
  const app = getApp();
  const accounts = await app.getTokenCache().getAllAccounts();
  if (accounts.length === 0) return null;
  account = accounts[0];
  return account;
}

/**
 * Acquire a Graph access token. Uses silent refresh when possible,
 * otherwise starts device-code flow (prints URL + code to stderr).
 */
export async function getAccessToken({ forceLogin = false } = {}) {
  const app = getApp();

  if (!forceLogin) {
    const existing = account || (await pickAccount());
    if (existing) {
      try {
        const silent = await app.acquireTokenSilent({
          account: existing,
          scopes: SCOPES,
        });
        account = silent.account;
        return silent.accessToken;
      } catch {
        // Fall through to device code.
      }
    }
  }

  const result = await app.acquireTokenByDeviceCode({
    scopes: SCOPES,
    deviceCodeCallback: (response) => {
      // MCP clients surface stderr; never write auth prompts to stdout.
      console.error(response.message);
    },
  });

  account = result.account;
  return result.accessToken;
}

export async function getAuthStatus() {
  getApp();
  const existing = account || (await pickAccount());
  if (!existing) {
    return { signedIn: false };
  }

  try {
    await getAccessToken();
    return {
      signedIn: true,
      username: existing.username,
      name: existing.name,
      homeAccountId: existing.homeAccountId,
    };
  } catch (error) {
    return {
      signedIn: false,
      error: error instanceof Error ? error.message : String(error),
    };
  }
}

export async function signOut() {
  const app = getApp();
  const accounts = await app.getTokenCache().getAllAccounts();
  for (const a of accounts) {
    await app.getTokenCache().removeAccount(a);
  }
  account = null;

  const file = cachePath();
  if (fs.existsSync(file)) {
    fs.unlinkSync(file);
  }

  return { signedOut: true };
}

export { SCOPES };
