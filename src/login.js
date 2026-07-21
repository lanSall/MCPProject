#!/usr/bin/env node
/**
 * One-shot device-code login helper.
 * Usage: npm run login
 * (with MS_CLIENT_ID / MS_TENANT_ID set in the environment)
 */
import { getAccessToken, getAuthStatus } from "./auth.js";
import { getProfile } from "./graph.js";

async function main() {
  console.error("Starting Outlook device-code login...");
  await getAccessToken({ forceLogin: true });
  const status = await getAuthStatus();
  const profile = await getProfile();
  console.log(
    JSON.stringify(
      {
        ok: true,
        status,
        profile,
      },
      null,
      2
    )
  );
}

main().catch((error) => {
  console.error(error instanceof Error ? error.message : error);
  process.exit(1);
});
