#!/usr/bin/env node
/**
 * One-time delegated sign-in: runs the OAuth device-code flow against Entra
 * and prints the refresh token to register as a personal credential (the MCP
 * gateway's /me page, field x-ms-refresh-token) or to set as MS_REFRESH_TOKEN.
 *
 *   node scripts/device-login.mjs --tenant <tenant-id> --client <public-client-id>
 *
 * The app registration must be a PUBLIC client ("Allow public client flows")
 * with the delegated scopes below. The refresh token is a secret: it acts as
 * you. It stays valid ~90 days past its last use; re-run this script when it
 * expires.
 */

const SCOPES = "openid profile offline_access Tasks.ReadWrite Group.Read.All User.ReadBasic.All";

function arg(name) {
  const idx = process.argv.indexOf(name);
  return idx === -1 ? undefined : process.argv[idx + 1];
}

const tenant = arg("--tenant") ?? process.env.MS_TENANT_ID;
const client = arg("--client") ?? process.env.MS_CLIENT_ID;
if (!tenant || !client) {
  console.error("Usage: node scripts/device-login.mjs --tenant <tenant-id> --client <public-client-id>");
  process.exit(1);
}

const base = `https://login.microsoftonline.com/${tenant}/oauth2/v2.0`;

const dc = await (
  await fetch(`${base}/devicecode`, {
    method: "POST",
    headers: { "content-type": "application/x-www-form-urlencoded" },
    body: new URLSearchParams({ client_id: client, scope: SCOPES }),
  })
).json();
if (!dc.device_code) {
  console.error("Device-code request failed:", dc.error_description ?? JSON.stringify(dc));
  process.exit(1);
}

console.log(`\n${dc.message}\n`);

const startedAt = Date.now();
const intervalMs = (dc.interval ?? 5) * 1000;
for (;;) {
  if (Date.now() - startedAt > (dc.expires_in ?? 900) * 1000) {
    console.error("Timed out waiting for sign-in — run the script again.");
    process.exit(1);
  }
  await new Promise((r) => setTimeout(r, intervalMs));
  const token = await (
    await fetch(`${base}/token`, {
      method: "POST",
      headers: { "content-type": "application/x-www-form-urlencoded" },
      body: new URLSearchParams({
        grant_type: "urn:ietf:params:oauth:grant-type:device_code",
        client_id: client,
        device_code: dc.device_code,
      }),
    })
  ).json();
  if (token.error === "authorization_pending") continue;
  if (token.error) {
    console.error("Sign-in failed:", token.error_description ?? token.error);
    process.exit(1);
  }
  console.log("Signed in. Register this as your personal credential (field x-ms-refresh-token):\n");
  console.log(token.refresh_token);
  console.log("\nTreat it like a password — it acts as you. Valid ~90 days past last use.");
  process.exit(0);
}
