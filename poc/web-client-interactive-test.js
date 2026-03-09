/**
 * Tests acquireTokenInteractive (PKCE auth code flow) with the Teams WEB
 * client ID. Opens a system browser, handles auth via loopback redirect.
 *
 * This is the same flow MSAL.js uses in the browser — no client secret needed.
 *
 * Run: node web-client-interactive-test.js
 */

const { PublicClientApplication } = require("@azure/msal-node");
const { exec } = require("node:child_process");
const fs = require("node:fs");

const TEAMS_WEB_CLIENT_ID = "5e3ce6c0-2b1f-4285-8d4b-75ee78787346";
const TENANT_ID           = "6a3548ab-7570-4271-91a8-58da00697029";
const CACHE_FILE          = "./msal_cache_web.json";

const SCOPES = [
  "offline_access",
  "https://api.spaces.skype.com/.default",
];

const beforeCacheAccess = async (ctx) => {
  if (fs.existsSync(CACHE_FILE))
    ctx.tokenCache.deserialize(fs.readFileSync(CACHE_FILE, "utf8"));
};
const afterCacheAccess = async (ctx) => {
  if (ctx.cacheHasChanged)
    fs.writeFileSync(CACHE_FILE, ctx.tokenCache.serialize());
};

const pca = new PublicClientApplication({
  auth: {
    clientId: TEAMS_WEB_CLIENT_ID,
    authority: `https://login.microsoftonline.com/${TENANT_ID}`,
  },
  cache: { cachePlugin: { beforeCacheAccess, afterCacheAccess } },
});

function decodeJwt(token) {
  try {
    return JSON.parse(Buffer.from(token.split(".")[1], "base64url").toString());
  } catch { return null; }
}

async function main() {
  console.log("=== Teams Web Client — Interactive Auth (PKCE) ===\n");
  console.log("A browser window will open. Complete the BME login.");
  console.log("After login, you will see a success page in the browser.\n");

  let result;
  try {
    result = await pca.acquireTokenInteractive({
      scopes: SCOPES,
      openBrowser: (url) => {
        console.log("Opening browser for:", url.split("?")[0]);
        exec(`xdg-open "${url}"`);
      },
      successTemplate: `
        <h2 style="font-family:sans-serif;color:green">
          Authentication successful — you can close this tab.
        </h2>`,
      errorTemplate: `
        <h2 style="font-family:sans-serif;color:red">
          Authentication failed: {errorMessage}
        </h2>`,
    });
  } catch (err) {
    console.error("\nInteractive auth failed:", err.message);
    if (err.message.includes("redirect_uri")) {
      console.log("\nThe web client may not allow localhost redirect URIs.");
      console.log("We need a different injection strategy.");
    }
    process.exit(1);
  }

  console.log("\n=== SUCCESS ===\n");

  const decoded = decodeJwt(result.accessToken);
  if (decoded) {
    const now = Math.floor(Date.now() / 1000);
    console.log("AT client (appid) :", decoded.appid || decoded.azp);
    console.log("AT audience       :", decoded.aud);
    console.log("AT upn            :", decoded.upn);
    console.log("AT expires in     :", Math.floor((decoded.exp - now) / 60), "min");
  }

  const cache = JSON.parse(pca.getTokenCache().serialize());
  const rtEntries = Object.values(cache.RefreshToken || {});
  console.log("\nRefresh token issued :", rtEntries.length > 0 ? "YES ✓" : "NO ✗");

  if (rtEntries.length > 0) {
    const rt = rtEntries[0];
    console.log("RT client_id         :", rt.client_id);
    console.log("RT family_id         :", rt.familyId ?? "(none)");
    console.log("Cache saved to       :", CACHE_FILE);
    console.log("\nNext step: build the localStorage injection script.");
  }
}

main().catch(console.error);
