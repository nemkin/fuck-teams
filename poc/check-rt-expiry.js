/**
 * Extracts the refresh token from the MSAL cache (written by auth-poc.js)
 * and makes a raw HTTP POST to AAD to see the full token response,
 * including refresh_token_expires_in which MSAL doesn't expose.
 *
 * Run AFTER auth-poc.js has already logged in and written msal_cache.json.
 * Run: node check-rt-expiry.js
 */

const { PublicClientApplication } = require("@azure/msal-node");
const https = require("node:https");
const fs = require("node:fs");

const TEAMS_CLIENT_ID = "1fec8e78-bce4-4aaf-ab1b-5451cc387264";
const TENANT_ID = "6a3548ab-7570-4271-91a8-58da00697029";
const SCOPES = ["offline_access", "https://teams.microsoft.com/.default"];
const CACHE_FILE = "./msal_cache.json";

// ── MSAL setup with file-backed cache ─────────────────────────────────────────

const beforeCacheAccess = async (cacheContext) => {
  if (fs.existsSync(CACHE_FILE)) {
    cacheContext.tokenCache.deserialize(fs.readFileSync(CACHE_FILE, "utf8"));
  }
};

const afterCacheAccess = async (cacheContext) => {
  if (cacheContext.cacheHasChanged) {
    fs.writeFileSync(CACHE_FILE, cacheContext.tokenCache.serialize());
  }
};

const pca = new PublicClientApplication({
  auth: { clientId: TEAMS_CLIENT_ID, authority: `https://login.microsoftonline.com/${TENANT_ID}` },
  cache: { cachePlugin: { beforeCacheAccess, afterCacheAccess } },
});

// ── Raw HTTP POST helper ───────────────────────────────────────────────────────

function postForm(url, params) {
  return new Promise((resolve, reject) => {
    const body = new URLSearchParams(params).toString();
    const options = {
      method: "POST",
      headers: {
        "Content-Type": "application/x-www-form-urlencoded",
        "Content-Length": Buffer.byteLength(body),
      },
    };
    const req = https.request(url, options, (res) => {
      let data = "";
      res.on("data", (chunk) => (data += chunk));
      res.on("end", () => {
        try { resolve({ status: res.statusCode, body: JSON.parse(data) }); }
        catch { resolve({ status: res.statusCode, body: data }); }
      });
    });
    req.on("error", reject);
    req.write(body);
    req.end();
  });
}

function formatSeconds(s) {
  if (!s) return "not provided";
  const days = Math.floor(s / 86400);
  const hours = Math.floor((s % 86400) / 3600);
  return `${s}s = ${days}d ${hours}h`;
}

// ── Main ──────────────────────────────────────────────────────────────────────

async function main() {
  // Load cache from file (written by auth-poc.js if you add cache persistence,
  // or we re-read from MSAL's in-memory cache after a silent refresh)
  let refreshToken;

  if (fs.existsSync(CACHE_FILE)) {
    console.log("Loading token cache from", CACHE_FILE);
    const cache = JSON.parse(fs.readFileSync(CACHE_FILE, "utf8"));
    const rtEntries = Object.values(cache.RefreshToken || {});
    if (rtEntries.length === 0) {
      console.error("No refresh token in cache file. Run auth-poc.js first.");
      process.exit(1);
    }
    refreshToken = rtEntries[0].secret;
    console.log("Refresh token found in cache file.\n");
  } else {
    // No cache file — do a fresh interactive login via device code
    console.log("No cache file found. Doing fresh device code login...\n");
    await pca.acquireTokenByDeviceCode({
      scopes: SCOPES,
      deviceCodeCallback: (r) => {
        console.log("----------------------------------------------");
        console.log(r.message);
        console.log("----------------------------------------------\n");
      },
    });

    const cacheData = JSON.parse(await pca.getTokenCache().serialize());
    const rtEntries = Object.values(cacheData.RefreshToken || {});
    if (rtEntries.length === 0) {
      console.error("Still no refresh token after login. IdP is not issuing one.");
      process.exit(1);
    }
    refreshToken = rtEntries[0].secret;
    // Save for next time
    fs.writeFileSync(CACHE_FILE, await pca.getTokenCache().serialize());
    console.log("Logged in. Cache saved to", CACHE_FILE, "\n");
  }

  // ── Raw refresh token exchange ──────────────────────────────────────────────
  console.log("=== RAW TOKEN REFRESH REQUEST ===\n");
  console.log("POSTing to AAD token endpoint with grant_type=refresh_token...\n");

  const tokenUrl = `https://login.microsoftonline.com/${TENANT_ID}/oauth2/v2.0/token`;
  const { status, body } = await postForm(tokenUrl, {
    client_id: TEAMS_CLIENT_ID,
    grant_type: "refresh_token",
    refresh_token: refreshToken,
    scope: SCOPES.join(" "),
  });

  console.log("HTTP status:", status);
  console.log("\n=== FULL RAW RESPONSE ===\n");

  if (body.error) {
    console.error("Error:", body.error);
    console.error("Description:", body.error_description);
    process.exit(1);
  }

  // Print everything except the actual token values (security)
  for (const [key, value] of Object.entries(body)) {
    if (["access_token", "refresh_token", "id_token"].includes(key)) {
      console.log(`  ${key}: <present, ${String(value).length} chars>`);
    } else {
      console.log(`  ${key}:`, value);
    }
  }

  console.log("\n=== TOKEN LIFETIME SUMMARY ===\n");
  console.log("access_token  expires_in         :", formatSeconds(body.expires_in));
  console.log("access_token  ext_expires_in     :", formatSeconds(body.ext_expires_in));
  console.log("refresh_token refresh_token_expires_in:", formatSeconds(body.refresh_token_expires_in));

  if (!body.refresh_token_expires_in) {
    console.log("\nNOTE: refresh_token_expires_in not returned.");
    console.log("This usually means it's a sliding-window token (Microsoft default: 90 days).");
    console.log("The token stays alive as long as it's used at least once every 90 days.");
  } else {
    const days = Math.floor(body.refresh_token_expires_in / 86400);
    console.log(`\nRefresh token is valid for ${days} days from issuance.`);
  }

  // Save the new refresh token to cache for future runs
  if (body.refresh_token) {
    const existingCache = JSON.parse(fs.existsSync(CACHE_FILE)
      ? fs.readFileSync(CACHE_FILE, "utf8")
      : "{}");
    const rtKey = Object.keys(existingCache.RefreshToken || {})[0];
    if (rtKey) {
      existingCache.RefreshToken[rtKey].secret = body.refresh_token;
      fs.writeFileSync(CACHE_FILE, JSON.stringify(existingCache, null, 2));
      console.log("\nUpdated refresh token saved to cache.");
    }
  }
}

main().catch(console.error);
