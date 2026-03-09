/**
 * Tests whether we can authenticate directly with the Teams WEB client ID
 * using msal-node's device code flow.
 *
 * If this works, we can inject the resulting RT directly into MSAL.js
 * localStorage with matching client ID — no FOCI needed.
 *
 * Run: node web-client-auth-test.js
 */

const { PublicClientApplication } = require("@azure/msal-node");
const https = require("node:https");
const fs = require("node:fs");

const TEAMS_WEB_CLIENT_ID = "5e3ce6c0-2b1f-4285-8d4b-75ee78787346";
const TENANT_ID           = "6a3548ab-7570-4271-91a8-58da00697029";
const CACHE_FILE          = "./msal_cache_web.json";

// The primary scope Teams web actually uses
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

function postForm(url, params) {
  return new Promise((resolve, reject) => {
    const body = new URLSearchParams(params).toString();
    const req = https.request(url, {
      method: "POST",
      headers: { "Content-Type": "application/x-www-form-urlencoded", "Content-Length": Buffer.byteLength(body) },
    }, (res) => {
      let data = "";
      res.on("data", (c) => (data += c));
      res.on("end", () => { try { resolve({ status: res.statusCode, body: JSON.parse(data) }); } catch { resolve({ status: res.statusCode, body: data }); } });
    });
    req.on("error", reject);
    req.write(body);
    req.end();
  });
}

async function main() {
  console.log("=== Teams Web Client Auth Test ===\n");
  console.log("Client ID:", TEAMS_WEB_CLIENT_ID);
  console.log("Scopes   :", SCOPES.join(", "));
  console.log();

  let result;
  try {
    result = await pca.acquireTokenByDeviceCode({
      scopes: SCOPES,
      deviceCodeCallback: (r) => {
        console.log("----------------------------------------------");
        console.log(r.message);
        console.log("----------------------------------------------\n");
      },
    });
  } catch (err) {
    console.error("Device code flow failed:", err.message);

    if (err.message.includes("client_secret") || err.message.includes("client_assertion")) {
      console.log("\nThis client requires a secret → it is a confidential client.");
      console.log("Device code flow is not available. Trying interactive flow next...");
    } else if (err.message.includes("AADSTS7000218")) {
      console.log("\nClient does not support public client flows.");
    } else {
      console.log("\nUnexpected error — see above.");
    }
    process.exit(1);
  }

  console.log("=== SUCCESS ===\n");
  console.log("Token type    :", result.tokenType);
  console.log("Scopes granted:", result.scopes.join(", "));

  const decoded = decodeJwt(result.accessToken);
  if (decoded) {
    const now = Math.floor(Date.now() / 1000);
    console.log("AT client (appid):", decoded.appid || decoded.azp);
    console.log("AT audience      :", decoded.aud);
    console.log("AT upn           :", decoded.upn);
    console.log("AT expires in    :", Math.floor((decoded.exp - now) / 60), "min");
  }

  // Check RT was issued
  const cacheRaw = pca.getTokenCache().serialize();
  const cache = JSON.parse(cacheRaw);
  const rtEntries = Object.values(cache.RefreshToken || {});
  console.log("\nRefresh token issued:", rtEntries.length > 0 ? "YES ✓" : "NO ✗");

  if (rtEntries.length > 0) {
    const rt = rtEntries[0];
    console.log("RT client_id     :", rt.client_id);
    console.log("RT family_id     :", rt.familyId ?? "(none — not FOCI)");

    // Immediately test raw refresh to confirm RT works and get expiry info
    console.log("\n=== Testing raw RT exchange ===\n");
    const tokenUrl = `https://login.microsoftonline.com/${TENANT_ID}/oauth2/v2.0/token`;
    const { status, body } = await postForm(tokenUrl, {
      client_id:     TEAMS_WEB_CLIENT_ID,
      grant_type:    "refresh_token",
      refresh_token: rt.secret,
      scope:         SCOPES.join(" "),
    });
    console.log("HTTP status:", status);
    if (body.error) {
      console.error("RT exchange failed:", body.error, body.error_description);
    } else {
      console.log("RT exchange: SUCCESS");
      console.log("expires_in              :", body.expires_in, "s");
      console.log("refresh_token_expires_in:", body.refresh_token_expires_in ?? "not returned (sliding window)");
      console.log("foci                    :", body.foci ?? "(not present)");
      console.log("new RT issued           :", !!body.refresh_token);
    }

    console.log("\n=== VERDICT ===\n");
    console.log("We have a working RT for the Teams WEB client ID.");
    console.log("We can now inject the MSAL.js localStorage cache directly.");
    console.log("Cache saved to:", CACHE_FILE);
  }
}

main().catch(console.error);
