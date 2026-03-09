/**
 * Tests whether our FOCI refresh token (obtained with the Teams desktop
 * client ID) can be exchanged for tokens under the Teams WEB client ID.
 *
 * If this works, we can build the full MSAL.js localStorage injection.
 *
 * Run: node foci-web-client-test.js
 */

const https = require("node:https");
const fs = require("node:fs");

const TEAMS_DESKTOP_CLIENT_ID = "1fec8e78-bce4-4aaf-ab1b-5451cc387264";
const TEAMS_WEB_CLIENT_ID     = "5e3ce6c0-2b1f-4285-8d4b-75ee78787346";
const TENANT_ID               = "6a3548ab-7570-4271-91a8-58da00697029";
const CACHE_FILE              = "./msal_cache.json";

// Scopes Teams web actually uses (from localStorage screenshot)
const WEB_SCOPES = [
  "offline_access",
  "https://api.spaces.skype.com/.default",
];

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

function decodeJwt(token) {
  try {
    return JSON.parse(Buffer.from(token.split(".")[1], "base64url").toString("utf8"));
  } catch { return null; }
}

async function main() {
  if (!fs.existsSync(CACHE_FILE)) {
    console.error("No msal_cache.json found. Run check-rt-expiry.js first.");
    process.exit(1);
  }

  const cache = JSON.parse(fs.readFileSync(CACHE_FILE, "utf8"));
  const rtEntries = Object.values(cache.RefreshToken || {});
  if (rtEntries.length === 0) {
    console.error("No refresh token in cache.");
    process.exit(1);
  }

  const refreshToken = rtEntries[0].secret;
  const tokenUrl = `https://login.microsoftonline.com/${TENANT_ID}/oauth2/v2.0/token`;

  console.log("=== FOCI Cross-Client Test ===\n");
  console.log("Using RT from client  :", TEAMS_DESKTOP_CLIENT_ID);
  console.log("Requesting token for  :", TEAMS_WEB_CLIENT_ID);
  console.log("Scopes                :", WEB_SCOPES.join(", "));
  console.log();

  const { status, body } = await postForm(tokenUrl, {
    client_id:     TEAMS_WEB_CLIENT_ID,
    grant_type:    "refresh_token",
    refresh_token: refreshToken,
    scope:         WEB_SCOPES.join(" "),
  });

  console.log("HTTP status:", status, "\n");

  if (body.error) {
    console.error("FAILED:", body.error);
    console.error(body.error_description);
    console.log("\nFOCI cross-client exchange does NOT work for this tenant.");
    console.log("We will need a different injection strategy.");
    process.exit(1);
  }

  console.log("SUCCESS - FOCI exchange worked!\n");

  // Decode the access token to confirm it's for the right client/tenant
  const decoded = decodeJwt(body.access_token);
  if (decoded) {
    console.log("Access token claims:");
    console.log("  appid / azp :", decoded.appid || decoded.azp);
    console.log("  aud         :", decoded.aud);
    console.log("  upn         :", decoded.upn);
    console.log("  tid         :", decoded.tid);
    console.log("  scp         :", decoded.scp);
    const now = Math.floor(Date.now() / 1000);
    console.log("  expires in  :", Math.floor((decoded.exp - now) / 60), "minutes");
  }

  console.log("\nNew refresh token issued:", !!body.refresh_token);
  console.log("foci field             :", body.foci ?? "(not present)");

  // ── Build the MSAL.js localStorage cache structure ─────────────────────────
  console.log("\n=== MSAL.js localStorage Cache Preview ===\n");
  console.log("This is what we would inject into Electron before loading Teams.\n");

  const homeAccountId = `${decoded?.oid ?? "UNKNOWN"}.${TENANT_ID}`;
  const environment   = "login.windows.net";
  const now           = Math.floor(Date.now() / 1000).toString();

  // Refresh token cache entry
  const rtKey = `${homeAccountId}-${environment}-refreshtoken-${TEAMS_WEB_CLIENT_ID}----`;
  const rtValue = {
    homeAccountId,
    environment,
    credentialType: "RefreshToken",
    clientId: TEAMS_WEB_CLIENT_ID,
    realm: "",
    target: "",
    secret: body.refresh_token || refreshToken,
    tokenType: "Bearer",
    cachedAt: now,
    lastModificationTime: now,
    familyId: "1",
  };

  // Access token cache entry (spaces.skype.com)
  const scopeStr = body.scope || WEB_SCOPES.join(" ");
  const atKey = `${homeAccountId}-${environment}-accesstoken-${TEAMS_WEB_CLIENT_ID}-${TENANT_ID}-${scopeStr}--`;
  const atValue = {
    homeAccountId,
    environment,
    credentialType: "AccessToken",
    clientId: TEAMS_WEB_CLIENT_ID,
    realm: TENANT_ID,
    target: scopeStr,
    secret: body.access_token,
    tokenType: "Bearer",
    cachedAt: now,
    lastModificationTime: now,
    expiresOn: (Math.floor(Date.now() / 1000) + body.expires_in).toString(),
    extendedExpiresOn: (Math.floor(Date.now() / 1000) + body.ext_expires_in).toString(),
  };

  // msal.token.keys index
  const tokenKeysValue = {
    idToken: [],
    accessToken: [atKey],
    refreshToken: [rtKey],
  };

  console.log("msal.token.keys." + TEAMS_WEB_CLIENT_ID);
  console.log(" =>", JSON.stringify(tokenKeysValue, null, 2));
  console.log("\n" + rtKey);
  console.log(" =>", JSON.stringify({ ...rtValue, secret: "<RT present>" }, null, 2));
  console.log("\n" + atKey.slice(0, 80) + "...");
  console.log(" =>", JSON.stringify({ ...atValue, secret: "<AT present>" }, null, 2));

  // Save the web-client tokens to a separate file for the injection PoC
  const injectionPayload = {
    [`msal.token.keys.${TEAMS_WEB_CLIENT_ID}`]: JSON.stringify(tokenKeysValue),
    [rtKey]: JSON.stringify(rtValue),
    [atKey]: JSON.stringify(atValue),
  };
  fs.writeFileSync("./web_client_tokens.json", JSON.stringify(injectionPayload, null, 2));
  console.log("\nSaved injection payload to web_client_tokens.json");
  console.log("\nNext step: inject this into Electron localStorage before loading teams.microsoft.com");
}

main().catch(console.error);
