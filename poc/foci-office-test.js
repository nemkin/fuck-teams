/**
 * Tests FOCI exchange using the Microsoft Office client ID.
 * Office is a well-known FOCI member — if Teams web is also FOCI,
 * this RT (obtained via Office device code) will exchange across.
 *
 * Run: node foci-office-test.js
 */

const { PublicClientApplication } = require("@azure/msal-node");
const https = require("node:https");
const fs = require("node:fs");

const OFFICE_CLIENT_ID    = "d3590ed6-52b3-4102-aeff-aad2292ab01c";
const TEAMS_WEB_CLIENT_ID = "5e3ce6c0-2b1f-4285-8d4b-75ee78787346";
const TENANT_ID           = "6a3548ab-7570-4271-91a8-58da00697029";
const CACHE_FILE          = "./msal_cache_office.json";

const SCOPES = ["offline_access", "https://api.spaces.skype.com/.default"];

const beforeCacheAccess = async (ctx) => {
  if (fs.existsSync(CACHE_FILE))
    ctx.tokenCache.deserialize(fs.readFileSync(CACHE_FILE, "utf8"));
};
const afterCacheAccess = async (ctx) => {
  if (ctx.cacheHasChanged)
    fs.writeFileSync(CACHE_FILE, ctx.tokenCache.serialize());
};

const pca = new PublicClientApplication({
  auth: { clientId: OFFICE_CLIENT_ID, authority: `https://login.microsoftonline.com/${TENANT_ID}` },
  cache: { cachePlugin: { beforeCacheAccess, afterCacheAccess } },
});

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
  console.log("Step 1: Get a fresh FOCI RT via the Office client (device code)...\n");

  let officeRT;

  if (fs.existsSync(CACHE_FILE)) {
    const cache = JSON.parse(fs.readFileSync(CACHE_FILE, "utf8"));
    const rtEntries = Object.values(cache.RefreshToken || {});
    if (rtEntries.length > 0) {
      officeRT = rtEntries[0].secret;
      console.log("Found existing Office RT in cache, using it.\n");
    }
  }

  if (!officeRT) {
    await pca.acquireTokenByDeviceCode({
      scopes: SCOPES,
      deviceCodeCallback: (r) => {
        console.log("----------------------------------------------");
        console.log(r.message);
        console.log("----------------------------------------------\n");
      },
    });
    const cache = JSON.parse(pca.getTokenCache().serialize());
    const rtEntries = Object.values(cache.RefreshToken || {});
    officeRT = rtEntries[0]?.secret;
    if (!officeRT) { console.error("No RT issued for Office client. Giving up."); process.exit(1); }
    console.log("foci field on Office RT:", rtEntries[0]?.familyId ?? "(check raw response)");
  }

  // Verify it's a FOCI token by checking the raw response
  console.log("Step 2: Confirm FOCI flag on Office RT via raw exchange...");
  const tokenUrl = `https://login.microsoftonline.com/${TENANT_ID}/oauth2/v2.0/token`;
  const checkResp = await postForm(tokenUrl, {
    client_id: OFFICE_CLIENT_ID,
    grant_type: "refresh_token",
    refresh_token: officeRT,
    scope: SCOPES.join(" "),
  });
  if (checkResp.body.error) {
    console.error("Office RT exchange failed:", checkResp.body.error);
    process.exit(1);
  }
  console.log("foci in Office RT response:", checkResp.body.foci ?? "(not present — not FOCI!)");
  const freshOfficeRT = checkResp.body.refresh_token || officeRT;

  console.log("\nStep 3: Try FOCI cross-client exchange → Teams web client...");
  const { status, body } = await postForm(tokenUrl, {
    client_id:     TEAMS_WEB_CLIENT_ID,
    grant_type:    "refresh_token",
    refresh_token: freshOfficeRT,
    scope:         ["offline_access", "https://api.spaces.skype.com/.default"].join(" "),
  });

  console.log("HTTP status:", status);
  if (body.error) {
    console.error("\nFailed:", body.error);
    console.error(body.error_description?.split("\r\n")[0]);
    console.log("\nConclusion: Teams web client is NOT in the FOCI family.");
    console.log("We need the Electron-internal auth intercept approach.");
  } else {
    console.log("\nSUCCESS — FOCI works via Office client!");
    console.log("Scopes:", body.scope);
    console.log("New RT:", !!body.refresh_token);
    console.log("\nWe can now build the localStorage injection using this RT.");
  }
}

main().catch(console.error);
