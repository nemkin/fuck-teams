/**
 * MSAL-node auth proof of concept.
 *
 * Uses the Teams desktop client ID + offline_access scope to do a
 * device code flow login. Prints back token lifetimes so we can see
 * whether BME's IdP gives native clients longer-lived tokens.
 *
 * Run: node auth-poc.js
 */

const { PublicClientApplication, LogLevel } = require("@azure/msal-node");

// Official Microsoft Teams desktop client ID (public, non-secret)
const TEAMS_CLIENT_ID = "1fec8e78-bce4-4aaf-ab1b-5451cc387264";

// BME tenant - "common" works for federated accounts, but you can also
// try "organizations" or the specific BME tenant ID if you know it.
const AUTHORITY = "https://login.microsoftonline.com/common";

// offline_access = ask for a refresh token
// We also request a Teams-relevant scope so the token is actually useful
const SCOPES = [
  "offline_access",
  "https://teams.microsoft.com/.default",
];

const msalConfig = {
  auth: {
    clientId: TEAMS_CLIENT_ID,
    authority: AUTHORITY,
  },
  system: {
    loggerOptions: {
      loggerCallback: (level, message) => {
        if (level === LogLevel.Error || level === LogLevel.Warning) {
          console.error("[MSAL]", message);
        }
      },
      logLevel: LogLevel.Warning,
    },
  },
};

const pca = new PublicClientApplication(msalConfig);

function decodeJwt(token) {
  try {
    const payload = token.split(".")[1];
    return JSON.parse(Buffer.from(payload, "base64url").toString("utf8"));
  } catch {
    return null;
  }
}

function formatDate(ts) {
  return new Date(ts * 1000).toISOString();
}

function printTokenInfo(label, token) {
  const decoded = decodeJwt(token);
  if (!decoded) {
    console.log(`  ${label}: (could not decode)`);
    return;
  }
  const now = Math.floor(Date.now() / 1000);
  const expiresIn = decoded.exp - now;
  const lifetimeHours = ((decoded.exp - decoded.iat) / 3600).toFixed(1);

  console.log(`\n  ${label}:`);
  console.log(`    issued  : ${formatDate(decoded.iat)}`);
  console.log(`    expires : ${formatDate(decoded.exp)}`);
  console.log(`    lifetime: ${lifetimeHours}h`);
  console.log(`    expires in ${Math.floor(expiresIn / 60)} minutes`);
  if (decoded.tid)   console.log(`    tenant  : ${decoded.tid}`);
  if (decoded.upn)   console.log(`    upn     : ${decoded.upn}`);
  if (decoded.scp)   console.log(`    scopes  : ${decoded.scp}`);
  if (decoded.amr)   console.log(`    amr     : ${decoded.amr.join(", ")}`);
}

async function main() {
  console.log("=== MSAL-node Teams Auth PoC ===\n");
  console.log("Client ID:", TEAMS_CLIENT_ID);
  console.log("Authority:", AUTHORITY);
  console.log("Scopes   :", SCOPES.join(", "));
  console.log();

  const deviceCodeRequest = {
    scopes: SCOPES,
    deviceCodeCallback: (response) => {
      // This is what the user needs to do to authenticate
      console.log("----------------------------------------------");
      console.log(response.message);
      console.log("----------------------------------------------\n");
    },
  };

  let result;
  try {
    result = await pca.acquireTokenByDeviceCode(deviceCodeRequest);
  } catch (err) {
    console.error("Auth failed:", err.message);
    process.exit(1);
  }

  console.log("\n=== TOKEN RESULT ===\n");
  console.log("Token type      :", result.tokenType);
  console.log("Scopes granted  :", result.scopes.join(", "));

  // Access token info
  if (result.accessToken) {
    printTokenInfo("Access Token", result.accessToken);
  }

  // The refresh token itself is not exposed in the result object directly,
  // but its presence is indicated and it lives in the token cache.
  // We can inspect the cache to see it.
  const cache = pca.getTokenCache();
  const cacheData = JSON.parse(await cache.serialize());

  console.log("\n=== TOKEN CACHE CONTENTS ===\n");

  const refreshTokens = Object.values(cacheData.RefreshToken || {});
  if (refreshTokens.length > 0) {
    console.log(`Refresh Tokens found: ${refreshTokens.length}`);
    for (const rt of refreshTokens) {
      console.log("\n  Refresh Token entry:");
      console.log(`    client_id   : ${rt.client_id}`);
      console.log(`    home_account: ${rt.home_account_id}`);
      console.log(`    environment : ${rt.environment}`);
      console.log(`    realm       : ${rt.realm}`);
      // last_modification_time is when it was issued
      if (rt.last_modification_time) {
        const issued = parseInt(rt.last_modification_time, 10);
        console.log(`    issued at   : ${formatDate(issued)}`);
      }
      // MSAL doesn't expose refresh token expiry directly in cache,
      // but we can note the target scopes
      console.log(`    target      : ${rt.target}`);
      console.log(`    has secret  : ${rt.secret ? "YES (token present)" : "NO"}`);
    }
  } else {
    console.log("WARNING: No refresh token in cache!");
    console.log("This means offline_access was not granted.");
    console.log("This is the key signal - if no RT, the IdP won't allow silent refresh.");
  }

  const accounts = Object.values(cacheData.Account || {});
  if (accounts.length > 0) {
    console.log(`\nAccounts found: ${accounts.length}`);
    for (const acct of accounts) {
      console.log("\n  Account:");
      console.log(`    username    : ${acct.username}`);
      console.log(`    environment : ${acct.environment}`);
      console.log(`    realm       : ${acct.realm}`);
      console.log(`    authority_type: ${acct.authority_type}`);
    }
  }

  // Now try a silent refresh immediately to confirm the RT works
  console.log("\n=== TESTING SILENT REFRESH ===\n");
  const accounts2 = await pca.getAllAccounts();
  if (accounts2.length === 0) {
    console.log("No accounts cached, can't test silent refresh.");
    return;
  }

  try {
    const silentResult = await pca.acquireTokenSilent({
      scopes: SCOPES,
      account: accounts2[0],
      forceRefresh: true, // force an actual RT exchange, not just cache hit
    });
    console.log("Silent refresh: SUCCESS");
    console.log("New access token expires:", new Date(silentResult.expiresOn).toISOString());
  } catch (err) {
    console.log("Silent refresh: FAILED");
    console.log("Error:", err.message);
    console.log("This is the critical signal - if silent refresh fails, the RT approach won't work.");
  }

  console.log("\n=== DONE ===");
}

main().catch(console.error);
