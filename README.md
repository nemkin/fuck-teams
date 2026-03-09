# fuck-teams

A fork of [teams-for-linux](https://github.com/IsmaelMartinez/teams-for-linux) with a vibe coded fix for the daily Shibboleth re-authentication problem at BME (Budapest University of Technology and Economics) — and likely any other university or enterprise using federated AAD login with short session policies.

## The Problem

BME uses Shibboleth (SAML) as its identity provider, federated into Microsoft AAD. On Linux, Teams requires re-authentication every day. On Android, it stays logged in for months.

The difference is not magical — it comes down to how Chromium (and Electron) handles session cookies.

When you log into Teams, Microsoft AAD sets a cookie called `ESTSAUTH` on `login.microsoftonline.com`. This is the AAD session cookie. If it is present and valid, AAD will silently re-authenticate MSAL.js (Teams web's auth library) without redirecting to Shibboleth. If it is missing or expired, AAD redirects to Shibboleth, which asks for your BME username and password again.

The `ESTSAUTH` cookie is issued as a **session cookie** — it has no `Max-Age` or `Expires` attribute. Chromium's standard behavior, even in a persistent profile, is to clear session cookies when the browser (or app) closes. So every time you close teams-for-linux, the ESTSAUTH cookie is wiped from disk. Next morning, no cookie → Shibboleth login.

The Android app stays logged in because it does not use a browser-based cookie session at all — it uses a native OAuth flow with a long-lived refresh token stored in the Android Account Manager.

## The Fix

In `app/mainAppWindow/index.js`, we intercept every `Set-Cookie` response header from `login.microsoftonline.com` using Electron's `session.webRequest.onHeadersReceived`. For any auth-related cookie (`ESTSAUTH`, `ESTSAUTHPERSISTENT`, `ESTSAUTHLIGHT`, `FedAuth`, etc.), we strip any existing `Max-Age`/`Expires` attributes and replace them with `Max-Age=7776000` (90 days). Electron then stores the cookie persistently on disk instead of treating it as a session cookie.

On next startup, the cookie is still there. AAD sees it, validates it, and silently re-authenticates Teams without touching Shibboleth.

The relevant functions are `persistAuthCookie()` and the modified `onHeadersReceivedHandler()`.

## What Was Investigated First

Before landing on this fix, several other approaches were explored and ruled out:

**Upstream PR #2311** (`fix/auth-recovery-stale-session-2296`) addresses the same symptom but in the opposite direction: it detects when auth has gone stale and triggers a clean re-login. It makes the daily login less painful but does not prevent it. This fork is based on that branch.

**MSAL-node native auth flow** — the Teams desktop client ID (`1fec8e78-bce4-4aaf-ab1b-5451cc387264`) does produce a 90-day sliding-window refresh token from BME's AAD tenant, confirming that long-lived tokens are possible. However, injecting these into the Teams web session proved impossible:
- The Teams web client ID (`5e3ce6c0-2b1f-4285-8d4b-75ee78787346`) does not support device code or interactive public client flows
- FOCI (Family of Client IDs) cross-client token exchange was blocked by BME's tenant for the web client
- The web client does not have `http://localhost` registered as a redirect URI

The cookie lifetime extension turned out to be the simpler and more direct fix.

## What Could Go Wrong

**The server-side AAD session may have a hard 24-hour expiry.**
The `ESTSAUTH` cookie contains an encrypted session identifier. If BME's AAD tenant is configured to expire server-side sessions after 24 hours regardless of cookie presence, AAD will return `login_required` even when the cookie is present and not expired on the client side. In this case the fix will not help and the problem is purely a server-side policy enforced by BME's IT department.

**BME may change their session policy.**
If BME tightens their Conditional Access policies (shorter session lifetimes, require fresh MFA, etc.), this fix stops working even if it works today.

**The fix only applies to the main Teams window.**
`onHeadersReceivedHandler` is registered on `window.webContents.session`, which covers the main app window. If any hidden auth windows or iframes use a different session, those cookies would not be intercepted. In practice this has not been an issue but is worth knowing.

**This is a fork, not a patch to upstream.**
The upstream teams-for-linux project is actively maintained. This fork will diverge over time. Merging upstream changes requires manual rebasing.

## Building

```bash
npm install
npm run dist:linux:deb   # produces dist/teams-for-linux_*.deb
```

Install:
```bash
sudo dpkg -i dist/teams-for-linux_*_amd64.deb
```

## Verifying the Fix

The `[AUTH_PERSIST]` log messages are at `debug` level and won't appear in normal runs. To see them, enable debug logging via the app config file, clear your session, and run from source:

```bash
mkdir -p ~/.config/Electron
cat > ~/.config/Electron/config.json << 'EOF'
{
  "logConfig": {
    "transports": {
      "console": { "level": "debug" },
      "file": { "level": false }
    }
  }
}
EOF
rm -rf ~/.config/Electron/Partitions/
npm start
```

On login you should see:

```
[AUTH_PERSIST] Stamped 90-day Max-Age on cookie: ESTSAUTH
[AUTH_PERSIST] Stamped 90-day Max-Age on cookie: ESTSAUTHPERSISTENT
[AUTH_PERSIST] Stamped 90-day Max-Age on cookie: ESTSAUTHLIGHT
```

Then close the app and reopen it the next day — if Teams loads without a Shibboleth redirect, the fix is working.

## Credits

- [IsmaelMartinez/teams-for-linux](https://github.com/IsmaelMartinez/teams-for-linux) — the upstream project this is based on
- PR #2311 by @IsmaelMartinez — auth recovery work that this branch is based on
