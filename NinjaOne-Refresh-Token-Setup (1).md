# NinjaOne Refresh Token — Setup Guide

This guide walks you through configuring NinjaOne and running `Get-NinjaOneRefreshToken.ps1` to obtain an OAuth 2.0 **refresh token** for your own NinjaOne tenant.

A refresh token lets your scripts and integrations get fresh access tokens without prompting you to log in again. NinjaOne access tokens expire (typically within an hour); the refresh token is the long-lived credential you store and reuse.

---

## Prerequisites

- A **NinjaOne system administrator** account in the tenant you want to authenticate against. Only system admins can create API client apps.
- **Windows** with **PowerShell 5.1 or PowerShell 7+**.
- Ability to run PowerShell **as Administrator**. The script's local listener binds to port 80 (because the redirect URI has no explicit port), and port 80 is privileged on Windows.
- The script file: `Get-NinjaOneRefreshToken.ps1`.
- Knowing your **NinjaOne login URL**:

  | Region                | Login URL                  |
  |-----------------------|----------------------------|
  | North America         | `https://app.ninjarmm.com` |
  | Europe                | `https://eu.ninjarmm.com`  |
  | Canada                | `https://ca.ninjarmm.com`  |
  | Australia / Oceania   | `https://oc.ninjarmm.com`  |

  If you log into a different URL than the ones above, contact NinjaOne support to confirm the correct base URL — do not guess.

---

## Step 1 — Create an API Client App in NinjaOne

1. Sign in to NinjaOne as a system administrator.
2. Go to **Administration → Apps → API**.
3. Open the **Client App IDs** tab and click **Add**.
4. Fill in the application configuration:
   - **Application Platform**: `Web`
   - **Name**: anything descriptive, e.g. `PowerShell Refresh Token Helper`.
   - **Redirect URIs**: `http://localhost/`
     - The trailing slash matters. Match it exactly.
   - **Scopes**: select the scopes you actually need. Common choices:
     - `monitoring` — read device/alert data
     - `management` — make changes to devices/orgs
     - `control` — run actions / scripts
     - `offline_access` — **required** to receive a refresh token
   - **Allowed Grant Types**: tick at least
     - `Authorization Code`
     - `Refresh Token`
5. Click **Save**.
6. Copy the **Client ID**.
7. Copy the **Client Secret**. The secret is shown only once — store it somewhere safe (a password manager) immediately.

---

## Step 2 — Edit the script's `$Config` block

Open `Get-NinjaOneRefreshToken.ps1` in a text editor and fill in the values at the top:

```powershell
$Config = @{
    BaseUrl       = 'https://app.ninjarmm.com'                 # your NinjaOne login URL
    TokenEndpoint = 'https://app.ninjarmm.com/ws/oauth/token'  # BaseUrl + /ws/oauth/token
    ClientId      = 'paste-client-id-here'
    ClientSecret  = 'paste-client-secret-here'
    RedirectUri   = 'http://localhost/'
    Scope         = 'monitoring management control offline_access'
    OutputFile    = $null   # or e.g. 'C:\Secrets\ninja-token.json'
}
```

Notes:
- `BaseUrl` and the prefix of `TokenEndpoint` should match.
- `RedirectUri` must match what you registered in NinjaOne **exactly**, trailing slash included.
- Trim `Scope` down to what you need — `offline_access` is the one you can't drop.
- `OutputFile` is optional. If set, the token bundle is saved DPAPI-encrypted (decryptable only by the same Windows user on the same machine).

---

## Step 3 — Run the script (as Administrator)

1. Right-click the Start menu → **Windows PowerShell (Admin)** or **Terminal (Admin)**.
2. `cd` to wherever you saved the script.
3. If PowerShell complains about execution policy on the file, unblock it once:
   ```powershell
   Unblock-File -Path .\Get-NinjaOneRefreshToken.ps1
   ```
   Or run a single bypassed session:
   ```powershell
   powershell -ExecutionPolicy Bypass -File .\Get-NinjaOneRefreshToken.ps1
   ```
4. Run it:
   ```powershell
   .\Get-NinjaOneRefreshToken.ps1
   ```

### What happens

1. The script starts a temporary local HTTP listener on `http://localhost/`.
2. It opens your default browser to NinjaOne's authorization page.
3. You sign in (if needed) and click **Authorize**.
4. NinjaOne redirects back to `localhost`, the script captures the authorization code, and the browser shows a "you can close this window" page.
5. The script POSTs the code to your `TokenEndpoint` and prints the access token, refresh token, expiry, and scope.

If you set `OutputFile`, the bundle is also written as a DPAPI-encrypted string at that path.

---

## Step 4 — Reading back a saved token bundle

```powershell
$encrypted = Get-Content 'C:\Secrets\ninja-token.json' -Raw
$secure    = ConvertTo-SecureString -String $encrypted
$plain     = [System.Net.NetworkCredential]::new('', $secure).Password
$bundle    = $plain | ConvertFrom-Json

$bundle.refresh_token
```

---

## Step 5 — Using the refresh token to get a new access token

```powershell
$body = @{
    grant_type    = 'refresh_token'
    refresh_token = $bundle.refresh_token
    client_id     = $bundle.client_id
    client_secret = '<your-client-secret>'
}

$resp = Invoke-RestMethod -Method Post `
    -Uri "$($bundle.base_url)/ws/oauth/token" `
    -Body $body `
    -ContentType 'application/x-www-form-urlencoded'

$accessToken = $resp.access_token
```

Then call the API:

```powershell
Invoke-RestMethod -Uri "$($bundle.base_url)/api/v2/devices" `
    -Headers @{ Authorization = "Bearer $accessToken" }
```

---

## Troubleshooting

| Symptom | Likely cause / fix |
|---|---|
| `Could not start the local listener on http://localhost/` | You're not running as Administrator, **or** something else is bound to port 80 (IIS, World Wide Web Publishing Service, Skype, a dev server). Re-launch PowerShell as Administrator. If port 80 is in use, either stop the offending service or switch to a port-based redirect URI: change `RedirectUri` to `http://localhost:8765/` in the script and update the same URI in NinjaOne. |
| Browser shows `redirect_uri_mismatch` or `invalid_redirect_uri`. | The redirect URI configured in NinjaOne doesn't exactly match `RedirectUri` in the script. Check the trailing slash. |
| Token response contains `access_token` but no `refresh_token`. | The `offline_access` scope is missing, **or** `Refresh Token` isn't enabled under Allowed Grant Types on the client app. Fix the client app config and re-run. |
| `invalid_client` error during token exchange. | Wrong Client ID/Secret, or you copied a stale secret. Regenerate the secret in NinjaOne if needed. |
| `invalid_grant` error. | The authorization code is single-use and short-lived — re-run the script to get a fresh one. Also check the redirect URI matches exactly between auth and token requests. |
| Browser hangs and nothing happens. | A firewall / endpoint security tool may be blocking the local listener. The listener is loopback-only, but some EDRs still intervene. |
| `state parameter mismatch`. | Don't reuse old browser tabs. Just re-run the script — it generates a fresh state every time. |

---

## Security notes

- **Treat the refresh token like a password.** Anyone with it (plus the Client ID and Secret) can act as you against the API for whatever scopes you granted.
- The `OutputFile` option uses Windows DPAPI under the current user account. If you need to share the token between users or machines, use a proper secrets manager (Azure Key Vault, AWS Secrets Manager, 1Password, etc.) rather than copying the encrypted file.
- Refresh tokens can be revoked from **Administration → Apps → API → OAuth Tokens** in NinjaOne. Revoke any tokens that are no longer needed.
- Use the narrowest scope set that gets the job done. If a script only reads device data, `monitoring offline_access` is enough — don't add `control`.
