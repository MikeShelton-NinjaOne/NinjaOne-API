# NinjaOne — Run a Script as SYSTEM (Refresh Token Edition)

This script lets you run a PowerShell command on a remote device through NinjaOne, running as the **SYSTEM** account. Unlike the one-time auth code version, this script **remembers your login** — after the first time you authorize it, it will renew its own access automatically every time you run it. No browser login required after the first run.

---

## How It Works (Plain English)

The first time you run the script, it opens your browser and asks you to log in to NinjaOne. After you authorize it, NinjaOne gives the script a **refresh token** — think of it like a long-lived key. The script saves that key to a small text file on your computer.

Every time you run the script after that, it uses the saved key to get a fresh short-lived access token on its own — no browser, no copy-pasting URLs. If the key ever expires or gets deleted, the script will ask you to log in again and get a new one.

---

## What You Will Need

- A **NinjaOne account** with administrator access
- **Windows PowerShell** (already installed on any modern Windows PC — no download needed)
- The `.ps1` script file from this repository downloaded to your computer

---

## Step 1 — Create an API App in NinjaOne

You only need to do this once.

1. Log in to your NinjaOne portal.
2. In the left sidebar, click **Administration**.
3. Go to **Apps** → **API** → **Client App**.
4. Click **Add** to create a new app (or open an existing one).
5. Fill in the following:
   - **Name:** Any name you like, e.g. `Run Script Tool`
   - **Redirect URI:** Enter exactly the following — do not add anything extra:
     ```
     https://localhost
     ```
   - **Scopes / Allowed Grants:** Make sure **Refresh Token** (sometimes labeled `offline_access`) is enabled. This is what allows the script to stay logged in between runs.
6. Save the app. NinjaOne will show you a **Client ID** and **Client Secret** — copy both and keep them somewhere safe.

> **Note:** The Client Secret is only shown once. If you lose it, you will need to generate a new one.

---

## Step 2 — Find Your Device ID

1. In NinjaOne, go to **Devices** and click on the device you want to target.
2. Look at the URL in your browser's address bar. It will look something like:
   ```
   https://app.ninjarmm.com/#/deviceDashboard/12345/overview
   ```
3. The number near the end (`12345` in the example) is your **Device ID**. Write it down.

---

## Step 3 — Edit the Script

Open the `.ps1` file in a text editor. **Notepad works fine** — right-click the file and choose **Edit** or **Open with → Notepad**.

Near the top you will see this section:

```powershell
$BaseUrl         = 'https://<your login URL>'
$TokenEndpoint   = 'https://<your login URL>/ws/oauth/token'
$ClientId        = '<Your Client ID>'
$ClientSecret    = '<Your Client Secret>'

$DeviceId        = '<Your Device ID>'

$TokenFile       = Join-Path $PSScriptRoot 'ninja_refresh_token.txt'
```

Replace each placeholder with your real values:

| Placeholder | What to put here |
|---|---|
| `https://<your login URL>` | The URL you use to log in to NinjaOne, e.g. `https://app.ninjarmm.com` |
| `https://<your login URL>/ws/oauth/token` | Same URL with `/ws/oauth/token` added to the end |
| `<Your Client ID>` | The Client ID from Step 1 |
| `<Your Client Secret>` | The Client Secret from Step 1 |
| `<Your Device ID>` | The number you found in Step 2 |

**`$TokenFile`** controls where the refresh token is saved. The default saves it in the same folder as the script, which is fine for most people. You can change it to a different path if you prefer — for example:
```powershell
$TokenFile = 'C:\Scripts\ninja_refresh_token.txt'
```

**Example of a filled-in config block:**

```powershell
$BaseUrl         = 'https://app.ninjarmm.com'
$TokenEndpoint   = 'https://app.ninjarmm.com/ws/oauth/token'
$ClientId        = 'abc123def456'
$ClientSecret    = 'supersecretvalue'

$DeviceId        = '12345'

$TokenFile       = Join-Path $PSScriptRoot 'ninja_refresh_token.txt'
```

---

## Step 4 — Enter the Commands You Want to Run

Find the `$ScriptBody` block just below the config section:

```powershell
$ScriptBody      = @'
Write-Output "Hello from SYSTEM context"
whoami
'@
```

Replace the example commands between `@'` and `'@` with whatever PowerShell commands you want to run on the device.

**Do not delete or modify the `@'` and `'@` lines themselves.**

Save the file when done.

---

## Step 5 — Run the Script

1. Open **PowerShell** on your computer.
   - Press the **Windows key**, type `PowerShell`, and press **Enter**.
2. Navigate to the folder where you saved the script:
   ```powershell
   cd C:\Users\YourName\Downloads
   ```
3. Run the script:
   ```powershell
   .\NinjaOne-RunScriptAsSystem-RefreshToken.ps1
   ```

> **If you see an error about "running scripts is disabled"**, run this first and then try again:
> ```powershell
> Set-ExecutionPolicy -Scope Process -ExecutionPolicy Bypass
> ```

---

## First Run — Log In and Authorize

The first time you run the script, it will open your browser.

1. Log in with your NinjaOne credentials and click **Authorize**.
2. Your browser will land on a **blank page** — this is normal.
3. Copy the **entire URL** from the browser address bar. It will look like:
   ```
   https://localhost/?code=abc123xyz&state=somethinglong
   ```
4. Go back to the PowerShell window and paste the URL when prompted, then press **Enter**.

The script will log in, save a refresh token to `ninja_refresh_token.txt`, and then run your script on the device.

---

## Every Run After That

No browser, no copy-pasting. The script will print:

```
Found saved refresh token. Attempting to renew access...
Access token renewed successfully using saved refresh token.
```

And then go straight to running your script.

---

## What is the Token File?

After the first run, you will see a new file called `ninja_refresh_token.txt` in the same folder as the script. This file contains the refresh token that keeps the script logged in.

**Keep this file safe:**
- Do not share it or upload it to GitHub — it grants access to your NinjaOne account.
- If you delete it, the script will simply ask you to log in again on the next run.
- If you think the token has been compromised, delete the file and revoke the API app in NinjaOne under **Administration → Apps → API**.

> **Tip for GitHub users:** Add `ninja_refresh_token.txt` to your `.gitignore` file so it never gets accidentally committed.

---

## Troubleshooting

**"Found saved refresh token" but then it asks me to log in again**
The refresh token has expired. This is normal — NinjaOne refresh tokens have a limited lifetime. Just log in again and a new token will be saved.

**"WARNING: NinjaOne did not return a refresh token"**
This usually means the `offline_access` scope is not enabled on your API app in NinjaOne. Go back to **Administration → Apps → API → Client App**, open your app, and make sure Refresh Token access is enabled.

**"The URL does not look right (state mismatch)"**
You may have copied the URL from the wrong tab, or the page refreshed. Re-run the script and try again immediately after the blank page appears.

**"No authorization code found in the URL"**
Make sure you copied the full URL from the address bar, not just part of it.

**"Could not get an access token"**
Double-check your `$ClientId`, `$ClientSecret`, and `$TokenEndpoint` in the script. Make sure there are no extra spaces.

**"Failed to send the script"**
Confirm the `$DeviceId` is correct and that the target device is online in NinjaOne.

---

## Security Notes

- **Protect your `ninja_refresh_token.txt` file** — anyone who has it can access your NinjaOne account via the API.
- **Protect your `$ClientSecret`** — do not share the script with the secret already filled in.
- This script runs commands as **SYSTEM**, which has full control over the local machine. Always double-check your `$ScriptBody` before running.
- To fully revoke access, delete the token file **and** delete or regenerate the Client Secret in NinjaOne under **Administration → Apps → API**.
