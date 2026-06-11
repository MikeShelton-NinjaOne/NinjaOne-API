# Tenable.sc → NinjaOne — Vulnerability CSV Uploader

This script connects to your **Tenable.sc** (Security Center) instance, pulls vulnerability scan data, formats it as a CSV file, and uploads it to a device group in **NinjaOne** as a document — all in one run.

After the first time you log in to NinjaOne, the script saves a refresh token so it can renew its own access automatically. No browser login required on subsequent runs.

---

## What Gets Pulled

Each row in the CSV represents one vulnerability found on one host and includes:

| Column | Description |
|---|---|
| Plugin ID | Tenable's unique identifier for the vulnerability check |
| Name | The name of the vulnerability |
| Severity | Informational, Low, Medium, High, or Critical |
| IP Address | The IP of the affected host |
| Hostname | The DNS name of the affected host (if available) |
| CVE | CVE identifier(s) associated with the vulnerability |
| First Seen | The date the vulnerability was first detected |
| Last Seen | The most recent date the vulnerability was detected |

---

## What You Will Need

- A **NinjaOne account** with administrator access
- A **Tenable.sc (Security Center)** account with API key access
- **Windows PowerShell** (already installed on any modern Windows PC)
- The `.ps1` script file downloaded to your computer

---

## Step 1 — Create an API App in NinjaOne

You only need to do this once.

1. Log in to your NinjaOne portal.
2. In the left sidebar click **Administration**.
3. Go to **Apps** → **API** → **Client App**.
4. Click **Add** to create a new app (or open an existing one).
5. Set the following:
   - **Name:** Any name you like, e.g. `Tenable CSV Uploader`
   - **Redirect URI:** Enter exactly the following — nothing extra:
     ```
     https://localhost
     ```
   - **Scopes / Allowed Grants:** Make sure **Refresh Token** (sometimes labeled `offline_access`) is enabled.
6. Save the app. Copy the **Client ID** and **Client Secret** — keep them somewhere safe.

> **Note:** The Client Secret is only shown once. If you lose it, generate a new one.

---

## Step 2 — Generate a Tenable.sc API Key

1. Log in to Tenable.sc.
2. Click your **username** in the top-right corner and select **Profile**.
3. Scroll down to the **API Keys** section and click **Generate**.
4. Copy both the **Access Key** and the **Secret Key** — you will need both.

> **Note:** Generating a new API key will invalidate any previously generated key for your account.

---

## Step 3 — Find Your NinjaOne Organization ID

The script uploads the CSV to a specific Organization in NinjaOne.

1. In NinjaOne, go to **Organizations** in the left sidebar.
2. Click on the Organization you want to upload to.
3. Look at the URL in your browser's address bar:
   ```
   https://app.ninjarmm.com/#/organizations/42/overview
   ```
4. The number in the URL (`42` in the example) is your **Organization ID**.

---

## Step 4 — Find Your Tenable.sc Repository ID

The script pulls vulnerabilities from a specific Repository in Tenable.sc.

1. Log in to Tenable.sc and go to **Repositories** (under the **Scans** menu or **Administration** depending on your version).
2. The **ID** column in the repository list is your **Repository ID**.

---

## Step 5 — Edit the Script

Open the `.ps1` file in a text editor. **Notepad works fine** — right-click the file and choose **Edit** or **Open with → Notepad**.

Find the configuration section near the top and fill in your values:

```powershell
# --- NinjaOne credentials ---
$NinjaBaseUrl        = 'https://<your NinjaOne login URL>'
$NinjaTokenEndpoint  = 'https://<your NinjaOne login URL>/ws/oauth/token'
$NinjaClientId       = '<Your NinjaOne Client ID>'
$NinjaClientSecret   = '<Your NinjaOne Client Secret>'
$NinjaOrganizationId = '<Your Organization ID>'

# --- Tenable.sc credentials ---
$TenableBaseUrl      = 'https://<your Tenable.sc hostname or IP>'
$TenableAccessKey    = '<Your Tenable.sc Access Key>'
$TenableSecretKey    = '<Your Tenable.sc Secret Key>'
$TenableRepositoryId = '<Your Repository ID>'
```

| Placeholder | What to put here |
|---|---|
| `https://<your NinjaOne login URL>` | The URL you use to log in to NinjaOne, e.g. `https://app.ninjarmm.com` |
| `https://<your NinjaOne login URL>/ws/oauth/token` | Same URL with `/ws/oauth/token` added to the end |
| `<Your NinjaOne Client ID>` | Client ID from Step 1 |
| `<Your NinjaOne Client Secret>` | Client Secret from Step 1 |
| `<Your Organization ID>` | The number from Step 3 |
| `https://<your Tenable.sc hostname or IP>` | The URL or IP address of your Tenable.sc server, e.g. `https://tenable.company.com` |
| `<Your Tenable.sc Access Key>` | Access Key from Step 2 |
| `<Your Tenable.sc Secret Key>` | Secret Key from Step 2 |
| `<Your Repository ID>` | The number from Step 4 |

**Example of a filled-in config block:**

```powershell
$NinjaBaseUrl        = 'https://app.ninjarmm.com'
$NinjaTokenEndpoint  = 'https://app.ninjarmm.com/ws/oauth/token'
$NinjaClientId       = 'abc123def456'
$NinjaClientSecret   = 'supersecretvalue'
$NinjaOrganizationId = '42'

$TenableBaseUrl      = 'https://tenable.company.com'
$TenableAccessKey    = 'aaabbbccc111222333'
$TenableSecretKey    = 'xxxyyyzzzaaa444555'
$TenableRepositoryId = '1'
```

Save the file when done.

---

## Step 6 — Run the Script

1. Open **PowerShell** on your computer.
   - Press the **Windows key**, type `PowerShell`, and press **Enter**.
2. Navigate to the folder where you saved the script:
   ```powershell
   cd C:\Users\YourName\Downloads
   ```
3. Run the script:
   ```powershell
   .\Tenable-To-NinjaOne.ps1
   ```

> **If you see an error about "running scripts is disabled"**, run this first and then try again:
> ```powershell
> Set-ExecutionPolicy -Scope Process -ExecutionPolicy Bypass
> ```

---

## First Run — Log In to NinjaOne

The first time you run the script, it will open your browser.

1. Log in with your NinjaOne credentials and click **Authorize**.
2. Your browser will show a **blank page** — this is normal.
3. Copy the **entire URL** from the browser address bar. It will look like:
   ```
   https://localhost/?code=abc123xyz&state=somethinglong
   ```
4. Go back to the PowerShell window, paste the URL when prompted, and press **Enter**.

The script will save a refresh token to `ninja_refresh_token.txt` in the same folder, then continue pulling data from Tenable.sc and uploading to NinjaOne.

---

## Every Run After That

No browser login needed. The script will print:

```
[NinjaOne] Found saved refresh token. Renewing access...
[NinjaOne] Access token renewed successfully.
```

Then it will go straight to pulling from Tenable.sc and uploading.

---

## What Gets Created

After the script runs you will see two new files in the same folder as the script:

| File | What it is |
|---|---|
| `Tenable-Vulnerabilities-YYYY-MM-DD.csv` | The vulnerability report, named with today's date |
| `ninja_refresh_token.txt` | The saved NinjaOne login token (keep this safe — see Security Notes) |

The CSV is also uploaded to NinjaOne and will appear as a document under the Organization you specified.

---

## A Note on Tenable.sc Certificates

If your Tenable.sc server uses a **self-signed certificate** (common for on-premises installs), PowerShell may refuse to connect with a certificate error. The script includes `-SkipCertificateCheck` to handle this automatically.

If your server has a **valid trusted certificate** from a real certificate authority, you can remove that line from the script — it is noted with a comment.

---

## Troubleshooting

**"Could not connect to Tenable.sc"**
Check that `$TenableBaseUrl` is correct and reachable from this computer. Make sure the Access Key and Secret Key are right and that the API is enabled on your Tenable.sc account. If your Tenable.sc uses a non-standard port, include it in the URL, e.g. `https://tenable.company.com:8443`.

**"No vulnerabilities returned"**
Confirm the `$TenableRepositoryId` is correct and that scans have been run and results exist in that repository. You can verify by logging in to Tenable.sc and browsing the repository directly.

**"Failed to upload CSV to NinjaOne"**
Check that `$NinjaOrganizationId` is correct and that your NinjaOne account has permission to upload documents to that organization.

**"Found saved refresh token" but then asks me to log in again**
The refresh token has expired. This is normal — just log in again and a new one will be saved automatically.

**"WARNING: No refresh token returned"**
Make sure the `offline_access` / Refresh Token grant is enabled on your NinjaOne API app under **Administration → Apps → API → Client App**.

**"The URL does not look right (state mismatch)"**
Re-run the script and copy the URL from the browser address bar immediately after the blank page appears, without navigating away.

---

## Security Notes

- **`ninja_refresh_token.txt`** grants API access to your NinjaOne account. Do not share it or commit it to source control.
- **Your Tenable.sc API keys** and **NinjaOne Client Secret** are sensitive credentials — do not share the script with those values already filled in.
- **Add these to `.gitignore`** if you are storing this script in a Git repository:
  ```
  ninja_refresh_token.txt
  Tenable-Vulnerabilities-*.csv
  ```
- To fully revoke NinjaOne access, delete `ninja_refresh_token.txt` **and** regenerate the Client Secret in NinjaOne under **Administration → Apps → API**.
- To revoke Tenable.sc access, log in to Tenable.sc and regenerate your API keys under **Profile → API Keys**.
