# Sync-InsightVM-To-NinjaOne

A PowerShell script that pulls vulnerability (CVE) findings from **Rapid7 InsightVM (Insight Platform, cloud)** using an **API key**, filters them by severity, builds a CSV in the format NinjaOne's Rapid7 Vulnerability Importer expects, and uploads it to a NinjaOne **Scan Group**.

> **Status note:** Items that could not be verified from public documentation are listed in [Section 12](#12-known-limitations-and-unverified-items). Anything marked **null** or **unverified** in this document should be confirmed in your own environment before relying on it.

---

## Table of contents

1. [What this does](#1-what-this-does)
2. [How it works (step by step)](#2-how-it-works-step-by-step)
3. [Data read and written](#3-data-read-and-written)
4. [Requirements](#4-requirements)
5. [One-time setup](#5-one-time-setup)
6. [Settings reference](#6-settings-reference)
7. [Running the script](#7-running-the-script)
8. [Scheduling it](#8-scheduling-it)
9. [Output files](#9-output-files)
10. [Reading the log](#10-reading-the-log)
11. [Troubleshooting](#11-troubleshooting)
12. [Known limitations and unverified items](#12-known-limitations-and-unverified-items)
13. [Security guidance](#13-security-guidance)
14. [Quick reference](#14-quick-reference)
15. [Supporting documents](#15-supporting-documents)

---

## 1. What this does

| Item | Detail |
|---|---|
| **Source** | Rapid7 InsightVM on the Insight Platform (cloud) |
| **Source authentication** | Rapid7 **API key** sent in the `X-Api-Key` header. No username or password. |
| **Destination** | NinjaOne Rapid7 Vulnerability Importer, into an existing Scan Group |
| **Destination authentication** | NinjaOne OAuth 2.0 **Client Credentials** (Client ID + Client Secret) |
| **What is sent** | Rows of `Vendor`, `Device ID`, `CVE ID` |
| **Filtering** | Minimum CVSS score (v3 preferred, v2 as fallback) |
| **Upload behavior** | Full current list on every run; each run is intended to replace the previous upload |
| **Logging** | Timestamped plain-English log to screen and to a text file |

### Who this is for
IT or security staff who maintain both Rapid7 InsightVM and NinjaOne and want vulnerability data in NinjaOne without manually exporting and importing a CSV.

---

## 2. How it works (step by step)

1. **Connect to Rapid7.** The script calls the regional Rapid7 API address, `https://{region}.api.insight.rapid7.com/vm/v4/integration`, with your API key.
2. **Pull assets.** It requests the asset list using a POST "search" call. Rapid7 returns results a page at a time with a **cursor**; the script keeps requesting pages until no cursor is returned. Each asset comes back with its findings attached, grouped as `new`, `same` and `remediated`.
3. **Pull vulnerability definitions.** It requests the vulnerability list the same way, once, and stores it in memory by vulnerability ID. This supplies the CVE number(s) and CVSS scores.
4. **Build the rows.** For each asset:
   - Choose the device identifier (`host_name` or `ip`, per your `$DeviceIdField` setting). Assets without that identifier are skipped.
   - Take the **current** findings (`new` + `same`). `remediated` findings are left out.
   - Look up each finding's definition. Use the CVSS v3 score, falling back to v2, falling back to 0.
   - Drop anything below `$MinimumCVSSSeverity`.
   - Split the CVE field into individual CVE IDs. Findings with no CVE are skipped, because NinjaOne's importer expects a CVE ID.
5. **De-duplicate.** Identical `Device ID` + `CVE ID` pairs are collapsed to one row.
6. **Write the CSV** (`insightvm-export.csv`).
7. **Log in to NinjaOne** and obtain a bearer token.
8. **Upload the CSV** to the Scan Group you specified.
9. **Log the result.** If anything fails, the log says what failed and what to check.

If zero rows qualify, the script logs that and **exits without uploading**.

---

## 3. Data read and written

| Direction | System | Data |
|---|---|---|
| **Read** | Rapid7 `assets` | Asset ID, host name, IP address, and each asset's findings (`new`, `same`, `remediated`) |
| **Read** | Rapid7 `vulnerabilities` | Vulnerability ID, CVE ID(s), CVSS v3 score, CVSS v2 score |
| **Write** | Local disk | `insightvm-export.csv` and `sync-log.txt` |
| **Write** | NinjaOne Scan Group | CSV rows of `Vendor`, `Device ID`, `CVE ID` |

The script only reads from Rapid7. It never changes anything in Rapid7. In NinjaOne, it only uploads to the Scan Group you specify.

---

## 4. Requirements

| Requirement | Detail |
|---|---|
| **PowerShell 7 or newer** | The NinjaOne upload uses `Invoke-RestMethod -Form`, which is not available in Windows PowerShell 5.1. Check your version with `$PSVersionTable.PSVersion`. |
| **Network access** | Outbound HTTPS to `{region}.api.insight.rapid7.com` and to your NinjaOne host (for example `app.ninjarmm.com`). |
| **Rapid7 subscription** | InsightVM on the Insight Platform (cloud) with API access. |
| **Rapid7 API key** | User key or organization key. See [5.1](#51-rapid7-create-an-api-key). |
| **NinjaOne API app** | Machine-to-machine app with Monitoring and Management scopes. See [5.2](#52-ninjaone-create-an-oauth-api-app). |
| **NinjaOne Rapid7 app** | Rapid7 Vulnerability Importer enabled, with an existing Scan Group. |

### On-premises Security Console users
If you run the **on-premises InsightVM Security Console** (not the Insight Platform), this script **will not work as written**. It is built for the cloud API and API-key authentication only.

---

## 5. One-time setup

### 5.1 Rapid7: create an API key

1. Go to <https://insight.rapid7.com> and log in.
2. Open **InsightVM**.
3. Click the **gear icon** (top right) and choose **API Keys**.
4. Click **Generate New User Key** (or create an **organization key**).
5. Select your organization, give the key a recognizable name (for example `ninjaone-sync`), and click **Submit**.
6. **Copy the key immediately.** Rapid7 only shows it once.

**Permissions.** The key acts as its user, so that user needs a role that can read asset data. Roles reported to work: **Asset Owner**, **Security Manager**, **Site Owner**, or a custom role with the right permissions and scope.

**Key type.** One integrator reported that an **organization key** connected more reliably than a user key. If a user key is rejected, try an organization key.

**Region.** Note which Rapid7 region your account is in. It must match `$Rapid7Region`.

| `$Rapid7Region` value | Region |
|---|---|
| `us` | US |
| `us2` | US2 |
| `us3` | US3 |
| `eu` | Europe |
| `ca` | Canada |
| `au` | Australia |
| `ap` | Japan / Asia-Pacific |
| `aps2` | India |
| `me1` | Middle East |

> Region codes come from Rapid7 and third-party documentation. If your region is not listed, confirm the correct code with Rapid7.

#### Optional: test the key before running the script
Replace `us` and the key with your own values:

```powershell
$headers = @{ "X-Api-Key" = "YOUR-KEY"; "Accept" = "application/json"; "Content-Type" = "application/json" }
Invoke-RestMethod -Method Post -Headers $headers -Body "{}" `
  -Uri "https://us.api.insight.rapid7.com/vm/v4/integration/assets?size=2&includeSame=true"
```

A successful call returns a `data` list and a `metadata` section. A 401 or 403 means the key, its permissions, or the region is wrong.

### 5.2 NinjaOne: create an OAuth API app

1. In NinjaOne go to **Administration > Apps > API**.
2. Open the **Client App IDs** tab and click **Add**.
3. For **Application Platform**, choose **API Services (machine-to-machine)**.
4. Give it a recognizable name.
5. Under **Scopes**, check both **Monitoring** and **Management**.
6. Under **Allowed Grant Types**, make sure **Client Credentials** is selected.
7. Click **Save**.
8. Copy the **Client ID** and **Client Secret**. The Secret is only shown once.

Notes:
- Only system administrators can create OAuth client applications.
- NinjaOne access tokens expire after one hour. The script requests a fresh token on every run, so this does not need action.

### 5.3 NinjaOne: enable the Rapid7 Vulnerability Importer

1. Go to **Administration > Apps** and find **Rapid7**.
2. Click **Enable**.

### 5.4 NinjaOne: create the Scan Group

1. Open the Rapid7 app and go to the **Scan Groups** tab.
2. Click **Create scan group**.
3. Name it (for example `Rapid7 - All Servers`) and finish the wizard. The wizard includes an attribute-mapping step, where NinjaOne reads the CSV's column headings and lets you map them.
4. Note the **Scan Group ID** and enter it in `$NinjaOneScanGroupID`.

> **Unverified:** Where the Scan Group ID is displayed in the console could not be confirmed from public documentation. If you cannot find it, ask NinjaOne support.

> **Tip:** Because the attribute-mapping step reads the headings of the file you upload, creating the Scan Group with a sample file produced by this script (headings `Vendor`, `Device ID`, `CVE ID`) helps make sure the mapping lines up.

---

## 6. Settings reference

All settings are in the **YOUR SETTINGS** block at the top of the script. Edit only that block.

| Setting | Required | Example | Description |
|---|---|---|---|
| `$Rapid7Region` | Yes | `"us"` | Your Rapid7 region code. Uncomment exactly one line. |
| `$Rapid7ApiKey` | Yes | `"abc123..."` | The Rapid7 API key. No username/password. |
| `$NinjaOneBaseURL` | Yes | `"app.ninjarmm.com"` | The NinjaOne host you log in to. Options: `app.ninjarmm.com` (US), `us2.ninjarmm.com`, `eu.ninjarmm.com`, `ca.ninjarmm.com`, `oc.ninjarmm.com`. Uncomment exactly one line. |
| `$NinjaOneClientID` | Yes | `"..."` | Client ID from the NinjaOne API app. |
| `$NinjaOneClientSecret` | Yes | `"..."` | Client Secret from the NinjaOne API app. |
| `$NinjaOneScanGroupID` | Yes | `1234` | ID of the existing Scan Group to update. The script does not create it. |
| `$MinimumCVSSSeverity` | Yes | `7.0` | Only CVEs at or above this score are sent. `10.0` = critical only, `7.0` = high and above, `0.0` = everything. |
| `$DeviceIdField` | Yes | `"Hostname"` | `"Hostname"` or `"IPAddress"`. Must match how your NinjaOne Scan Group matches devices. |
| `$LogFilePath` | Yes | `".\sync-log.txt"` | Where the log is written. |

### Choosing `$DeviceIdField`
The value here must correspond to how devices are matched in your NinjaOne Scan Group. If they don't match, NinjaOne won't be able to attach CVEs to your devices. Hostnames in Rapid7 and NinjaOne need to be spelled the same way (watch for domain suffixes and capitalization). **Whether NinjaOne's matching is case-sensitive or tolerant of domain suffixes: null (not verified).**

---

## 7. Running the script

1. Open **PowerShell 7** (`pwsh`).
2. Move to the folder containing the script:
   ```powershell
   cd C:\Scripts\InsightVM-NinjaOne
   ```
3. If scripts are blocked by your execution policy, allow it for this session only:
   ```powershell
   Set-ExecutionPolicy -Scope Process -ExecutionPolicy Bypass
   ```
4. Run it:
   ```powershell
   .\Sync-InsightVM-To-NinjaOne.ps1
   ```

### Recommended first run
1. Set `$MinimumCVSSSeverity` to your normal value.
2. Run the script and read the log.
3. Open `insightvm-export.csv` and confirm that `Device ID` and `CVE ID` are filled in and look right.
4. Check NinjaOne to confirm the data arrived in the Scan Group.

If the CSV is empty or the log reports zero rows, see [Troubleshooting](#11-troubleshooting).

---

## 8. Scheduling it

Because each run uploads the full current list, the script is suitable for a recurring schedule (for example daily). A Windows Task Scheduler example (adjust paths, time and account to your environment):

```powershell
$action  = New-ScheduledTaskAction -Execute "pwsh.exe" `
           -Argument '-NoProfile -ExecutionPolicy Bypass -File "C:\Scripts\InsightVM-NinjaOne\Sync-InsightVM-To-NinjaOne.ps1"' `
           -WorkingDirectory "C:\Scripts\InsightVM-NinjaOne"
$trigger = New-ScheduledTaskTrigger -Daily -At 3am
Register-ScheduledTask -TaskName "InsightVM to NinjaOne Sync" -Action $action -Trigger $trigger
```

Notes:
- Set the **working directory** so that the relative paths (`.\sync-log.txt`, `.\insightvm-export.csv`) land in the script's folder.
- Run the task under an account that can read the script file, and restrict access to that file (it contains secrets).
- A sensible frequency depends on how often Rapid7 scans run. **Rapid7 rate limits: null (not verified).**

---

## 9. Output files

| File | Created | Purpose |
|---|---|---|
| `sync-log.txt` (path set by `$LogFilePath`) | Every run (appended) | Timestamped record of what happened |
| `insightvm-export.csv` | When at least one row qualifies | The file uploaded to NinjaOne. Overwritten each run. |

### CSV format
```csv
"Vendor","Device ID","CVE ID"
"Rapid7","SERVER01","CVE-2024-0001"
"Rapid7","SERVER01","CVE-2024-0002"
"Rapid7","WS-FINANCE-07","CVE-2024-0001"
```

The CSV contains no secrets, but it does contain your device names and vulnerability data. Treat it as sensitive.

---

## 10. Reading the log

A normal run looks like this (numbers are illustrative):

```
[2026-11-10 03:00:01] ===== Starting InsightVM -> NinjaOne sync =====
[2026-11-10 03:00:01] Connecting to Rapid7 (region: us) and retrieving asset list...
[2026-11-10 03:00:19] Found 412 assets in InsightVM.
[2026-11-10 03:00:21] Pulled 9,850 raw findings from InsightVM. 1,204 rows remain after applying your severity filter (7+) and removing entries without a CVE ID or device identifier.
[2026-11-10 03:00:21] 1,187 unique device/CVE rows after removing duplicates.
[2026-11-10 03:00:21] Wrote CSV file to .\insightvm-export.csv.
[2026-11-10 03:00:21] Connecting to NinjaOne...
[2026-11-10 03:00:22] Connected to NinjaOne successfully.
[2026-11-10 03:00:22] Uploading CSV to Scan Group ID 1234...
[2026-11-10 03:00:25] Upload succeeded. 1,187 rows sent to NinjaOne Scan Group 1234.
[2026-11-10 03:00:25] ===== Sync finished =====
```

| Log message | Meaning |
|---|---|
| `Found N assets` | Rapid7 connection and key worked. |
| `Pulled X raw findings ... Y rows remain` | X is every current finding on devices with an identifier. Y is what's left after the severity and CVE filters. A big drop is normal. |
| `No qualifying CVEs found - nothing to upload` | Nothing met your severity filter (or fields weren't found). Nothing was uploaded. |
| `COULD NOT CONNECT TO RAPID7` | Region, API key or permissions problem. |
| `COULD NOT LOG IN TO NINJAONE` | NinjaOne host, Client ID or Client Secret problem. |
| `THE UPLOAD TO NINJAONE FAILED` | Scan Group ID, scopes, or the upload address. |
| `Sync stopped early because of an error` | A step above failed; the message above it says which. |

---

## 11. Troubleshooting

### Rapid7

| Symptom | Likely cause | What to do |
|---|---|---|
| `COULD NOT CONNECT TO RAPID7` with 401 or 403 | Key invalid, expired, missing permissions, or the wrong region | Generate a fresh key (gear icon > API Keys), confirm `$Rapid7Region`, confirm the user's role can read asset data, and try an **organization key**. |
| Connection error / timeout | Network or proxy blocking `*.api.insight.rapid7.com` | Allow outbound HTTPS to the regional address; test with the snippet in [5.1](#51-rapid7-create-an-api-key). |
| `Found 0 assets` | Key's user has no asset scope, or wrong organization | Check the user's role and organization. |
| Assets found but `0 rows remain` | Severity filter too high, device identifier missing, **or field names don't match** | Lower `$MinimumCVSSSeverity` to `0.0` for a test run; check `$DeviceIdField`; see Section 12 item 1. |
| Rows missing for some devices | Those assets have no `host_name` (or `ip`) | Try the other `$DeviceIdField` value. |

### NinjaOne

| Symptom | Likely cause | What to do |
|---|---|---|
| `COULD NOT LOG IN TO NINJAONE` | Wrong host, wrong Client ID/Secret, wrong app type | Confirm `$NinjaOneBaseURL` matches where you log in. Recreate the app as **API Services (machine-to-machine)** with **Client Credentials** and copy the Secret again. |
| `THE UPLOAD TO NINJAONE FAILED` | Wrong Scan Group ID, missing scopes, **or the upload address is not valid** | Confirm the ID, confirm both scopes, and see Section 12 item 3. |
| Upload succeeds but nothing shows on devices | Device names in the CSV don't match NinjaOne device names | Compare `Device ID` values with the names in NinjaOne; try the other `$DeviceIdField`. |
| Data didn't change after re-run | Upload not replacing as expected | Confirm in NinjaOne; see Section 12 item 4. |

### PowerShell

| Symptom | Likely cause | What to do |
|---|---|---|
| `A parameter cannot be found that matches parameter name 'Form'` | Running Windows PowerShell 5.1 | Run with **PowerShell 7** (`pwsh`). |
| `running scripts is disabled on this system` | Execution policy | Use `Set-ExecutionPolicy -Scope Process -ExecutionPolicy Bypass`. |
| Log or CSV appears in an unexpected folder | Relative paths use the current folder | Run from the script's folder or use full paths. |

---

## 12. Known limitations and unverified items

Please review these points before relying on this script.

1. **Rapid7 response field names are unverified.** The script reads `host_name`, `ip`, `vulnerability_id`, `cves`, `cvss_v3_score` and `cvss_v2_score`. These could not be confirmed against Rapid7's official schema. If a field name is wrong, the corresponding value is empty (null), and rows will be dropped (for example, no `cves` means no rows). The CVE field is handled whether Rapid7 returns it as a list or as comma-separated text.
2. **Which findings count as "current" is an assumption.** The script uses `new` + `same` findings and skips `remediated`, with `includeSame=true`. How Rapid7 defines these groups by default (for example, what time window "new" is measured against) was not confirmed.
3. **The NinjaOne upload address is unverified.** The script posts to `/api/v2/vulnerability/scan-groups/{id}/upload`. This endpoint could not be found in NinjaOne's public documentation. NinjaOne's documented method is a CSV import in the console (Scan Groups > Import file). If the upload fails, confirm the correct address with NinjaOne support or use the console import with the generated CSV.
4. **"Each run replaces the last" is not verified.** NinjaOne's documentation says it resolves vulnerabilities when you upload new scans that no longer contain them, which is consistent with this, but behavior through the API was not verified.
5. **CSV columns are unverified.** `Vendor`, `Device ID` and `CVE ID` were not verified against NinjaOne's importer. NinjaOne's console lets you map CSV columns, so the headings may need mapping when the Scan Group is created.
6. **Findings without a CVE are skipped.** Vendor-specific advisories with no CVE will not appear in NinjaOne.
7. **Cloud (Insight Platform) only.** On-premises Security Console is not supported.
8. **Large environments.** The script holds all assets and vulnerability definitions in memory. **Performance limits and Rapid7 rate limits: null (not tested).**
9. **No retry logic.** A temporary failure stops the run; it will be picked up on the next scheduled run.
10. **Scan Group ID location in the console: null (not verified).**

### Recommended validation before production use
1. Run once manually with `$MinimumCVSSSeverity = 0.0` and confirm the CSV has populated `Device ID` and `CVE ID` columns.
2. Spot-check 3 to 5 devices against Rapid7 and NinjaOne.
3. Confirm the NinjaOne upload works, or fall back to the console import.
4. Only then set the real severity threshold and schedule it.

---

## 13. Security guidance

- **The script file contains secrets** (Rapid7 API key and NinjaOne Client Secret) once filled in. **Do not commit the filled-in file to GitHub** or share it. Keep the repository copy with placeholder values only.
- Restrict file permissions on the script to the account that runs it.
- Use a **dedicated** Rapid7 user and API key for this integration, with the minimum role needed (asset read access).
- Give the NinjaOne API app only the scopes this script needs (Monitoring and Management).
- **Rotate** keys and secrets if they may have been exposed. Rapid7 keys can be regenerated from the API Keys page; NinjaOne client secrets can be regenerated from the app's settings.
- The log file contains device counts and error messages but not credentials. The CSV contains device names and CVEs; store it appropriately.
- Consider moving secrets out of the file (for example, into a secrets manager or environment variables) if your organization requires it. The script as written reads them from the settings block.

---

## 14. Quick reference

```text
Run:            pwsh .\Sync-InsightVM-To-NinjaOne.ps1
Needs:          PowerShell 7+, Rapid7 API key + region, NinjaOne Client ID/Secret, Scan Group ID
Rapid7 auth:    X-Api-Key header (no username/password)
Rapid7 URL:     https://{region}.api.insight.rapid7.com/vm/v4/integration/{assets|vulnerabilities}
NinjaOne auth:  OAuth 2.0 client credentials (scopes: monitoring management)
Output:         insightvm-export.csv  +  sync-log.txt
First check:    Open the CSV. Are Device ID and CVE ID filled in?
```

---

## 15. Supporting documents

Rapid7:
- InsightVM Cloud Integrations API (v4) overview: authorization via `X-Api-Key`, regional base path, versioning: <https://qveris.ai/providers/insightvm_cloud_integrations>
- Official Rapid7 reference (cited by third parties): <https://help.rapid7.com/insightvm/en-us/api/integrations.html>
- API key creation steps and required roles (third-party guide): <https://support.preludesecurity.com/rapid7-vulnerability-management>
- Test requests and note that organization keys connected more reliably: <https://learn.microsoft.com/en-us/security-exposure-management/rapid7-data-connector>
- Cloud Integrations API client behavior (header, regions, cursor-paginated search): <https://github.com/PostHog/posthog/pull/71148>

NinjaOne:
- Public API reference: <https://app.ninjarmm.com/apidocs-beta/core-resources>
- How to set up an API OAuth token: <https://www.ninjaone.com/docs/integrations/how-to-set-up-api-oauth-token/>
- Rapid7 Vulnerability Importer setup (CSV import and Scan Groups): <https://www.ninjaone.com/de/docs/integrationen-und-anwendungen-von-drittanbietern/schwachstellen-management/rapid7-vulnerability-importer-setup-and-configuration/>
- Access tokens last one hour (client-credentials flow): <https://docs.airbyte.com/integrations/sources/ninjaone-rmm-migrations>

---

*Document covers `Sync-InsightVM-To-NinjaOne.ps1`. Last updated: November 10, 2026.*
