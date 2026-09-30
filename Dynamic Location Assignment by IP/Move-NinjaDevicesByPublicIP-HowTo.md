# How-To: Move NinjaRMM Devices to the Correct Location by Public IP

This guide explains how to configure and run `Move-NinjaDevicesByPublicIP.ps1`, a PowerShell script that automatically moves NinjaRMM devices into the correct Location based on their current Public IP address.

---

## 1. Overview

You maintain a CSV file that says *"this Public IP belongs to this Location."* The script:

1. Logs in to the NinjaRMM API using your credentials.
2. Pulls a list of all devices (or just devices in one Organization, if you choose).
3. Reads your CSV mapping of Public IPs to Locations.
4. For every device, checks its current Public IP against that mapping:
   - **IP matches, already in the right Location** → does nothing.
   - **IP matches, but device is in the wrong Location** → moves it.
   - **IP matches, but the move fails** (e.g. API error) → logs it as a failure.
   - **IP has no match in the CSV at all** → logs it as unmatched.
5. Prints a summary to the screen and saves three result CSV files.

---

## 2. Files You Need

| File | Purpose |
|------|---------|
| `Move-NinjaDevicesByPublicIP.ps1` | The script itself. |
| `ninjapublicips.csv` | The input file you maintain — maps Public IPs to Locations. |

Both of these should exist before you run the script. The CSV must be in the same folder you configure as `$WorkingFolder` (see Section 3).

---

## 3. Step 1 — Fill In the Configuration Section

Open `Move-NinjaDevicesByPublicIP.ps1` in a text editor. Near the top, under the header comment, you'll find a clearly marked section:

```powershell
# USER CONFIGURATION — EDIT THE VALUES IN THIS SECTION BEFORE RUNNING THE SCRIPT
```

Fill in each value as described below.

### 3.1 `$NinjaClientId` and `$NinjaClientSecret` — **Required**

These are your NinjaRMM API credentials.

**Where to get them:**
1. Log in to your NinjaRMM dashboard.
2. Go to **Administration → Apps → API**.
3. Create a new API Client (or use an existing one) with these scopes enabled:
   - `monitoring`
   - `management`
4. Copy the **Client ID** and **Client Secret** it generates.

**In the script:**
```powershell
$NinjaClientId     = 'your-client-id-here'
$NinjaClientSecret = 'your-client-secret-here'
```

> ⚠️ **Security note:** These values are stored in plain text inside the script file. Keep the file in a secure location and restrict who can open it. If this script will be shared or reused by others, consider adapting it to pull credentials from a secure vault instead of hardcoding them.

### 3.2 `$NinjaRegionUrl` — **Required**

NinjaRMM has different base URLs depending on which region your account is hosted in. If this is wrong, the script cannot authenticate.

Common values:

| Region | URL |
|--------|-----|
| US | `https://app.ninjarmm.com` |
| EU | `https://eu.ninjarmm.com` |
| Oceania | `https://oc.ninjarmm.com` |
| Canada | `https://ca.ninjarmm.com` |

**In the script:**
```powershell
$NinjaRegionUrl = 'https://app.ninjarmm.com'
```

If you're not sure which region applies to you, check the URL you use to log into the NinjaRMM web dashboard — it should match one of the above.

### 3.3 `$WorkingFolder` — **Required**

This is the folder on the machine running the script where:
- The input CSV (`ninjapublicips.csv`) must already exist.
- The three output report CSVs will be saved once the script finishes.

**In the script:**
```powershell
$WorkingFolder = 'C:\Temp'
```

Notes:
- This can be any folder you have read/write access to — it doesn't need to be `C:\Temp`.
- Do **not** include a trailing slash (use `C:\Temp`, not `C:\Temp\`).
- The folder must already exist — the script does not create it for you.

### 3.4 `$InputCsvFileName` — Usually leave as-is

This is the exact filename the script looks for inside `$WorkingFolder`.

**In the script:**
```powershell
$InputCsvFileName = 'ninjapublicips.csv'
```

Only change this if you've renamed your CSV file to something else. Otherwise, leave it as the default.

### 3.5 `$OrganizationIdFilter` — **Optional**

Controls whether the script processes every device across your entire NinjaRMM tenant, or just one Organization.

**In the script:**
```powershell
$OrganizationIdFilter = $null
```

- Leave as `$null` to check devices across **all** Organizations.
- Set it to a specific Organization ID (a number) to limit the script to just that Organization, e.g.:

  ```powershell
  $OrganizationIdFilter = 12
  ```

**Finding an Organization ID:** In the NinjaRMM dashboard, open the Organization and check the URL — the ID is usually the number in the address bar, or it can be pulled via the NinjaRMM API `/v2/organizations` endpoint.

---

## 4. Step 2 — Prepare the Input CSV

The CSV is how you tell the script which Public IP belongs to which Location. It must be named `ninjapublicips.csv` (or whatever you set `$InputCsvFileName` to) and saved inside `$WorkingFolder`.

### 4.1 Required Columns

The header row must contain exactly these four column names:

| Column | Type | Description |
|--------|------|-------------|
| `PublicIP` | Text | The public IP address associated with a location. |
| `OrgID` | Number | The NinjaRMM Organization ID that this location belongs to. |
| `Location` | Text | A friendly display name for the location (used only in reports — not sent to the API). |
| `LocationID` | Number | The NinjaRMM Location ID that matching devices should be moved into. |

### 4.2 Example File

```csv
PublicIP,OrgID,Location,LocationID
203.0.113.10,12,Main Office,45
203.0.113.55,12,Branch Office,46
198.51.100.2,18,Warehouse,52
```

A ready-to-edit version of this file, `ninjapublicips.csv`, is provided alongside this guide — just open it and replace the sample rows with your real data.

### 4.3 Validation Rules

- **No duplicate `PublicIP` values.** Every row must have a unique Public IP. If two rows share the same IP, the script will stop immediately and list which IPs are duplicated — fix the CSV and re-run.
- **Blank `PublicIP` rows are ignored automatically** — you don't need to delete empty rows, though keeping the file tidy is good practice.
- **Column names are case-sensitive in spirit** — keep them exactly as shown (`PublicIP`, `OrgID`, `Location`, `LocationID`) to avoid matching issues.
- Save the file as a standard `.csv` (comma-separated). UTF-8 encoding is safest if you're editing in Excel or a text editor.

### 4.4 Where to Find Location IDs

To fill in `LocationID` correctly, you'll need the numeric ID NinjaRMM uses internally for each location — this is different from the location's display name. You can find these via:
- The NinjaRMM API `/v2/organization/{id}/locations` endpoint, or
- Looking at the URL when viewing a specific Location in the NinjaRMM dashboard.

---

## 5. Step 3 — Run the Script

1. Open PowerShell (Windows PowerShell 5.1+ or PowerShell 7+).
2. Navigate to the folder containing the script, or reference it by full path.
3. Run it:

   ```powershell
   .\Move-NinjaDevicesByPublicIP.ps1
   ```

   or, using the full path:

   ```powershell
   C:\Path\To\Move-NinjaDevicesByPublicIP.ps1
   ```

No parameters need to be passed at the command line — everything is controlled by the configuration values you set inside the script in Step 1.

### 5.1 What Happens If Something's Missing

- If `$NinjaClientId` or `$NinjaClientSecret` are left blank, the script stops immediately with a clear error message telling you to fill them in.
- If the input CSV can't be found at `$WorkingFolder\$InputCsvFileName`, the script stops and tells you exactly which path it looked for.
- If the CSV has duplicate Public IPs, the script stops and lists each duplicated IP.

### 5.2 What You'll See While It Runs

- `INFO:` messages as it imports the CSV, requests an API token, and fetches the device list.
- A progress bar labeled **"Comparing Public IP to Location..."** while it works through matched devices.
- A color-coded summary at the end:

  | Color | Meaning |
  |-------|---------|
  | 🟢 Dark Green | Devices successfully moved. |
  | 🟡 Dark Yellow | Devices with no matching Public IP found in the CSV. |
  | 🔴 Dark Red | Devices that failed to move, or setup/validation errors. |
  | 🔵 Dark Cyan | No devices needed to move — everything already matched. |

---

## 6. Step 4 — Review the Output

All output files are saved inside `$WorkingFolder`, alongside your input CSV.

| File | Created When | Contents |
|------|---------------|----------|
| `movedendpoints.csv` | At least one device was moved | Device name, old Org/Location, new Location, Public IP. |
| `unmatchedendpoints.csv` | At least one device had no IP match | Device name, Org, old Location, Public IP. |
| `failedendpoints.csv` | At least one move attempt failed | Device name, Org, old Location, the Location it tried to move to, Public IP, and the specific error message. |

If none of the three categories have any entries, no files are created, and you'll see an informational message instead.

---

## 7. Troubleshooting

| Symptom | Likely Cause | Fix |
|---------|---------------|-----|
| `$NinjaClientId and $NinjaClientSecret must be filled in...` | Credentials left blank in the config section. | Fill in both values under **USER CONFIGURATION**. |
| `FAIL: Input CSV not found at ...` | CSV missing, misnamed, or in the wrong folder. | Confirm `$WorkingFolder` and `$InputCsvFileName` match where the file actually is. |
| `FAIL: Duplicate Public IPs found...` | Two or more CSV rows share the same `PublicIP`. | Edit the CSV to remove or correct the duplicate(s), then re-run. |
| Script fails right after "Requesting NinjaRMM API access token" | Invalid Client ID/Secret, missing scopes, or wrong `$NinjaRegionUrl`. | Double-check credentials and scopes in NinjaRMM; confirm the region URL matches your account. |
| A device appears in `unmatchedendpoints.csv` unexpectedly | Its current Public IP isn't in the CSV, or there's a typo in the IP. | Check the device's actual Public IP in NinjaRMM and compare against your CSV row. |
| A device appears in `failedendpoints.csv` | The move was attempted but the API rejected it. | Check the `Error` column in that file for the specific reason (e.g. invalid `LocationID`, insufficient permissions). |
| Script seems to log in more than once | Expected behavior. | The script requests a new token before each device move to avoid failures from an expired token during long runs. |

---

## 8. Quick Reference Checklist

Before running the script, confirm:

- [ ] `$NinjaClientId` and `$NinjaClientSecret` are filled in.
- [ ] `$NinjaRegionUrl` matches your NinjaRMM account's region.
- [ ] `$WorkingFolder` points to a real, accessible folder.
- [ ] `ninjapublicips.csv` exists inside that folder, with the correct headers and no duplicate Public IPs.
- [ ] `$OrganizationIdFilter` is set correctly (or left as `$null` to cover all Organizations).

---

**Script Author:** TawTek
**Script Version:** 2.0
**Guide Last Updated:** 2026-09-30
