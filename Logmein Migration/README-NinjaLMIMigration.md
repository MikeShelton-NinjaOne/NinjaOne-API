# Invoke-NinjaLMIMigration.ps1

> Reads a CSV exported from LogMeIn and migrates each device in NinjaOne — moving it to the correct organization and location, setting its display name, and writing notes to a custom field. Creates orgs and locations automatically if they don't already exist in NinjaOne.

---

## Table of Contents

- [What This Script Does](#what-this-script-does)
- [How It Works Row by Row](#how-it-works-row-by-row)
- [CSV Format](#csv-format)
- [Prerequisites](#prerequisites)
- [Part 1: Create the logmeinNotes Custom Field](#part-1-create-the-logmeinnotes-custom-field)
- [Part 2: Create the API App](#part-2-create-the-api-app)
- [Part 3: Configure the Script](#part-3-configure-the-script)
- [Part 4: Run the Script](#part-4-run-the-script)
- [Reading the Output](#reading-the-output)
- [How Empty Row Detection Works](#how-empty-row-detection-works)
- [How Device Matching Works](#how-device-matching-works)
- [API Endpoints Used](#api-endpoints-used)
- [Troubleshooting](#troubleshooting)
- [Pre-Flight Checklist](#pre-flight-checklist)

---

## What This Script Does

For each row in the CSV the script:

1. **Finds the device** in NinjaOne by matching the `Host Name` column against `systemName` and `dnsName`
2. **Finds or creates the organization** from the `Organization` column — if the org doesn't exist in NinjaOne it is created automatically
3. **Finds or creates the location** from the `Location` column inside that org — created automatically if missing
4. **Moves the device** to that org and location via `PATCH /v2/device/{id}`
5. **Sets the device display name** to the value in the `Computer description` column
6. **Writes the `Notes` column** to a device custom field named `logmeinNotes`
7. **Stops processing** after hitting 3 consecutive empty rows
8. **Exports a results CSV** alongside the input file showing the outcome of every row

---

## How It Works Row by Row

```
Row 2: CBT-GK17L84
  ✓ Device found in NinjaOne (ID: 12345)
  ✓ Org "3P Abstract" already exists (ID: 7)
  ✓ Location "Main Office" already exists (ID: 3)
  ✓ Device moved to org "3P Abstract", location "Main Office"
  ✓ Display name set to: CBT-GK17L84 (Michael Hirschler 5/6/25)
  ✓ Notes written to custom field 'logmeinNotes'
```

If an org or location doesn't exist:
```
Row 6: 2WMQFT3.priorityhcs.com
  ✓ Device found (hostname matched after stripping domain suffix)
  i Org "Above and Beyond Therapy" not found — creating...
  ✓ Org created (ID: 42)
  i Location "Main Office" not found in org — creating...
  ✓ Location created (ID: 88)
  ✓ Device moved to org "Above and Beyond Therapy", location "Main Office"
  ✓ Display name set to: 2WMQFT3(Rochel Cohen)
  i No notes to write for this device.
```

---

## CSV Format

The script expects this exact column order (matching your LogMeIn export):

| Column | Used For |
|---|---|
| `Computer description` | Sets the device display name in NinjaOne |
| `Host Name` | Finds the device in NinjaOne |
| `Organization` | Target org — found or created |
| `Location` | Target location inside the org — found or created |
| `secure name` | Not used |
| `username` | Not used |
| `password` | Not used |
| `Notes` | Written to the `logmeinNotes` custom field |

> ⚠️ The first row is treated as the header and skipped. Do not remove it.

**Example rows:**
```csv
Computer description,Host Name,Organization,Location,secure name,username,password,Notes
CBT-GK17L84 (Michael Hirschler 5/6/25),CBT-GK17L84,3P Abstract,Main Office,,,,Test note
2WMQFT3(Rochel Cohen),2WMQFT3.priorityhcs.com,Above and Beyond Therapy,Main Office,,,,
```

---

## Prerequisites

| Requirement | Notes |
|---|---|
| PowerShell 5.1+ | Built-in on Windows 10/11. Run `$PSVersionTable` to check. |
| NinjaOne System Administrator | Required to create the API app, custom field, and manage devices |
| No extra modules | Uses only built-in PowerShell — nothing to install |

---

## Part 1: Create the logmeinNotes Custom Field

The script writes notes to a custom field named `logmeinNotes`. This field **must be created manually in NinjaOne before running the script** — the script does not create it automatically.

1. Log into NinjaOne as a **System Administrator**
2. Go to: **Administration → Devices → Global Custom Fields**
3. Click **Add → Field**
4. Fill in the form exactly:

   | Setting | Value |
   |---|---|
   | **Label** | `LogMeIn Notes` |
   | **Name** | `logmeinNotes` ← must be exactly this |
   | **Type** | `Text` or `Multi-line Text` |
   | **Technician Permission** | `Read Only` (or `Read/Write` if technicians should edit it) |
   | **Script Permission** | `Read/Write` ← required |
   | **API Permission** | `Read/Write` ← required |

5. Click **Save**

> ⚠️ The field **Name** must be exactly `logmeinNotes` — this is what the script references. The Label can be anything you like.

> ⚠️ If Script Permission or API Permission is not set to `Read/Write`, the script will fail to write to the field with a permission error. It will still complete the device move but will log a warning for that row.

---

## Part 2: Create the API App

One-time setup — creates the silent credentials the script uses to authenticate.

1. Log into NinjaOne as a **System Administrator**
2. Go to: **Administration → Apps → API → Client App IDs → Add**
3. Fill in:

   | Field | Value |
   |---|---|
   | **Name** | Any name, e.g. `LMIMigrationScript` |
   | **Platform** | `API Services (Machine-to-Machine)` |
   | **Allowed Scopes** | ✅ `monitoring` AND ✅ `management` |
   | **Redirect URI** | Leave blank |

4. Click **Save** — copy the **Client ID** and **Client Secret** immediately

> ⚠️ The Client Secret is shown only once. If you miss it, delete the app and create a new one.

---

## Part 3: Configure the Script

Open `Invoke-NinjaLMIMigration.ps1` in any text editor. Find the **CONFIGURATION** block at the very top and fill in all five values:

```powershell
# ==============================================================================
#  CONFIGURATION — Fill in ALL values in this block before running
# ==============================================================================

$BaseUrl       = 'https://<your Login URL>'
$TokenEndpoint = 'https://<your Login URL>/ws/oauth/token'
$ClientId      = '<Your Client ID>'
$ClientSecret  = '<Your Client Secret>'
$CsvPath       = '<Path to your CSV file>'
```

| Variable | Example | Notes |
|---|---|---|
| `$BaseUrl` | `https://app.ninjarmm.com` | Your NinjaOne login URL — no trailing slash |
| `$TokenEndpoint` | `https://app.ninjarmm.com/ws/oauth/token` | Same as `$BaseUrl` + `/ws/oauth/token` |
| `$ClientId` | `abc123...` | From the API app in Part 2 |
| `$ClientSecret` | `s3cr3t...` | From the API app — shown once at creation |
| `$CsvPath` | `C:\Users\You\Downloads\LMI_X_Ninja.csv` | Full path to your CSV file |

**Regional URLs:**

| Region | Base URL |
|---|---|
| United States | `https://app.ninjarmm.com` |
| Europe | `https://eu.ninjarmm.com` |
| Oceania | `https://oc.ninjarmm.com` |
| Canada | `https://ca.ninjarmm.com` |

**CSV path tips:**
- Use the full path to avoid ambiguity: `C:\Users\You\Downloads\LMI_X_Ninja.csv`
- Or use `.\LMI_X_Ninja.csv` if the CSV is in the same folder as the script
- Wrap in single quotes — do not use double quotes for paths

---

## Part 4: Run the Script

1. Open **PowerShell** (does not need to be run as Administrator)
2. Navigate to the script folder:
   ```powershell
   cd C:\Scripts
   ```
3. Run the script:
   ```powershell
   .\Invoke-NinjaLMIMigration.ps1
   ```
4. Watch the output — each device is reported as it processes:
   ```
   ================================================================
   NinjaOne LogMeIn Migration Script
   CSV : C:\Users\You\Downloads\LMI_X_Ninja.csv
   ================================================================

   [1/4] Authenticating...
   [✓] Authenticated.

   [2/4] Loading existing organizations and devices from NinjaOne...
   [✓] Loaded 45 org(s) and 312 device(s).

   [3/4] Processing CSV rows...

   ── Row 2 : 1JDWK44 ──────────────────────────────────────────────
       [✓] Device found: 1JDWK44 (ID: 1234)
       [✓] Org found: 3P Abstract (ID: 7)
       [✓] Location found: Main Office (ID: 3)
       [✓] Device moved to org "3P Abstract", location "Main Office"
       [✓] Display name set to: 1JDWK44 (Maya Nesser Home 4/25/25)
       [i] No notes to write for this device.

   ── Row 3 : CBT-GK17L84 ──────────────────────────────────────────
       [✓] Device found: CBT-GK17L84 (ID: 1235)
       ...

   [4/4] Summary
   ================================================================
   [✓] COMPLETE
       Rows processed : 5
       Migrated OK    : 4
       Skipped        : 1
       Errors         : 0
   ================================================================
   ```

5. A results CSV is saved automatically next to your input file:
   ```
   NinjaMigration_Results_20260701_143022.csv
   ```

---

## Reading the Output

### Console output colour guide

| Colour | Meaning |
|---|---|
| 🟢 Green `[✓]` | Step completed successfully |
| 🟡 Yellow `[i]` or `[!]` | Warning — non-fatal, script continues |
| 🔴 Red `[!]` | Error — this row was skipped or failed |
| ⚪ Gray `[i]` | Informational — nothing to action |

### Results CSV columns

| Column | Description |
|---|---|
| `Row` | Row number in the source CSV |
| `Hostname` | The hostname from the CSV |
| `Org` | Target organization name |
| `Location` | Target location name |
| `Status` | `OK`, `NOT FOUND`, or `ERROR` |
| `Notes` | What happened — e.g. `Org created`, `Location created`, `Custom field write failed` |

---

## How Empty Row Detection Works

The script reads the CSV line by line and counts consecutive empty rows. Once it hits **3 consecutive empty rows** it stops processing entirely.

- A row counts as empty if it is blank or contains only commas
- The counter resets to 0 whenever a non-empty row is found
- You can change the threshold by editing `$EmptyRowLimit = 3` in the config block

This means you can have gaps in your CSV (e.g. a single blank row between groups) without the script stopping — it only stops on 3 in a row.

---

## How Device Matching Works

The script tries to find each device in NinjaOne using two strategies:

**1. Exact hostname match** — compares the `Host Name` from the CSV against `systemName` and `dnsName` on every device. Matching is case-insensitive.

**2. Short hostname fallback** — if no exact match is found, the script strips everything after the first `.` and tries again. This handles FQDNs like `2WMQFT3.priorityhcs.com` — it tries `2WMQFT3` as the hostname.

If neither match finds the device, the row is skipped and logged as `NOT FOUND` in the results CSV.

---

## API Endpoints Used

| Method | Endpoint | Purpose |
|---|---|---|
| `POST` | `/ws/oauth/token` | Get access token (Client Credentials) |
| `GET` | `/v2/organizations` | Load all existing orgs (paginated) |
| `GET` | `/v2/devices` | Load all existing devices (paginated) |
| `GET` | `/v2/organization/{id}/locations` | Load locations per org |
| `POST` | `/v2/organizations` | Create a new org if not found |
| `POST` | `/v2/organization/{id}/locations` | Create a new location if not found |
| `PATCH` | `/v2/device/{id}` | Move device, set display name |
| `PATCH` | `/v2/device/{id}/custom-fields` | Write Notes to `logmeinNotes` field |

---

## Troubleshooting

| Error / Symptom | Solution |
|---|---|
| `Fill in $BaseUrl` | The `<your Login URL>` placeholder is still in the config block. |
| `Fill in $CsvPath` | The `<Path to your CSV file>` placeholder is still in the config block. |
| `CSV file not found` | The path in `$CsvPath` doesn't exist. Check for typos and use a full absolute path. |
| `Authentication failed` | Check `$ClientId` and `$ClientSecret`. API app must be `API Services (Machine-to-Machine)` with `monitoring` AND `management` scopes. |
| Device shows `NOT FOUND` | The hostname in the CSV doesn't match `systemName` or `dnsName` in NinjaOne. Check exact spelling on the device record in NinjaOne. |
| `Custom field write failed` | The `logmeinNotes` custom field doesn't exist, or its API Permission is not `Read/Write`. See Part 1. |
| Org created but device not moved | The PATCH call failed after org creation. Check the `Errors` count and `ERROR` rows in the results CSV for the specific error message. |
| Script stops too early | The CSV has empty rows before the end of the data. Either remove the blank rows or increase `$EmptyRowLimit` in the config block. |
| Script processes rows past the data | The CSV has no empty rows at the end. This is fine — the script will just process every row and stop naturally when it runs out of lines. |
| Display name not updating | The `Computer description` column is blank for that row. An empty display name is silently skipped. |

---

## Pre-Flight Checklist

- [ ] NinjaOne System Administrator access confirmed
- [ ] `logmeinNotes` global custom field created — Name exact, API and Script permissions = `Read/Write`
- [ ] API app created — Platform: `API Services (Machine-to-Machine)`, Scopes: `monitoring` + `management`
- [ ] Client ID and Client Secret saved securely
- [ ] `$BaseUrl`, `$TokenEndpoint`, `$ClientId`, `$ClientSecret`, `$CsvPath` all filled in
- [ ] CSV file path is correct and file is accessible
- [ ] Test run on a single-row CSV first before running the full file
- [ ] Confirmed `[✓] COMPLETE` output and reviewed the results CSV
