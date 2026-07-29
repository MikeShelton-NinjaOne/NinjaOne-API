# NinjaOne Tenant-Wide Software Inventory Export

This guide walks through setting up and running `Get-NinjaSoftwareInventory-OC-TenantWide.ps1`, which exports a CSV of every installed application on every device across your entire NinjaOne Oceania (OC) tenant.

---

## What this script does

- Authenticates to the NinjaOne Public API v2 using a Client Credentials app (no browser login required — fully silent/unattended).
- Fetches a list of every device in the tenant (hostname, OS type, organization, location).
- Fetches the full software inventory for the tenant in one bulk, paginated call.
- Joins the two together and writes a single CSV with one row per installed application per device.
- Optionally, can be scoped to a single organization instead of the whole tenant.

It is hard-coded to the **Oceania (OC)** instance (`https://oc.ninjarmm.com`). It will not work against another region (US, US2, EU, CA) without changing the `BaseUrl` and `TokenEndpoint` values.

---

## Requirements

| Requirement | Notes |
|---|---|
| Windows PowerShell 5.1+ | Comes with Windows. Run `$PSVersionTable.PSVersion` to check. |
| Network access to `oc.ninjarmm.com` | Run from a machine that can reach the internet/your NinjaOne instance. |
| A NinjaOne API Client App | See setup steps below. |
| Permission to write to `C:\Windows\Temp` | This is where the output CSV is saved. Change the path in the script if needed. |

No external PowerShell modules are required — the script only uses built-in cmdlets (`Invoke-RestMethod`, `Export-Csv`, etc.).

---

## Step 1: Create the API Client App in NinjaOne

1. Log in to NinjaOne.
2. Go to **Administration > Apps > API > Client App IDs**.
3. Click **Add**.
4. Configure it with these exact settings:

   | Field | Value |
   |---|---|
   | Application Platform | **API Services (Machine-to-Machine)** |
   | Allowed Scopes | **Monitoring** |
   | Redirect URI | Leave blank (not used for this grant type) |

5. Click **Save**.
6. Copy the **Client ID** and **Client Secret** somewhere safe.
   - The Client Secret is only shown once. If you lose it, you'll need to regenerate it.

> This script only *reads* data (device list + software inventory), so the **Monitoring** scope alone is sufficient — you do not need to enable Management.

---

## Step 2: Edit the configuration block

Open `Get-NinjaSoftwareInventory-OC-TenantWide.ps1` in Notepad, VS Code, or the PowerShell ISE. Near the top, under `# === CONFIGURATION ===`, fill in:

```powershell
$ClientId      = '<Your Client ID>'
$ClientSecret  = '<Your Client Secret>'
```

Replace the placeholder text (including the angle brackets) with the values you copied in Step 1. Leave the quotation marks in place.

### Optional settings

| Setting | Default | What it does |
|---|---|---|
| `$OrganizationFilterId` | `0` | Set to a specific organization's ID to export software for just that org instead of the whole tenant. Leave at `0` to export everything. |
| `$PageSize` | `1000` | How many records to request per API page. 1000 is a safe default; you generally don't need to change this. |

**Finding an Organization ID:** open the organization in NinjaOne and look at the browser URL — the number after `/organization/` is the ID:

```
https://oc.ninjarmm.com/#/organization/7/overview
                                       ^
                                 Organization ID
```

---

## Step 3: Run the script

Open a PowerShell window, navigate to the folder containing the script, and run:

```powershell
.\Get-NinjaSoftwareInventory-OC-TenantWide.ps1
```

If PowerShell blocks the script from running due to execution policy, you may need to run this first (in an elevated/admin PowerShell window):

```powershell
Set-ExecutionPolicy -Scope Process -ExecutionPolicy Bypass
```

This only affects the current PowerShell window/session, not your whole system.

---

## What you'll see while it runs

The script prints progress through four stages:

1. **Authenticating** — obtains an access token using your Client ID/Secret.
2. **Fetching device list** — pages through every device in scope, printing a running count.
3. **Fetching software inventory** — pages through the tenant-wide software query, printing a running count.
4. **Building output and exporting to CSV** — joins the data, shows a preview of the first 25 rows in the console, then saves the CSV.

A typical run looks like:

```
  [1/4] Authenticating (Client Credentials)...
  [✓] Authenticated.

  [2/4] Fetching device list (this may take a few pages on large tenants)...
      ...1000 device(s) so far
      ...1840 device(s) so far
  [✓] 1840 device(s) indexed.

  [3/4] Fetching software inventory for the whole tenant...
      ...1000 software record(s) so far
      ...2000 software record(s) so far
      ...2650 software record(s) so far
  [✓] 2650 total software record(s) retrieved.

  [4/4] Building output and exporting to CSV...
```

---

## Output

The script writes a CSV to:

```
C:\Windows\Temp\NinjaSoftware_AllOrgs_<timestamp>.csv
```

(or `NinjaSoftware_Org<ID>_<timestamp>.csv` if you set `$OrganizationFilterId`).

### Columns

| Column | Description |
|---|---|
| `Publisher` | Software publisher, as reported by the OS. |
| `SoftwareName` | Application display name. |
| `Version` | Installed version string. |
| `OSType` | Device class/OS type (e.g. `WINDOWS_WORKSTATION`, `WINDOWS_SERVER`, `MAC`). |
| `InstallDate` | OS-reported install date, formatted `yyyy-MM-dd`. |
| `Hostname` | Device hostname. |
| `OrganizationName` | The NinjaOne organization the device belongs to. |
| `LocationName` | The location within that organization. |
| `DeviceId` | NinjaOne's internal numeric device ID. |

### A note on blank fields

**Blank does not mean "none" or "not applicable" — it means NinjaOne did not report a value for that field.** Not every device or OS reports every property (this varies especially between Windows, Mac, and Linux, and between workstations, servers, and network devices). The script never guesses or fills in a value that wasn't actually returned by the API — a blank cell is a blank cell.

If a software record references a device ID that isn't in the device list (rare — can happen if a device was deleted between the two API calls), the console will show a note like:

```
[i] 3 software record(s) referenced a device ID not found in the device list (left blank)
```

Those rows will still appear in the CSV with the software details filled in, but `Hostname`, `OSType`, `OrganizationName`, and `LocationName` left blank for that row.

---

## Install date limitation

The NinjaOne API only stores a single `installDate` per software record — it does not distinguish between first-install and last-install/reinstall dates. `InstallDate` in the CSV reflects whatever the operating system reported to NinjaOne at time of collection.

---

## Troubleshooting

| Symptom | Likely cause / fix |
|---|---|
| `Authentication failed` | Double-check `$ClientId` and `$ClientSecret`. Confirm the API app's platform is **API Services (Machine-to-Machine)** and the **Monitoring** scope is enabled. |
| `Please fill in ClientId and ClientSecret...` | You haven't replaced the `<...>` placeholders in the configuration block yet. |
| Script hangs or is very slow | Normal on large tenants — it's paging through devices and software in batches of `$PageSize`. Watch the console for the running counts to confirm it's progressing. |
| `No devices found — nothing to export` | Check `$OrganizationFilterId` — if set, confirm that organization ID actually exists and has devices. Set it back to `0` to test against the whole tenant. |
| `Failed to fetch device list` / `Failed to fetch software inventory` | Usually a network issue reaching `oc.ninjarmm.com`, or the API app is missing the Monitoring scope. The error message printed will include the underlying HTTP error. |
| CSV won't save | Confirm the account running the script has write access to `C:\Windows\Temp`, or edit the `$CsvPath` line near the bottom of the script to point somewhere else. |
| PowerShell won't run the `.ps1` file at all | Run `Set-ExecutionPolicy -Scope Process -ExecutionPolicy Bypass` in the same PowerShell window first (see Step 3). |

---

## Regional note

This script's `$BaseUrl` and `$TokenEndpoint` are hard-coded to `https://oc.ninjarmm.com` (Oceania). If your tenant is hosted in a different region, do not just change these two values and assume it will work — the rest of the script (query syntax, scopes) should still apply, but you should confirm against your instance before relying on the output. Regions include:

- `https://app.ninjarmm.com` (US primary)
- `https://us2.ninjarmm.com` (US secondary)
- `https://ca.ninjarmm.com` (Canada)
- `https://eu.ninjarmm.com` (Europe)
- `https://oc.ninjarmm.com` (Oceania / Australia)
