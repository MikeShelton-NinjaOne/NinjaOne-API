# Get-NinjaSoftwareInventory-OC-SingleDevice.ps1

> Pulls the complete software inventory for a single device from the NinjaOne Oceania (OC) instance and exports a CSV containing publisher, software name, version, OS type, install date, hostname, and location name.

---

## Table of Contents

- [What This Script Does](#what-this-script-does)
- [Output Columns](#output-columns)
- [Prerequisites](#prerequisites)
- [Part 1: Create the API App](#part-1-create-the-api-app)
- [Part 2: Configure the Script](#part-2-configure-the-script)
- [Part 3: Provide the Device ID](#part-3-provide-the-device-id)
  - [Method 1 — NinjaOne Script Variable (recommended for automations)](#method-1--ninjaone-script-variable-recommended-for-automations)
  - [Method 2 — Command-line parameter](#method-2--command-line-parameter)
  - [Method 3 — Config block](#method-3--config-block)
- [Finding the Device ID](#finding-the-device-id)
- [Running the Script](#running-the-script)
  - [Manually from PowerShell](#manually-from-powershell)
  - [As a NinjaOne Automation](#as-a-ninjaone-automation)
- [Where the CSV Is Saved](#where-the-csv-is-saved)
- [Reading the Output](#reading-the-output)
- [A Note on Install Dates](#a-note-on-install-dates)
- [API Endpoints Used](#api-endpoints-used)
- [Troubleshooting](#troubleshooting)
- [Pre-Flight Checklist](#pre-flight-checklist)

---

## What This Script Does

1. **Authenticates silently** to the NinjaOne OC instance using Client Credentials — no browser, no login prompt
2. **Fetches device details** — hostname, OS type, location ID, and organization ID
3. **Resolves the location name** by looking up the org's location list
4. **Fetches the full software list** for that specific device
5. **Exports a CSV** with one row per installed application, enriched with device context

---

## Output Columns

| Column | Description | Example |
|---|---|---|
| `Publisher` | Company that published the software | `Microsoft Corporation` |
| `SoftwareName` | Name of the installed application | `Microsoft Visual C++ 2019` |
| `Version` | Installed version number | `14.28.29914.0` |
| `OSType` | Device class reported by NinjaOne | `WINDOWS_WORKSTATION` |
| `InstallDate` | Date the software was installed (OS-reported) | `2024-11-15` |
| `Hostname` | System name of the device | `DESKTOP-ABC123` |
| `LocationName` | Name of the NinjaOne location the device belongs to | `Sydney Office` |

**OSType values you may see:**

| Value | Meaning |
|---|---|
| `WINDOWS_WORKSTATION` | Windows desktop or laptop |
| `WINDOWS_SERVER` | Windows Server |
| `MAC` | macOS device |
| `LINUX_WORKSTATION` | Linux desktop |
| `LINUX_SERVER` | Linux server |
| `NMS_ROUTER` / `NMS_SWITCH` | Network device — software list may be empty |

---

## Prerequisites

| Requirement | Notes |
|---|---|
| PowerShell 5.1 or later | Built-in on Windows 10/11. Run `$PSVersionTable` to check your version. |
| NinjaOne System Administrator access | Required to create the API app |
| No extra modules | Uses only built-in PowerShell — nothing to install |
| NinjaOne OC instance | This script is hard-coded to `oc.ninjarmm.com` |

---

## Part 1: Create the API App

This is a **one-time setup**. You are creating a silent set of credentials the script uses to authenticate.

1. Log into your NinjaOne OC portal as a **System Administrator**
   (`https://oc.ninjarmm.com`)
2. Go to: **Administration → Apps → API → Client App IDs → Add**
3. Fill in the form:

   | Field | What to Enter |
   |---|---|
   | **Name** | Any name, e.g. `SoftwareInventoryScript` |
   | **Platform** | `API Services (Machine-to-Machine)` |
   | **Allowed Scopes** | ✅ `monitoring` — this is the only scope needed |
   | **Redirect URI** | Leave blank |

4. Click **Save**
5. You will see a **Client ID** and **Client Secret** — **copy both immediately**

> ⚠️ The Client Secret is shown only once. If you miss it, delete the app and create a new one.

> ℹ️ Only the `monitoring` scope is required. This script is read-only — it does not create, update, or delete anything in NinjaOne.

---

## Part 2: Configure the Script

Open `Get-NinjaSoftwareInventory-OC-SingleDevice.ps1` in any text editor (Notepad, VS Code, PowerShell ISE). Find the **CONFIGURATION** block near the top:

```powershell
# ==============================================================================
#  CONFIGURATION — Fill in your Client ID and Secret below
# ==============================================================================

# Hard-coded to Oceania instance — do not change
$BaseUrl       = 'https://oc.ninjarmm.com'
$TokenEndpoint = 'https://oc.ninjarmm.com/ws/oauth/token'

# From Administration > Apps > API > Client App IDs
$ClientId      = '<Your Client ID>'

# From Administration > Apps > API > Client App IDs (shown once at creation)
$ClientSecret  = '<Your Client Secret>'

# Set this when running manually outside of NinjaOne.
# Leave as 0 if using a NinjaOne Script Variable or the -DeviceId parameter.
$ManualDeviceId = 0   # e.g. 12345
```

Replace `<Your Client ID>` and `<Your Client Secret>` with the values you copied in Part 1. Leave everything else as-is — the URLs are already set correctly for the OC instance.

> ⚠️ Do not commit the `$ClientSecret` value to a public repository. Treat it like a password.

---

## Part 3: Provide the Device ID

The script needs to know which device to pull software for. There are three ways to provide the Device ID — the script checks them in this order and uses whichever one is set:

### Method 1 — NinjaOne Script Variable (recommended for automations)

When running the script as a NinjaOne automation, create a **Script Variable** named `targetDeviceId`. NinjaOne injects it as an environment variable automatically and the script reads it without any extra configuration.

**How to create the Script Variable:**

1. Go to: **Administration → Scripting → Script Variables**
2. Click **Add Script Variable**
3. Fill in:

   | Field | Value |
   |---|---|
   | **Name** | `targetDeviceId` ← must be this exact spelling |
   | **Label** | `Target Device ID` |
   | **Type** | `Integer` |
   | **Required** | Yes |
   | **Description** | `The numeric ID of the device to pull software inventory for. Find it in the device URL.` |

4. Click **Save**

When a technician runs the script from NinjaOne, they will be prompted to enter the Device ID before it executes.

---

### Method 2 — Command-line parameter

When running the script manually from PowerShell, pass the Device ID with the `-DeviceId` flag:

```powershell
.\Get-NinjaSoftwareInventory-OC-SingleDevice.ps1 -DeviceId 12345
```

---

### Method 3 — Config block

Open the script and set `$ManualDeviceId` directly in the CONFIGURATION block:

```powershell
$ManualDeviceId = 12345
```

Then run the script normally without any parameters:

```powershell
.\Get-NinjaSoftwareInventory-OC-SingleDevice.ps1
```

This method is useful for one-off runs where you always want the same device.

---

## Finding the Device ID

The Device ID is the number that appears in the browser URL when you open a device in NinjaOne.

**Step-by-step:**

1. Log into your NinjaOne OC portal: `https://oc.ninjarmm.com`
2. Navigate to **Devices** in the left sidebar
3. Click on the device you want to pull inventory for
4. Look at the URL in your browser address bar:

```
https://oc.ninjarmm.com/#/deviceDashboard/12345/overview
                                           ^^^^^
                                    This is the Device ID
```

5. Copy that number — it is the Device ID

**Alternative — via the API:**
If you need to find a device ID programmatically, call `GET /v2/devices` and find the device by `systemName`. The `id` field on each device object is the Device ID.

---

## Running the Script

### Manually from PowerShell

1. Open **PowerShell** (does not need to be run as Administrator)
2. Navigate to the folder where the script is saved:
   ```powershell
   cd C:\Scripts
   ```
3. Run using whichever method suits your workflow:

   ```powershell
   # Method 1 — with -DeviceId parameter (most common for manual runs)
   .\Get-NinjaSoftwareInventory-OC-SingleDevice.ps1 -DeviceId 12345

   # Method 2 — using ManualDeviceId set in the config block
   .\Get-NinjaSoftwareInventory-OC-SingleDevice.ps1
   ```

4. Watch the progress output in the console:

   ```
   ============================================================
   NinjaOne Software Inventory — Single Device
   Instance  : oc.ninjarmm.com (Oceania)
   Device ID : 12345
   ============================================================

   [1/4] Authenticating (Client Credentials)...
   [✓] Authenticated.

   [2/4] Fetching device info (ID: 12345)...
   [✓] Device : DESKTOP-ABC123
       OS Type: WINDOWS_WORKSTATION
       Org ID : 7
       Location: Sydney Office

   [3/4] Fetching software inventory for DESKTOP-ABC123...
   [✓] 142 software record(s) found.

   [4/4] Building output and exporting to CSV...

   Preview (first 25 records):
   Publisher                 SoftwareName                    Version        ...
   ---------                 ------------                    -------        ...
   Google LLC                Google Chrome                   133.0.6943.142 ...
   Microsoft Corporation     Microsoft Edge                  121.0.2277.128 ...
   ...

   ============================================================
   [✓] COMPLETE
       Device    : DESKTOP-ABC123 (ID: 12345)
       OS Type   : WINDOWS_WORKSTATION
       Location  : Sydney Office
       Records   : 142 software records
       CSV saved : C:\Windows\Temp\NinjaSoftware_DESKTOP-ABC123_20260701_143022.csv
   ============================================================
   ```

5. The CSV is saved to `C:\Windows\Temp\` with a timestamped filename

---

### As a NinjaOne Automation

Running the script as a NinjaOne automation means it executes on the endpoint itself and the CSV is saved locally on that device.

**Step 1 — Upload the script to NinjaOne**

1. Go to: **Administration → Scripting → Scripts**
2. Click **Add Script**
3. Fill in:

   | Field | Value |
   |---|---|
   | **Name** | `Get-NinjaSoftwareInventory-SingleDevice` |
   | **Language** | `PowerShell` |
   | **Run As** | `System` |
   | **Timeout** | `120` seconds |
   | **Script Body** | Paste the full contents of the `.ps1` file |

4. Click **Save**

**Step 2 — Create the Script Variable** (if not already done — see [Method 1](#method-1--ninjaone-script-variable-recommended-for-automations) above)

**Step 3 — Run the automation**

1. Navigate to the target device in NinjaOne
2. Right-click the device → **Run Script**, or click the **Scripts** button on the device page
3. Search for and select `Get-NinjaSoftwareInventory-SingleDevice`
4. Fill in the **Target Device ID** Script Variable when prompted
   > ℹ️ You can enter the ID of the device you are currently on, or the ID of a completely different device — the script fetches data for whichever ID you enter, regardless of which machine it physically runs on.
5. Click **Run**
6. Monitor the output in the NinjaOne **Script Activity** log

**Step 4 — Retrieve the CSV**

Since the script runs under the `System` account on the endpoint, the CSV is saved to:

```
C:\Windows\Temp\NinjaSoftware_HOSTNAME_TIMESTAMP.csv
```

To retrieve it, either:
- RDP into the device and navigate to `C:\Windows\Temp\`
- Use NinjaOne's **File Browser** feature to browse to that path
- Copy it to a network share by adding a line to the script after the `Export-Csv` call:
  ```powershell
  Copy-Item -Path $CsvPath -Destination '\\your-server\reports\'
  ```

---

## Where the CSV Is Saved

| How the script is run | CSV location |
|---|---|
| Manually from PowerShell (your workstation) | `C:\Windows\Temp\NinjaSoftware_HOSTNAME_TIMESTAMP.csv` on your machine |
| As a NinjaOne automation on an endpoint | `C:\Windows\Temp\NinjaSoftware_HOSTNAME_TIMESTAMP.csv` on **that endpoint** |

The filename is always in this format:
```
NinjaSoftware_DESKTOP-ABC123_20260701_143022.csv
```
Where `DESKTOP-ABC123` is the device hostname and `20260701_143022` is the date and time the script ran.

---

## Reading the Output

Open the CSV in Excel or any spreadsheet application. Each row represents one installed application on the target device.

**Example output:**

| Publisher | SoftwareName | Version | OSType | InstallDate | Hostname | LocationName |
|---|---|---|---|---|---|---|
| Google LLC | Google Chrome | 133.0.6943.142 | WINDOWS_WORKSTATION | 2025-01-15 | DESKTOP-ABC123 | Sydney Office |
| Microsoft Corporation | Microsoft Edge | 121.0.2277.128 | WINDOWS_WORKSTATION | 2024-11-20 | DESKTOP-ABC123 | Sydney Office |
| Microsoft Corporation | Microsoft Visual C++ 2019 | 14.28.29914.0 | WINDOWS_WORKSTATION | 2024-08-01 | DESKTOP-ABC123 | Sydney Office |

The records are sorted alphabetically by `SoftwareName` then `Hostname`.

**Empty fields:**
- `Publisher` may be blank for some applications that don't register a publisher
- `Version` may be blank for some older or manually installed software
- `InstallDate` may be blank if the OS did not record an install date

---

## A Note on Install Dates

The script exposes a single `InstallDate` column. This reflects the install date as recorded by the Windows registry or installer at the time the application was installed.

**The NinjaOne API does not provide separate first-install and last-install timestamps.** The `/v2/device/{id}/software` endpoint returns one record per installed application with one `installDate` field. There is no install history or event log exposed through the API.

If a piece of software has been reinstalled or updated, the `installDate` will reflect the most recent installation as reported by the OS — but this depends entirely on how the installer wrote to the registry, which varies by application.

---

## API Endpoints Used

| Method | Endpoint | Purpose |
|---|---|---|
| `POST` | `/ws/oauth/token` | Get access token (Client Credentials) |
| `GET` | `/v2/device/{id}` | Fetch device hostname, OS type, locationId, orgId |
| `GET` | `/v2/organization/{id}/locations` | Resolve location ID to location name |
| `GET` | `/v2/device/{id}/software` | Fetch full software list for this device |

All calls go to `https://oc.ninjarmm.com` — hard-coded for this customer's instance.

---

## Troubleshooting

| Error / Symptom | Solution |
|---|---|
| `Please fill in ClientId and ClientSecret` | The placeholder text is still in the config block. Replace `<Your Client ID>` and `<Your Client Secret>` with your real values. |
| `No device ID provided` | You haven't set the device ID through any of the three methods. See [Part 3](#part-3-provide-the-device-id). |
| `Authentication failed` | Check that `$ClientId` and `$ClientSecret` are correct. The API app must have Platform = `API Services (Machine-to-Machine)` and the `monitoring` scope enabled. |
| `Device ID not found (HTTP 404)` | The number you entered doesn't match any device in NinjaOne. Double-check the ID from the device URL in the OC portal. |
| `Failed to fetch software list` | The device type may not report a software inventory (common for NMS network devices and cloud monitors). This is a NinjaOne limitation — software lists are only available for agent-managed endpoints. |
| `LocationName shows 'Unknown'` | The script couldn't resolve the location. This is non-fatal — the CSV will still export with `Unknown` in that column. Check that the device has a location assigned in NinjaOne. |
| CSV opens with all data in one column | The file is comma-delimited. In Excel, use **Data → Text to Columns → Delimited → Comma** to split it, or open via **File → Import** and select comma as the delimiter. |
| CSV is empty except the header row | The device has no software records in NinjaOne. Confirm the NinjaOne agent is installed and reporting on the device. |
| Script runs instantly with no output in NinjaOne | Run As may not be set to `System`. Edit the script in NinjaOne and confirm `Run As = System`. |
| `targetDeviceId` Script Variable not appearing | The Script Variable name must be exactly `targetDeviceId` (camelCase, no spaces). Check spelling in Administration → Scripting → Script Variables. |

---

## Pre-Flight Checklist

- [ ] Logged into NinjaOne OC portal (`oc.ninjarmm.com`) as System Administrator
- [ ] API app created — Platform: `API Services (Machine-to-Machine)`, Scope: `monitoring`
- [ ] Client ID and Client Secret copied and stored securely
- [ ] `$ClientId` and `$ClientSecret` filled in the script CONFIGURATION block
- [ ] Device ID identified from the device URL in NinjaOne
- [ ] Device ID provided via Script Variable, `-DeviceId` parameter, or `$ManualDeviceId`
- [ ] Test run completed — confirmed `[✓] COMPLETE` in the output
- [ ] CSV file located at `C:\Windows\Temp\NinjaSoftware_HOSTNAME_TIMESTAMP.csv`
- [ ] CSV opened and columns verified
