# Invoke-NinjaOffboardCapture.ps1

> Captures full device information and an offboard reason, then writes a structured report to an **Apps & Services document** on the device's organization in NinjaOne. Each offboarded device gets its own named document — they stack up permanently under the org's Documentation tab.

---

## Table of Contents

- [What This Script Does](#what-this-script-does)
- [Where the Data Goes](#where-the-data-goes)
- [What Gets Captured](#what-gets-captured)
- [Prerequisites](#prerequisites)
- [Part 1: Create the API App](#part-1-create-the-api-app)
- [Part 2: Configure the Script](#part-2-configure-the-script)
- [Part 3: Add the Script to NinjaOne](#part-3-add-the-script-to-ninjaone)
- [Part 4: Create the Script Variables](#part-4-create-the-script-variables)
- [Part 5: Run the Script](#part-5-run-the-script)
- [The Document Template](#the-document-template)
- [Finding a Device ID](#finding-a-device-id)
- [Authentication — Why No Login Prompt?](#authentication--why-no-login-prompt)
- [API Endpoints Used](#api-endpoints-used)
- [Troubleshooting](#troubleshooting)
- [Pre-Flight Checklist](#pre-flight-checklist)

---

## What This Script Does

When a technician offboards a device, this script:

1. **Authenticates silently** using Client Credentials — no browser, no login prompt
2. **Pulls device data** from the NinjaOne API — hostname, IP, OS, hardware, last contact, installed software, recent activity
3. **Takes the offboard reason** entered as a NinjaOne Script Variable when the script is triggered
4. **Checks whether the "Device Offboard Report" document template exists** — creates it automatically via the API if it does not
5. **Creates a new Apps & Services document** on the organization named `Offboard — <DeviceHostname>` with all captured data written into structured fields

The result is a permanent, searchable record sitting directly on the organization in NinjaOne, visible to any technician who opens that org's Documentation tab.

---

## Where the Data Goes

The offboard report is written to an **Apps & Services document** on the organization — not a custom field, not a global field, not a device field.

To find it after the script runs:

```
NinjaOne → Organizations → [Org Name] → Documentation tab → Apps & Services
```

Each device gets its own document named `Offboard — HOSTNAME`. If the same device is run through the script more than once (e.g. a re-used asset), the existing document is updated rather than creating a duplicate.

> ℹ️ The "Device Offboard Report" document template is created automatically the very first time the script runs. You do not need to create it manually.

---

## What Gets Captured

Each document has six structured fields:

| Field | Contents |
|---|---|
| **Offboard Reason** | Exactly what the technician typed when triggering the script |
| **Capture Date/Time** | Timestamp of when the script ran |
| **Device Details** | Hostname, DNS name, IP addresses, OS, device class, last contact, agent version, organization |
| **Hardware Info** | Manufacturer, model, serial number, CPU, RAM *(if available for device type)* |
| **Last Activity** | Most recent activity entry — timestamp, type, message |
| **Installed Software** | First 20 installed applications with version numbers |

---

## Prerequisites

| Requirement | Notes |
|---|---|
| PowerShell 5.1+ | Built-in on Windows 10/11. Run `$PSVersionTable` to check. |
| NinjaOne System Administrator | Required to create API credentials, upload scripts, and create Script Variables |
| No extra modules | Uses only built-in PowerShell — nothing to install |

---

## Part 1: Create the API App

This is a **one-time setup**. You are creating a silent set of credentials that let the script talk to NinjaOne without a browser or login prompt.

1. Log into NinjaOne as a **System Administrator**
2. Go to: **Administration → Apps → API → Client App IDs → Add**
3. Fill in the form:

   | Field | What to Enter |
   |---|---|
   | **Name** | Any name, e.g. `OffboardCaptureScript` |
   | **Platform** | `API Services (Machine-to-Machine)` |
   | **Allowed Scopes** | Check ✅ `monitoring` AND ✅ `management` |
   | **Redirect URI** | Leave blank — not needed for this flow |

4. Click **Save**
5. You will see a **Client ID** and **Client Secret** — **copy both immediately**

> ⚠️ The Client Secret is only shown once. If you miss it, delete the app and create a new one.

> ⚠️ Do not commit the Client Secret to source control. Treat it like a password.

---

## Part 2: Configure the Script

Open `Invoke-NinjaOffboardCapture.ps1` in any text editor. Find the **CONFIGURATION** block at the very top and fill in all four values:

```powershell
# ==============================================================================
#  CONFIGURATION — Fill in ALL four values below before saving/running the script
# ==============================================================================

$BaseUrl       = 'https://<your Login URL>'
$TokenEndpoint = 'https://<your Login URL>/ws/oauth/token'
$ClientId      = '<Your Client ID>'
$ClientSecret  = '<Your Client Secret>'
```

| Variable | Example | Where to Find It |
|---|---|---|
| `$BaseUrl` | `https://app.ninjarmm.com` | Your NinjaOne login URL |
| `$TokenEndpoint` | `https://app.ninjarmm.com/ws/oauth/token` | Same as `$BaseUrl` + `/ws/oauth/token` |
| `$ClientId` | `abc123...` | Shown after creating the app in Part 1 |
| `$ClientSecret` | `s3cr3t...` | Shown once at app creation |

**Regional URLs:**

| Region | Base URL |
|---|---|
| United States | `https://app.ninjarmm.com` |
| Europe | `https://eu.ninjarmm.com` |
| Oceania | `https://oc.ninjarmm.com` |
| Canada | `https://ca.ninjarmm.com` |

---

## Part 3: Add the Script to NinjaOne

1. Go to: **Administration → Scripting → Scripts → Add Script**
2. Fill in:

   | Field | Value |
   |---|---|
   | **Name** | `Invoke-NinjaOffboardCapture` |
   | **Language** | `PowerShell` |
   | **Run As** | `System` |
   | **Timeout** | `120` seconds |
   | **Script Body** | Paste the entire contents of the `.ps1` file |

3. Click **Save**

---

## Part 4: Create the Script Variables

Script Variables are the inputs a technician fills in each time they run the script. You need exactly two.

Go to: **Administration → Scripting → Script Variables**

### Variable 1 — Target Device ID

| Field | Value |
|---|---|
| **Name** | `targetDeviceId` ← must be exact, including capitalisation |
| **Label** | `Target Device ID` |
| **Type** | `Integer` |
| **Required** | Yes |
| **Description** | `The numeric ID of the device being offboarded. Find it in the device URL in NinjaOne.` |

### Variable 2 — Offboard Reason

| Field | Value |
|---|---|
| **Name** | `offboardReason` ← must be exact, including capitalisation |
| **Label** | `Offboard Reason` |
| **Type** | `Text` |
| **Required** | Yes |
| **Description** | `Why is this device being offboarded? E.g. employee departure, hardware failure, end of life.` |

> ⚠️ The variable **Names** must match exactly — `targetDeviceId` and `offboardReason`. The script reads them as environment variables by those exact names. A typo will cause the script to stop with an error.

---

## Part 5: Run the Script

1. In NinjaOne, navigate to **any managed device**
   > The script targets `targetDeviceId`, not the device it physically runs on — so it can run on any managed device, including the offboarded device itself.
2. Right-click the device → **Run Script**, or use the **Scripts** button on the device page
3. Search for and select **Invoke-NinjaOffboardCapture**
4. Fill in the two Script Variables:
   - **Target Device ID** — the numeric ID from the device URL (see [Finding a Device ID](#finding-a-device-id))
   - **Offboard Reason** — type a clear explanation, e.g.:
     `Employee Jane Smith departed 2026-06-01. Laptop returned to IT stock. Agent removed.`
5. Click **Run**

The script runs in the background. Watch the output in the NinjaOne script activity log. On success you will see:

```
[✓] OFFBOARD CAPTURE COMPLETE
    Device       : DESKTOP-ABC123 (ID: 12345)
    Organization : Acme Corp (ID: 7)
    Document     : Offboard — DESKTOP-ABC123
    Template     : Device Offboard Report (ID: 3)
    Timestamp    : 2026-06-25 14:30:00

To view: NinjaOne > Organizations > Acme Corp > Documentation > Apps & Services
```

---

## The Document Template

The script automatically creates a document template called **"Device Offboard Report"** on the very first run. You do not need to build it manually.

The template has six fields:

| Field Name (internal) | Label | Type |
|---|---|---|
| `offboardReason` | Offboard Reason | WYSIWYG |
| `captureTimestamp` | Capture Date/Time | Text |
| `deviceDetails` | Device Details | WYSIWYG |
| `hardwareInfo` | Hardware Info | WYSIWYG |
| `lastActivity` | Last Activity | WYSIWYG |
| `softwareSummary` | Installed Software | WYSIWYG |

After the first run the template is visible in NinjaOne under:
**Administration → Documentation → Templates → Device Offboard Report**

You can edit the template's visual layout and field labels in the NinjaOne UI freely. Just do not rename or delete the internal field names (`offboardReason`, `captureTimestamp`, etc.) as the script references them by those names.

---

## Finding a Device ID

**Method 1 — From the URL (easiest):**
Open the device in NinjaOne. The ID is in the browser address bar:
```
https://app.ninjarmm.com/#/deviceDashboard/12345/overview
                                             ^^^^^
                                        This is the device ID
```

**Method 2 — From the device list:**
Hover over the device name in the NinjaOne device list. The ID may appear in the browser status bar.

**Method 3 — Via the API:**
Call `GET /v2/devices` and find the device by `systemName`. The `id` field is the device ID.

---

## Authentication — Why No Login Prompt?

This script uses **Client Credentials** (machine-to-machine) authentication instead of Authorization Code.

| | Authorization Code | Client Credentials *(this script)* |
|---|---|---|
| Login prompt | Yes — browser opens | ❌ None — fully silent |
| Someone must be at the machine | Yes | ❌ No |
| Works on headless/server devices | Sometimes | ✅ Always |
| Works without a browser | No | ✅ Yes |
| Audit identity | Named technician via OAuth | NinjaOne script activity log |

Because this script is triggered **from within NinjaOne by a technician**, NinjaOne's script activity log already captures who ran it and when. Client Credentials is the right choice — silent, reliable, and works on any device or environment.

---

## API Endpoints Used

| Method | Endpoint | Purpose |
|---|---|---|
| `POST` | `/ws/oauth/token` | Get access token (Client Credentials) |
| `GET` | `/v2/device/{id}` | Fetch device record |
| `GET` | `/v2/device/{id}/system-info` | Fetch hardware details |
| `GET` | `/v2/device/{id}/software` | Fetch installed software |
| `GET` | `/v2/device/{id}/activities` | Fetch recent activity |
| `GET` | `/v2/organization/{id}` | Fetch organization name |
| `GET` | `/v2/document-templates` | Check if template already exists |
| `GET` | `/v2/document-templates/{id}` | Fetch template attribute IDs for field mapping |
| `POST` | `/v2/document-templates` | Create the template (first run only) |
| `GET` | `/v2/organization/{id}/documents` | Check if a document for this device already exists |
| `POST` | `/v2/organization/documents` | Create the offboard document |
| `PATCH` | `/v2/organization/documents` | Update an existing document (if same device re-run) |

---

## Troubleshooting

| Error / Symptom | Solution |
|---|---|
| `Please fill in BaseUrl` | The `<your Login URL>` placeholder is still in the config. Replace it with your actual NinjaOne URL. |
| `Authentication failed` | Check `$BaseUrl`, `$ClientId`, `$ClientSecret`. API app must be `API Services (Machine-to-Machine)` with `monitoring` AND `management` scopes. |
| `Script Variable 'targetDeviceId' is empty` | The Script Variable name in NinjaOne must be exactly `targetDeviceId` — check for typos. |
| `Script Variable 'offboardReason' is empty` | The Script Variable name must be exactly `offboardReason`. Both must be filled before clicking Run. |
| `Device ID not found (HTTP 404)` | The number entered doesn't exist. Double-check the device URL in NinjaOne. |
| `Failed to create document template (HTTP 403)` | API app is missing the `management` scope. Edit the app in Administration → Apps → API. |
| `Failed to create document (HTTP 403)` | Same — `management` scope required for writing documentation. |
| `Failed to create document (HTTP 400)` | Template field names may not match. Delete the template in NinjaOne UI and re-run — the script will recreate it cleanly. |
| Document appears but hardware fields are empty | Normal for network devices, cloud monitors, and some agent types that don't report hardware inventory. |
| Script finishes in under 2 seconds with no output | Run As is likely not set to `System`. Edit the script in NinjaOne and confirm Run As = System. |

---

## Pre-Flight Checklist

- [ ] NinjaOne System Administrator access confirmed
- [ ] API app created — Platform: `API Services (Machine-to-Machine)`
- [ ] API app has both `monitoring` and `management` scopes
- [ ] Client ID and Client Secret saved somewhere safe
- [ ] All four variables filled in the CONFIGURATION block of the script
- [ ] Script uploaded to NinjaOne — Run As = System, Timeout = 120s
- [ ] Script Variable `targetDeviceId` created — Type: Integer, exact name match
- [ ] Script Variable `offboardReason` created — Type: Text, exact name match
- [ ] Device ID of a test device found from its NinjaOne URL
- [ ] Test run completed — confirmed `[✓] OFFBOARD CAPTURE COMPLETE` in output
- [ ] Opened Organizations → [Org] → Documentation → Apps & Services and confirmed document appeared
- [ ] "Device Offboard Report" template visible under Administration → Documentation → Templates
