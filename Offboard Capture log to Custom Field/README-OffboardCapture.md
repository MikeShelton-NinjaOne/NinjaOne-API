# Invoke-NinjaOffboardCapture.ps1

> Captures device information and an offboard reason via the NinjaOne API, automatically creates an organization-level custom field named after the device hostname, and writes a structured offboard report to that field — all triggered silently from within NinjaOne with no browser login required.

---

## Table of Contents

- [What This Script Does](#what-this-script-does)
- [How It Works (Plain English)](#how-it-works-plain-english)
- [Prerequisites](#prerequisites)
- [Part 1: Create the API App in NinjaOne](#part-1-create-the-api-app-in-ninjaone)
- [Part 2: Configure the Script](#part-2-configure-the-script)
- [Part 3: Add the Script to NinjaOne](#part-3-add-the-script-to-ninjaone)
- [Part 4: Create the Script Variables](#part-4-create-the-script-variables)
- [Part 5: Run the Script](#part-5-run-the-script)
- [Finding a Device ID](#finding-a-device-id)
- [Where to Find the Output](#where-to-find-the-output)
- [What Gets Captured](#what-gets-captured)
- [Authentication: Why No Browser Login?](#authentication-why-no-browser-login)
- [API Endpoints Used](#api-endpoints-used)
- [Troubleshooting](#troubleshooting)
- [Pre-Flight Checklist](#pre-flight-checklist)

---

## What This Script Does

When a technician offboards a device, this script:

1. **Authenticates silently** to NinjaOne using an API Client ID and Secret (no browser, no login prompt)
2. **Pulls detailed device information** — hostname, IP, OS, hardware, last contact, recent activity, installed software
3. **Takes the offboard reason** you enter when running the script
4. **Creates a new organization-level custom field** named after the device (e.g. `desktopAbc123`) if it doesn't already exist
5. **Writes a structured offboard report** to that custom field on the organization record

The result is a permanent, searchable record attached to the organization in NinjaOne showing exactly what the device looked like and why it was offboarded.

---

## How It Works (Plain English)

Think of it like filling out a form — you tell the script which device you're offboarding (by ID) and why, and it automatically:

- Goes out and grabs all the device's information from NinjaOne
- Creates a labeled storage slot on the organization (the custom field)
- Writes everything into that slot so it lives permanently in NinjaOne

You don't need to log in anywhere, open a browser, or copy anything manually.

---

## Prerequisites

| Requirement | Notes |
|---|---|
| NinjaOne account | Must be a **System Administrator** to create API credentials and scripts |
| PowerShell 5.1+ | Built-in on Windows 10/11. Run `$PSVersionTable` to check your version. |
| No extra modules needed | This script uses only built-in PowerShell commands |

---

## Part 1: Create the API App in NinjaOne

This is a **one-time setup**. You create an API application that gives the script permission to talk to NinjaOne on your behalf — silently, with no login prompt.

1. Log into NinjaOne as a **System Administrator**
2. Go to: **Administration → Apps → API → Client App IDs → Add**
3. Fill in the form:

   | Field | What to Enter |
   |---|---|
   | **Name** | Any name you like, e.g. `OffboardCaptureScript` |
   | **Platform** | `API Services (Machine-to-Machine)` |
   | **Allowed Scopes** | Check both `monitoring` and `management` |
   | **Redirect URI** | Leave blank — not needed for this flow |

4. Click **Save**
5. You will see a **Client ID** and **Client Secret** — **copy both and save them somewhere safe**

> ⚠️ The Client Secret is only shown once. If you lose it, you will need to regenerate it from the portal.

> ℹ️ **No Redirect URI is needed.** Client Credentials (machine-to-machine) authentication does not use a browser or redirect — that is the whole point. This is different from the Authorization Code flow used in some other scripts.

---

## Part 2: Configure the Script

Open `Invoke-NinjaOffboardCapture.ps1` in any text editor (Notepad works fine). Find the **CONFIGURATION** block near the top — it looks like this:

```powershell
# ==============================================================================
#  CONFIGURATION — Fill in ALL four values below before saving/running this script
# ==============================================================================

$BaseUrl       = 'https://<your Login URL>'
$TokenEndpoint = 'https://<your Login URL>/ws/oauth/token'
$ClientId      = '<Your Client ID>'
$ClientSecret  = '<Your Client Secret>'
```

Replace each placeholder with your real values:

| Variable | Example | Where to Find It |
|---|---|---|
| `$BaseUrl` | `https://app.ninjarmm.com` | Your NinjaOne login URL (see regional table below) |
| `$TokenEndpoint` | `https://app.ninjarmm.com/ws/oauth/token` | Same as `$BaseUrl`, just add `/ws/oauth/token` at the end |
| `$ClientId` | `abc123xyz...` | Shown after creating the app in Step 1 |
| `$ClientSecret` | `s3cr3tK3y...` | Shown once at app creation — save it immediately |

**Regional URLs:**

| Region | Base URL |
|---|---|
| United States | `https://app.ninjarmm.com` |
| Europe | `https://eu.ninjarmm.com` |
| Oceania | `https://oc.ninjarmm.com` |
| Canada | `https://ca.ninjarmm.com` |

> ⚠️ Do not commit the `$ClientSecret` value to a public Git repository. Treat it like a password.

---

## Part 3: Add the Script to NinjaOne

1. In NinjaOne, go to: **Administration → Scripting → Scripts**
2. Click **Add Script** (or the `+` button)
3. Fill in the script details:

   | Field | What to Enter |
   |---|---|
   | **Name** | `Invoke-NinjaOffboardCapture` |
   | **Language** | `PowerShell` |
   | **Script Body** | Paste the entire contents of `Invoke-NinjaOffboardCapture.ps1` |
   | **Description** | `Captures offboard data and writes it to an org-level custom field` |
   | **Run As** | `System` |
   | **Timeout** | `120` seconds (2 minutes is plenty) |

4. Click **Save**

---

## Part 4: Create the Script Variables

Script Variables are the fields a technician fills in when they run the script. You need to create two.

Go to: **Administration → Scripting → Script Variables**

### Variable 1 — Target Device ID

Click **Add Script Variable** and fill in:

| Field | Value |
|---|---|
| **Name** | `targetDeviceId` ← must be exact, including capitalization |
| **Label** | `Target Device ID` |
| **Type** | `Integer` |
| **Required** | Yes |
| **Description** | `The numeric ID of the device being offboarded. Find it in the device's NinjaOne URL.` |

### Variable 2 — Offboard Reason

Click **Add Script Variable** and fill in:

| Field | Value |
|---|---|
| **Name** | `offboardReason` ← must be exact, including capitalization |
| **Label** | `Offboard Reason` |
| **Type** | `Text` |
| **Required** | Yes |
| **Description** | `Explain why this device is being offboarded (e.g. employee departure, hardware failure, retired asset).` |

> ⚠️ The variable **Names** (`targetDeviceId` and `offboardReason`) must be spelled exactly as shown. The script reads them by those exact names from environment variables. If the name doesn't match, the script will stop with an error.

---

## Part 5: Run the Script

1. In NinjaOne, navigate to the **device** you want to run the script against
   - This can be the device being offboarded, or any managed device — the script targets `targetDeviceId`, not the device it physically runs on
2. Click **Run Script** (or right-click the device → Run Script)
3. Search for and select **Invoke-NinjaOffboardCapture**
4. Fill in the two Script Variables that appear:
   - **Target Device ID**: Enter the numeric ID of the device being offboarded (see [Finding a Device ID](#finding-a-device-id))
   - **Offboard Reason**: Type a clear reason, e.g. `Employee John Smith departed on 2026-06-01. Device returned to IT stock.`
5. Click **Run**

The script will execute and you can watch the output in the NinjaOne script activity log.

---

## Finding a Device ID

You need the numeric device ID to fill in the `targetDeviceId` variable. Here are two ways to find it:

**Method 1 — From the URL:**
Open the device in NinjaOne. Look at the URL in your browser:
```
https://app.ninjarmm.com/#/deviceDashboard/12345/overview
                                             ^^^^^
                                         This is the device ID
```

**Method 2 — From the API:**
Use `GET /v2/devices` to list all devices and find the one you want. The `id` field in the response is the device ID.

**Method 3 — From the device list:**
In NinjaOne, hover over the device name in the device list. The ID may appear in the status bar or tooltip depending on your browser.

---

## Where to Find the Output

After the script runs successfully:

1. Go to **NinjaOne → Organizations**
2. Open the organization the device belongs to
3. Click the **Custom Fields** tab (or look for a custom tab that was created)
4. Look for a field labeled **`Offboard - <DeviceHostname>`**

The field will contain the full structured offboard report.

> ℹ️ If you have many offboarded devices, each one gets its own custom field on the organization, named after the device hostname. They are all visible on the organization's Custom Fields tab.

---

## What Gets Captured

The offboard report written to the custom field includes:

| Section | Data Points |
|---|---|
| **Header** | Capture timestamp, device name, device ID, organization name and ID |
| **Offboard Reason** | Exactly what the technician typed into the Script Variable |
| **Device Details** | System name, DNS name, IP addresses, OS name and service pack, device class, last contact time, NinjaOne agent version |
| **Hardware** *(if available)* | Manufacturer, model, serial number, CPU, RAM |
| **Last Recorded Activity** *(if available)* | Timestamp, activity type, message |
| **Installed Software** | First 20 installed applications with version numbers |

---

## Authentication: Why No Browser Login?

This script uses **Client Credentials** (also called machine-to-machine) authentication instead of Authorization Code flow.

The difference:

| | Authorization Code | Client Credentials (this script) |
|---|---|---|
| **Login prompt** | Yes — browser opens | ❌ None — fully silent |
| **Requires someone at the machine** | Yes | ❌ No |
| **Works on headless/server devices** | Sometimes not | ✅ Always |
| **Identity tracked** | Named technician | API app identity |
| **Audit trail** | OAuth log + NinjaOne log | NinjaOne script activity log |

Because this script is triggered **from within NinjaOne by a technician**, NinjaOne's own activity log already records who ran it and when. The OAuth layer does not need to capture identity too — so silent Client Credentials is the better fit. It is more reliable, works on any device type, and removes the dependency on a browser being accessible on the endpoint.

---

## API Endpoints Used

| Method | Endpoint | Purpose |
|---|---|---|
| `POST` | `/ws/oauth/token` | Get access token via Client Credentials |
| `GET` | `/v2/device/{id}` | Fetch device record (hostname, IP, OS, last contact) |
| `GET` | `/v2/device/{id}/system-info` | Fetch hardware details (CPU, RAM, serial number) |
| `GET` | `/v2/device/{id}/software` | Fetch installed software list |
| `GET` | `/v2/device/{id}/activities` | Fetch recent device activity |
| `GET` | `/v2/organization/{id}` | Fetch organization name |
| `GET` | `/v2/custom-fields?scope=organization` | Check if the custom field already exists |
| `POST` | `/v2/custom-fields` | Create the new org-scoped custom field |
| `PATCH` | `/v2/organization/{id}/custom-fields` | Write the offboard report to the field |

---

## Troubleshooting

| Error / Symptom | Solution |
|---|---|
| `Authentication failed` | Check `$BaseUrl`, `$ClientId`, `$ClientSecret`. Verify the API app has `monitoring` and `management` scopes. |
| `Script Variable 'targetDeviceId' is empty` | Make sure the Script Variable name in NinjaOne is exactly `targetDeviceId` (camelCase, no spaces). |
| `Script Variable 'offboardReason' is empty` | Make sure the Script Variable name is exactly `offboardReason`. Both variables must be filled in before running. |
| `Device ID not found (404)` | The ID entered in `targetDeviceId` doesn't exist. Double-check by looking at the device URL in NinjaOne. |
| `Failed to create custom field (403)` | The API app is missing the `management` scope. Edit the app in Administration → Apps → API and add it. |
| `Failed to write to custom field (400)` | The field name may contain characters NinjaOne doesn't accept. Check the script output for the generated field name and look for unusual characters in the hostname. |
| `Failed to write to custom field (404)` | The field was just created and hasn't propagated yet. The script includes a 3-second wait — if you see this, try running the script again (it will skip creation and go straight to writing). |
| Custom field appears on device, not org | Confirm `definitionScopes` is set to `ORGANIZATION` in the script. Do not change this value. |
| Script runs but no output visible | Check the NinjaOne script activity log for the device. Expand the run entry to see `Write-Host` output. |

---

## Pre-Flight Checklist

Run through this before your first use:

- [ ] NinjaOne System Administrator access confirmed
- [ ] API app created in NinjaOne with platform `API Services (Machine-to-Machine)`
- [ ] API app has both `monitoring` and `management` scopes
- [ ] Client ID and Client Secret copied and stored safely
- [ ] `$BaseUrl`, `$TokenEndpoint`, `$ClientId`, `$ClientSecret` all filled in the script
- [ ] Script uploaded to NinjaOne (Administration → Scripting → Scripts)
- [ ] Script Variable `targetDeviceId` created (Type: Integer, Name exact match)
- [ ] Script Variable `offboardReason` created (Type: Text, Name exact match)
- [ ] Device ID of the test device identified from its NinjaOne URL
- [ ] Test run completed on a non-critical device before using in production
- [ ] Output verified in Organization → Custom Fields tab after test run
