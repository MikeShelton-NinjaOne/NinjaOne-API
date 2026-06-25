# Invoke-NinjaOffboardCapture-KB.ps1

> Captures full device information and an offboard reason, then creates or updates a formatted **Knowledge Base article** in NinjaOne. Each offboarded device gets its own named article — permanently searchable in the KB folder you choose.

---

## Table of Contents

- [What This Script Does](#what-this-script-does)
- [Where the Data Goes](#where-the-data-goes)
- [What Gets Captured](#what-gets-captured)
- [How It Differs from the Apps & Services Version](#how-it-differs-from-the-apps--services-version)
- [Prerequisites](#prerequisites)
- [Part 1: Create the API App](#part-1-create-the-api-app)
- [Part 2: Create the Knowledge Base Folder](#part-2-create-the-knowledge-base-folder)
- [Part 3: Configure the Script](#part-3-configure-the-script)
- [Part 4: Add the Script to NinjaOne](#part-4-add-the-script-to-ninjaone)
- [Part 5: Create the Script Variables](#part-5-create-the-script-variables)
- [Part 6: Run the Script](#part-6-run-the-script)
- [Finding a Device ID](#finding-a-device-id)
- [Finding a KB Folder ID](#finding-a-kb-folder-id)
- [Authentication — Why No Login Prompt?](#authentication--why-no-login-prompt)
- [API Endpoints Used](#api-endpoints-used)
- [Troubleshooting](#troubleshooting)
- [Pre-Flight Checklist](#pre-flight-checklist)

---

## What This Script Does

When a technician offboards a device, this script:

1. **Authenticates silently** using Client Credentials — no browser, no login prompt
2. **Pulls device data** — hostname, IP, OS, hardware, last contact, installed software, recent activity
3. **Takes the offboard reason** entered as a NinjaOne Script Variable
4. **Builds a fully formatted HTML report** with colour-coded sections and tables
5. **Checks whether a KB article already exists** for this device in the configured folder — creates it if new, updates it if it already exists

The result is a clean, readable article sitting permanently in your NinjaOne Knowledge Base, viewable by any technician with KB access.

---

## Where the Data Goes

```
NinjaOne → Knowledge Base → [Your Folder] → Offboard Report — HOSTNAME
```

Each offboarded device gets its own article named `Offboard Report — HOSTNAME`. Running the script on the same device twice updates the article rather than creating a duplicate.

---

## What Gets Captured

The KB article contains a formatted HTML report with five sections:

| Section | Contents |
|---|---|
| **Header** | Device hostname, organization name, capture timestamp |
| **Offboard Reason** | Exactly what the technician entered when triggering the script |
| **Device Details** | System name, DNS name, IP addresses, OS, device class, last contact, agent version |
| **Hardware Info** | Manufacturer, model, serial number, CPU, RAM *(where available)* |
| **Last Activity** | Most recent activity — timestamp, type, message |
| **Installed Software** | First 20 installed applications with version numbers |

---

## How It Differs from the Apps & Services Version

You may also have `Invoke-NinjaOffboardCapture.ps1` — the version that writes to an Apps & Services document. Here is when to use each:

| | KB Version *(this script)* | Apps & Services Version |
|---|---|---|
| **Where it lives** | Knowledge Base → folder | Organization → Documentation → Apps & Services |
| **Who can see it** | Anyone with KB access | Anyone viewing the specific org |
| **Searchable in KB** | ✅ Yes — full KB search | ❌ No |
| **Grouped by org** | ❌ All in one folder | ✅ Attached to the org record |
| **Template required** | ❌ No | ✅ Auto-created but required |
| **Best for** | Central searchable archive | Org-level documentation |

Use this script (KB version) if you want a **central searchable repository** of all offboard records. Use the Apps & Services version if you want offboard records attached to the **organization they belong to**.

---

## Prerequisites

| Requirement | Notes |
|---|---|
| PowerShell 5.1+ | Built-in on Windows 10/11. Run `$PSVersionTable` to check. |
| NinjaOne System Administrator | Required to create API credentials, upload scripts, create Script Variables, and manage the KB |
| NinjaOne Knowledge Base | Must be enabled on your account |
| No extra modules | Uses only built-in PowerShell — nothing to install |

---

## Part 1: Create the API App

One-time setup — creates the silent credentials the script uses to authenticate.

1. Log into NinjaOne as a **System Administrator**
2. Go to: **Administration → Apps → API → Client App IDs → Add**
3. Fill in the form:

   | Field | What to Enter |
   |---|---|
   | **Name** | Any name, e.g. `OffboardCaptureKB` |
   | **Platform** | `API Services (Machine-to-Machine)` |
   | **Allowed Scopes** | Check ✅ `monitoring` AND ✅ `management` |
   | **Redirect URI** | Leave blank |

4. Click **Save**
5. Copy the **Client ID** and **Client Secret** — the secret is only shown once

> ⚠️ If you lose the Client Secret, delete the app and create a new one.

---

## Part 2: Create the Knowledge Base Folder

The script needs an existing KB folder to post articles into.

1. In NinjaOne, open the **Knowledge Base** from the left sidebar
2. Click **New Folder**
3. Name it something clear, e.g. `Device Offboard Records`
4. Click **Save**
5. Open the folder and look at the browser URL:

   ```
   https://app.ninjarmm.com/#/knowledgeBase/folder/42
                                                     ^^
                                              This is your Folder ID
   ```

6. Copy that number — you need it in the next step

---

## Part 3: Configure the Script

Open `Invoke-NinjaOffboardCapture-KB.ps1` in any text editor. Find the **CONFIGURATION** block at the very top and fill in all six values:

```powershell
# ==============================================================================
#  CONFIGURATION — Fill in ALL six values below before saving/running the script
# ==============================================================================

$BaseUrl       = 'https://<your Login URL>'
$TokenEndpoint = 'https://<your Login URL>/ws/oauth/token'
$ClientId      = '<Your Client ID>'
$ClientSecret  = '<Your Client Secret>'
$KbFolderId    = 0        # <-- Replace with your KB folder ID number
$ArticlePrefix = 'Offboard Report'
```

| Variable | Example | Notes |
|---|---|---|
| `$BaseUrl` | `https://app.ninjarmm.com` | Your NinjaOne login URL — no trailing slash |
| `$TokenEndpoint` | `https://app.ninjarmm.com/ws/oauth/token` | Same as `$BaseUrl` plus `/ws/oauth/token` |
| `$ClientId` | `abc123...` | From the API app created in Part 1 |
| `$ClientSecret` | `s3cr3t...` | From the API app — shown once at creation |
| `$KbFolderId` | `42` | The number from the KB folder URL (Part 2) |
| `$ArticlePrefix` | `Offboard Report` | Article names become `Offboard Report — HOSTNAME` |

**Regional URLs:**

| Region | Base URL |
|---|---|
| United States | `https://app.ninjarmm.com` |
| Europe | `https://eu.ninjarmm.com` |
| Oceania | `https://oc.ninjarmm.com` |
| Canada | `https://ca.ninjarmm.com` |

> ⚠️ Do not commit `$ClientSecret` to source control. Treat it like a password.

---

## Part 4: Add the Script to NinjaOne

1. Go to: **Administration → Scripting → Scripts → Add Script**
2. Fill in:

   | Field | Value |
   |---|---|
   | **Name** | `Invoke-NinjaOffboardCapture-KB` |
   | **Language** | `PowerShell` |
   | **Run As** | `System` |
   | **Timeout** | `120` seconds |
   | **Script Body** | Paste the entire contents of the `.ps1` file |

3. Click **Save**

---

## Part 5: Create the Script Variables

Script Variables are the two inputs a technician fills in each time they run the script.

Go to: **Administration → Scripting → Script Variables**

### Variable 1 — Target Device ID

| Field | Value |
|---|---|
| **Name** | `targetDeviceId` ← exact spelling and capitalisation required |
| **Label** | `Target Device ID` |
| **Type** | `Integer` |
| **Required** | Yes |
| **Description** | `The numeric ID of the device being offboarded. Find it in the device's NinjaOne URL.` |

### Variable 2 — Offboard Reason

| Field | Value |
|---|---|
| **Name** | `offboardReason` ← exact spelling and capitalisation required |
| **Label** | `Offboard Reason` |
| **Type** | `Text` |
| **Required** | Yes |
| **Description** | `Why is this device being offboarded? E.g. employee departure, hardware failure, end of life.` |

> ⚠️ The variable **Names** must be exactly `targetDeviceId` and `offboardReason`. The script reads them as environment variables by those exact names. A typo means the script stops immediately with an error.

---

## Part 6: Run the Script

1. In NinjaOne, navigate to **any managed device**
   > The script uses `targetDeviceId` to find the device via the API — it does not have to run *on* that device. It can run on any managed device.
2. Right-click the device → **Run Script**, or use the **Scripts** button
3. Search for and select **Invoke-NinjaOffboardCapture-KB**
4. Fill in the two Script Variables:
   - **Target Device ID** — the numeric ID from the device URL
   - **Offboard Reason** — a clear description, e.g.:
     `Employee Jane Smith left on 2026-06-01. Laptop wiped and returned to stock. Agent removed.`
5. Click **Run**

Watch the output in the NinjaOne script activity log:

```
  ============================================================
  NinjaOne Offboard Capture  →  Knowledge Base
  Device ID  : 12345
  KB Folder  : 42
  ============================================================

  [1/5] Authenticating (Client Credentials)...
  [✓] Authenticated.

  [2/5] Fetching device information (ID: 12345)...
  [✓] Device: DESKTOP-ABC123  (Org ID: 7)

  [3/5] Fetching organization name (ID: 7)...
  [✓] Organization: Acme Corp

  [4/5] Building HTML report...
  [✓] HTML report built.

  [5/5] Posting to Knowledge Base folder (ID: 42)...
  [✓] KB article created.

  ============================================================
  [✓] OFFBOARD CAPTURE COMPLETE
      Device      : DESKTOP-ABC123 (ID: 12345)
      Org         : Acme Corp (ID: 7)
      KB Folder   : 42
      Article     : Offboard Report — DESKTOP-ABC123
      Timestamp   : 2026-06-25 14:30:00
  ============================================================

  To view: NinjaOne > Knowledge Base > your folder > Offboard Report — DESKTOP-ABC123
```

6. Open **Knowledge Base** in NinjaOne → your folder → look for `Offboard Report — DESKTOP-ABC123`

---

## Finding a Device ID

**From the browser URL (easiest):**
Open the device in NinjaOne and look at the address bar:

```
https://app.ninjarmm.com/#/deviceDashboard/12345/overview
                                             ^^^^^
                                        This number is the device ID
```

**From the device list:**
Hover over a device name — the ID may appear in the browser status bar.

**Via the API:**
Call `GET /v2/devices` and find the device by `systemName`. The `id` field is the device ID.

---

## Finding a KB Folder ID

Open the Knowledge Base in NinjaOne and navigate into the folder you created. Look at the browser URL:

```
https://app.ninjarmm.com/#/knowledgeBase/folder/42
                                                  ^^
                                           This is the folder ID
```

If you haven't created a folder yet, go to **Knowledge Base → New Folder**, name it, save it, then open it to get the ID.

---

## Authentication — Why No Login Prompt?

This script uses **Client Credentials** (machine-to-machine) — not Authorization Code.

| | Authorization Code | Client Credentials *(this script)* |
|---|---|---|
| Login prompt | Yes — browser opens | ❌ None — fully silent |
| Requires someone at the machine | Yes | ❌ No |
| Works on any device type | Sometimes | ✅ Always |
| Works without a browser | No | ✅ Yes |
| Audit trail | Named user via OAuth | NinjaOne script activity log |

Because this script is triggered **from within NinjaOne by a technician**, NinjaOne already logs who ran it and when. Client Credentials is the correct choice — silent, reliable, and compatible with all device types.

---

## API Endpoints Used

| Method | Endpoint | Purpose |
|---|---|---|
| `POST` | `/ws/oauth/token` | Get access token (Client Credentials) |
| `GET` | `/v2/device/{id}` | Fetch device record |
| `GET` | `/v2/device/{id}/system-info` | Fetch hardware details |
| `GET` | `/v2/device/{id}/software` | Fetch installed software list |
| `GET` | `/v2/device/{id}/activities` | Fetch most recent activity |
| `GET` | `/v2/organization/{id}` | Fetch organization name |
| `GET` | `/v2/knowledgebase/global/articles?folderId=` | Check if article already exists |
| `POST` | `/v2/knowledgebase/articles` | Create the KB article (first run) |
| `PATCH` | `/v2/knowledgebase/article/{id}` | Update the KB article (subsequent runs) |

---

## Troubleshooting

| Error / Symptom | Solution |
|---|---|
| `Please fill in BaseUrl` | The `<your Login URL>` placeholder is still in the config. Replace with your actual URL. |
| `Please set KbFolderId` | Change `$KbFolderId = 0` to your actual folder ID from the KB folder URL. |
| `Authentication failed` | Check `$BaseUrl`, `$ClientId`, `$ClientSecret`. API app must be `API Services (Machine-to-Machine)` with `monitoring` AND `management` scopes. |
| `Script Variable 'targetDeviceId' is empty` | The Script Variable name in NinjaOne must be exactly `targetDeviceId` — check capitalisation. |
| `Script Variable 'offboardReason' is empty` | The Script Variable name must be exactly `offboardReason`. |
| `Device ID not found (HTTP 404)` | The number entered doesn't match any device. Check the device URL in NinjaOne. |
| `Failed to create KB article (HTTP 403)` | API app is missing the `management` scope. Edit it in Administration → Apps → API. |
| `Failed to create KB article (HTTP 404)` | `$KbFolderId` doesn't exist. Verify the folder ID from the URL, or create the folder first. |
| Article appears in KB but hardware section shows "Not available" | Normal for network devices, cloud monitors, and some agent types that don't report hardware inventory to NinjaOne. |
| Article content appears as raw HTML text | NinjaOne renders HTML in KB articles — if you see raw tags, your NinjaOne plan may not support HTML article content. Contact NinjaOne support. |
| Script finishes instantly with no output | Run As is likely not `System`. Edit the script in NinjaOne and set Run As = System. |

---

## Pre-Flight Checklist

- [ ] NinjaOne System Administrator access confirmed
- [ ] Knowledge Base is enabled on your NinjaOne account
- [ ] API app created — Platform: `API Services (Machine-to-Machine)`
- [ ] API app has both `monitoring` and `management` scopes
- [ ] Client ID and Client Secret copied and stored safely
- [ ] KB folder created and folder ID copied from the URL
- [ ] All six variables filled in the CONFIGURATION block — `$KbFolderId` is not 0
- [ ] Script uploaded to NinjaOne — Run As = System, Timeout = 120s
- [ ] Script Variable `targetDeviceId` created — Type: Integer, exact name
- [ ] Script Variable `offboardReason` created — Type: Text, exact name
- [ ] Test run completed on a real device — confirmed `[✓] OFFBOARD CAPTURE COMPLETE`
- [ ] Opened Knowledge Base in NinjaOne and confirmed article appeared in the folder
