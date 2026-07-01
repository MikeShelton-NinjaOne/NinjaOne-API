# NinjaOne Software Patch Reports

> Three PowerShell scripts that pull software patch data from the NinjaOne API, cross-reference installed software to attach version numbers (which the patch endpoint alone does not return), and post interactive filterable HTML reports to a NinjaOne Knowledge Base folder.

---

## The Three Scripts

| Script | Scope | KB Article Name |
|---|---|---|
| `Invoke-NinjaSoftwarePatchReport-AllOrgs.ps1` | Every device across all organizations | `Software Patch Report — All Organizations` |
| `Invoke-NinjaSoftwarePatchReport-SingleOrg.ps1` | All devices in one organization | `Software Patch Report — <OrgName>` |
| `Invoke-NinjaSoftwarePatchReport-SingleDevice.ps1` | One specific device | `Software Patch Report — <Hostname>` |

All three scripts share the same configuration format, the same authentication method, and produce the same style of interactive report. The only difference is scope and which input ID they need.

---

## Table of Contents

- [Why Version Needs Cross-Referencing](#why-version-needs-cross-referencing)
- [What the Report Shows](#what-the-report-shows)
- [Prerequisites](#prerequisites)
- [Part 1: Create the API App](#part-1-create-the-api-app)
- [Part 2: Create the Knowledge Base Folder](#part-2-create-the-knowledge-base-folder)
- [Part 3: Configure a Script](#part-3-configure-a-script)
- [Part 4: Run the Scripts](#part-4-run-the-scripts)
  - [All Organizations](#all-organizations)
  - [Single Organization](#single-organization)
  - [Single Device](#single-device)
- [Running from NinjaOne (Script Variables)](#running-from-ninjaone-script-variables)
- [Finding IDs](#finding-ids)
- [API Endpoints Used](#api-endpoints-used)
- [Troubleshooting](#troubleshooting)
- [Pre-Flight Checklist](#pre-flight-checklist)

---

## Why Version Needs Cross-Referencing

The `/v2/queries/software-patches` endpoint returns patch records with these fields:

| Field | Description |
|---|---|
| `patchId` | Internal patch identifier |
| `name` | Patch/update title |
| `identifier` / `productIdentifier` | Product code used to identify the application |
| `type` | Patch type (e.g. PATCH, UPDATE) |
| `impact` | Severity (CRITICAL, IMPORTANT, MODERATE, LOW) |
| `status` | INSTALLED, PENDING, FAILED, REJECTED, APPROVED |
| `installedAt` | Unix timestamp of when the patch was applied |
| `deviceId` | Which device this record belongs to |

**`version` is not returned by this endpoint.** It is not hidden behind `productIdentifier` — that field is a product code string, not a version number.

Version data lives on the software inventory endpoint (`/v2/queries/software` or `/v2/device/{id}/software`), which returns `name`, `version`, `publisher`, `productCode`, and `installDate` for each installed application.

These scripts cross-reference the two datasets by matching `productIdentifier` or patch `name` against the software inventory's `productCode` or `name`, then attach the `version` field to each patch record before building the report.

> ℹ️ The version shown is the **currently installed version** of that application — not necessarily the version of the specific patch file. If a newer patch has been applied since, the version reflects the current state.

---

## What the Report Shows

Each report is a self-contained interactive HTML page posted as a KB article. It includes:

**Summary stat cards:**
- Total records shown (updates as you filter)
- Installed count
- Failed count
- Pending/Approved count

**Filter controls:**
- Status dropdown (All / INSTALLED / PENDING / FAILED / REJECTED / APPROVED)
- Severity dropdown (All / CRITICAL / IMPORTANT / MODERATE / LOW)
- Organization dropdown (All Orgs version only)
- Device dropdown (Single Org version only)
- Search bar (patch name and device name)
- Reset button

**Sortable table columns:**
- Patch Name
- Version (monospace, purple-highlighted — shows `N/A` if no software inventory match found)
- Status (colour-coded badge)
- Severity (colour-coded label)
- Type
- Device (all-org and single-org versions)
- Organization (all-org version)
- Installed At

---

## Prerequisites

| Requirement | Notes |
|---|---|
| PowerShell 5.1+ | Built-in on Windows 10/11. Run `$PSVersionTable` to confirm. |
| NinjaOne System Administrator | Required to create API credentials and manage the KB |
| NinjaOne Patch Management | Must be active — scripts read patching data |
| No extra modules | Uses only built-in PowerShell |

---

## Part 1: Create the API App

One-time setup — creates the silent credentials the scripts use.

1. Log into NinjaOne as a **System Administrator**
2. Go to: **Administration → Apps → API → Client App IDs → Add**
3. Fill in:

   | Field | Value |
   |---|---|
   | **Name** | Any name, e.g. `PatchReportScript` |
   | **Platform** | `API Services (Machine-to-Machine)` |
   | **Allowed Scopes** | ✅ `monitoring` AND ✅ `management` |
   | **Redirect URI** | Leave blank |

4. Click **Save** — copy the **Client ID** and **Client Secret** immediately

> ⚠️ The Client Secret is shown only once. Store it securely.

---

## Part 2: Create the Knowledge Base Folder

1. In NinjaOne, open the **Knowledge Base** from the left sidebar
2. Click **New Folder**
3. Name it, e.g. `Patch Reports`
4. Open the folder and look at the browser URL:

   ```
   https://app.ninjarmm.com/#/knowledgeBase/folder/42
                                                     ^^  ← Folder ID
   ```

5. Copy that number — you need it in the config

---

## Part 3: Configure a Script

Open the script in any text editor. Fill in the **CONFIGURATION** block at the top:

```powershell
# ==============================================================================
#  CONFIGURATION — Fill in ALL five values before running
# ==============================================================================

$BaseUrl       = 'https://<your Login URL>'
$TokenEndpoint = 'https://<your Login URL>/ws/oauth/token'
$ClientId      = '<Your Client ID>'
$ClientSecret  = '<Your Client Secret>'
$KbFolderId    = 0    # <-- Replace with your KB folder ID number
```

| Variable | Example | Notes |
|---|---|---|
| `$BaseUrl` | `https://app.ninjarmm.com` | Your NinjaOne login URL — no trailing slash |
| `$TokenEndpoint` | `https://app.ninjarmm.com/ws/oauth/token` | Same as `$BaseUrl` + `/ws/oauth/token` |
| `$ClientId` | `abc123...` | From the API app in Part 1 |
| `$ClientSecret` | `s3cr3t...` | From the API app — shown once |
| `$KbFolderId` | `42` | From the KB folder URL in Part 2 |

**Regional URLs:**

| Region | Base URL |
|---|---|
| United States | `https://app.ninjarmm.com` |
| Europe | `https://eu.ninjarmm.com` |
| Oceania | `https://oc.ninjarmm.com` |
| Canada | `https://ca.ninjarmm.com` |

---

## Part 4: Run the Scripts

### All Organizations

Pulls every patch record across your entire NinjaOne instance.

```powershell
.\Invoke-NinjaSoftwarePatchReport-AllOrgs.ps1
```

No additional parameters needed. The report groups by org and device in the filter dropdowns.

> ⚠️ This can take several minutes in large environments — it pages through all devices, all software records, and all patch records. Expect 3–10 minutes for 500+ devices.

---

### Single Organization

Pulls patch data for all devices belonging to one organization.

**Option A — Set it directly in the script:**
```powershell
# In the config block, set:
$ManualOrgId = 99   # your org ID
```
Then run:
```powershell
.\Invoke-NinjaSoftwarePatchReport-SingleOrg.ps1
```

**Option B — Pass it as a parameter (or NinjaOne Script Variable):**
The script also reads `$env:targetOrgId` automatically when triggered from NinjaOne.

---

### Single Device

Pulls patch data for one specific device.

**Option A — Set it directly in the script:**
```powershell
# In the config block, set:
$ManualDeviceId = 12345   # your device ID
```
Then run:
```powershell
.\Invoke-NinjaSoftwarePatchReport-SingleDevice.ps1
```

**Option B — NinjaOne Script Variable:**
The script reads `$env:targetDeviceId` when triggered from NinjaOne.

---

## Running from NinjaOne (Script Variables)

All three scripts can be uploaded to NinjaOne and triggered with Script Variables, so technicians can run them on demand from the portal.

### Upload a script

1. Go to: **Administration → Scripting → Scripts → Add Script**
2. Fill in: Name, Language = PowerShell, Run As = System, Timeout = 300 seconds
3. Paste the full script content and click Save

### Script Variables needed per script

**All Organizations** — no Script Variables needed

**Single Organization:**

| Name | Type | Label |
|---|---|---|
| `targetOrgId` | Integer | Target Organization ID |

**Single Device:**

| Name | Type | Label |
|---|---|---|
| `targetDeviceId` | Integer | Target Device ID |

> ⚠️ Variable names must be exact — `targetOrgId` and `targetDeviceId`. The scripts read them as environment variables by those exact names.

### Running a script in NinjaOne

1. Navigate to any managed device
2. Right-click → **Run Script** → select the script
3. Fill in the Script Variable if prompted
4. Click **Run** — monitor output in the script activity log

---

## Finding IDs

**Organization ID:**
Open the org in NinjaOne and look at the URL:
```
https://app.ninjarmm.com/#/customerDashboard/99/overview
                                              ^^  ← Org ID
```

**Device ID:**
Open the device in NinjaOne and look at the URL:
```
https://app.ninjarmm.com/#/deviceDashboard/12345/overview
                                            ^^^^^  ← Device ID
```

**KB Folder ID:**
Open the folder in the Knowledge Base:
```
https://app.ninjarmm.com/#/knowledgeBase/folder/42
                                                 ^^  ← Folder ID
```

---

## API Endpoints Used

| Script | Method | Endpoint | Purpose |
|---|---|---|---|
| All | `POST` | `/ws/oauth/token` | Client Credentials auth |
| All Orgs | `GET` | `/v2/organizations` | Org names (paginated) |
| All Orgs | `GET` | `/v2/devices` | Device names and org mapping (paginated) |
| All Orgs | `GET` | `/v2/queries/software` | Software inventory for version data (paginated) |
| All Orgs | `GET` | `/v2/queries/software-patches` | Patch records (paginated) |
| Single Org | `GET` | `/v2/organization/{id}` | Org name |
| Single Org | `GET` | `/v2/organization/{id}/devices` | Devices in org |
| Single Org | `GET` | `/v2/queries/software?organizationId=` | Software inventory scoped to org |
| Single Org | `GET` | `/v2/queries/software-patches?organizationId=` | Patches scoped to org |
| Single Device | `GET` | `/v2/device/{id}` | Device name and org ID |
| Single Device | `GET` | `/v2/organization/{id}` | Org name |
| Single Device | `GET` | `/v2/device/{id}/software` | Software inventory for this device |
| Single Device | `GET` | `/v2/device/{id}/software-patches` | Patches for this device |
| All | `GET` | `/v2/knowledgebase/global/articles?folderId=` | Check if article exists |
| All | `POST` | `/v2/knowledgebase/articles` | Create KB article (first run) |
| All | `PATCH` | `/v2/knowledgebase/article/{id}` | Update KB article (subsequent runs) |

---

## Troubleshooting

| Error / Symptom | Solution |
|---|---|
| `Fill in BaseUrl` | The `<your Login URL>` placeholder is still in the config. |
| `Fill in KbFolderId` | Change `$KbFolderId = 0` to your actual folder ID. |
| `Fill in credentials` | Replace the `<Your Client ID>` and `<Your Client Secret>` placeholders. |
| `Authentication failed` | Check URL, Client ID, Client Secret. API app must be `API Services (Machine-to-Machine)` with `monitoring` AND `management` scopes. |
| Device/Org not found (HTTP 404) | The ID doesn't exist in NinjaOne. Double-check from the URL. |
| `Failed to post to KB (HTTP 403)` | API app is missing the `management` scope. |
| `Failed to post to KB (HTTP 404)` | `$KbFolderId` doesn't exist. Verify from the folder URL. |
| Version shows `N/A` for all patches | The software inventory endpoint returned no data. Check that the device reports software to NinjaOne. Some device types (network devices, cloud monitors) may not have software inventory. |
| Version shows `N/A` for some patches | Normal — some patches don't have a matching `productCode` or `name` in the software inventory. This typically affects update-type patches that aren't tracked as discrete installed applications. |
| All Orgs script is very slow | Expected for large environments. Each page of data requires API calls. 1,000+ devices may take 10+ minutes. Consider scheduling this off-hours. |
| Report appears as raw HTML in KB | NinjaOne renders HTML in KB article content. If you see code instead of the rendered page, contact NinjaOne support to confirm HTML KB content is enabled on your plan. |
| Duplicate articles appearing | The script checks for an existing article by exact name before creating. If duplicates exist, manually delete the extras in the KB and the script will find the correct one on next run. |

---

## Pre-Flight Checklist

- [ ] NinjaOne System Administrator access confirmed
- [ ] NinjaOne Patch Management enabled on your account
- [ ] API app created — Platform: `API Services (Machine-to-Machine)`, scopes: `monitoring` + `management`
- [ ] Client ID and Client Secret saved securely
- [ ] KB folder created and folder ID copied from the URL
- [ ] All five config variables filled in (no `<placeholder>` text remaining, `$KbFolderId` ≠ 0)
- [ ] For Single Org: Org ID confirmed from the org URL
- [ ] For Single Device: Device ID confirmed from the device URL
- [ ] Test run completed — confirmed `[✓] COMPLETE` output
- [ ] KB article confirmed visible in NinjaOne: Knowledge Base → your folder
