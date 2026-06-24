# Invoke-NinjaTimeReport.ps1

> Generates an interactive HTML report showing how much time each technician has spent on tickets — broken down per organization — and posts it directly to a NinjaOne Knowledge Base folder. Runs silently with no login prompt required.

---

## Table of Contents

- [What This Script Does](#what-this-script-does)
- [What the Report Looks Like](#what-the-report-looks-like)
- [Prerequisites](#prerequisites)
- [Part 1: Create the API App](#part-1-create-the-api-app)
- [Part 2: Create the Knowledge Base Folder](#part-2-create-the-knowledge-base-folder)
- [Part 3: Configure the Script](#part-3-configure-the-script)
- [Part 4: Run the Script](#part-4-run-the-script)
- [Scheduling the Report](#scheduling-the-report)
- [How Time Tracking is Calculated](#how-time-tracking-is-calculated)
- [Do I Need to Change the Grant Type?](#do-i-need-to-change-the-grant-type)
- [API Endpoints Used](#api-endpoints-used)
- [Troubleshooting](#troubleshooting)
- [Pre-Flight Checklist](#pre-flight-checklist)

---

## What This Script Does

1. **Authenticates silently** using Client Credentials — no browser, no login prompt
2. **Pulls all technician accounts** from your NinjaOne instance
3. **Pulls all organizations**
4. **Fetches all tickets** updated in the last 90 days (paginated — handles large environments)
5. **Reads time-tracking log entries** from each ticket that has recorded time
6. **Aggregates hours per technician and per organization**
7. **Builds a self-contained interactive HTML report** with filters and search
8. **Posts the HTML to a NinjaOne Knowledge Base folder** — creates the article on first run, updates it on every subsequent run

Run it on a schedule (weekly, daily, monthly) and anyone with Knowledge Base access in NinjaOne can open it at any time.

---

## What the Report Looks Like

The HTML report renders directly inside the NinjaOne KB viewer and includes:

**Summary bar at the top:**
- Total technicians with tracked time
- Total hours across all technicians
- Average hours per technician
- Total time-log entries processed

**Per-technician cards (expandable):**
- Technician avatar (initials), name, and total hours
- Visual progress bar showing relative time vs. top tech
- Click to expand — reveals a per-organization breakdown table

**Filters and search:**
- **Date range picker** — From / To date inputs to narrow the view
- **Search bar** — Filter by technician name in real time
- **Reset button** — Returns to the default 90-day window

> ℹ️ The date range filter in the report UI reflects what was baked into the data when the script ran. To pull a genuinely different date range, change `$LookbackDays` in the script config and re-run.

---

## Prerequisites

| Requirement | Notes |
|---|---|
| PowerShell 5.1+ | Built-in on Windows 10/11. Run `$PSVersionTable` to check. |
| NinjaOne System Administrator access | Required to create API credentials and KB folders |
| NinjaOne Ticketing module | Must be enabled on your account — this script reads ticketing data |
| No extra modules needed | Uses only built-in PowerShell commands |

---

## Part 1: Create the API App

This is a **one-time setup**. You are creating a set of credentials that let the script talk to NinjaOne silently.

1. Log into NinjaOne as a **System Administrator**
2. Go to: **Administration → Apps → API → Client App IDs → Add**
3. Fill in the form exactly as shown:

   | Field | What to Enter |
   |---|---|
   | **Name** | Any name, e.g. `TimeReportScript` |
   | **Platform** | `API Services (Machine-to-Machine)` |
   | **Allowed Scopes** | Check ✅ `monitoring` AND ✅ `management` |
   | **Redirect URI** | Leave blank — not needed for this flow |

4. Click **Save**
5. You will see a **Client ID** and **Client Secret** — **copy both right now**

> ⚠️ The Client Secret is only shown once. If you miss it, you will need to delete the app and create a new one.

> ⚠️ Never share or commit your Client Secret to source control. Treat it like a password.

---

## Part 2: Create the Knowledge Base Folder

The script needs an existing KB folder to post the report into.

1. In NinjaOne, go to: **Knowledge Base** (in the left sidebar)
2. Click **New Folder**
3. Name it something like `Reports` or `Technician Reports`
4. Click **Save**
5. Open the folder you just created
6. Look at the URL in your browser — it will look like this:

   ```
   https://app.ninjarmm.com/#/knowledgeBase/folder/42
                                                     ^^
                                              This is your Folder ID
   ```

7. Copy that number — you'll need it in the next step

---

## Part 3: Configure the Script

Open `Invoke-NinjaTimeReport.ps1` in any text editor. Find the **CONFIGURATION** block at the top and fill in all six values:

```powershell
# ==============================================================================
#  CONFIGURATION — Fill in ALL values in this block before running the script
# ==============================================================================

$BaseUrl       = 'https://<your Login URL>'
$TokenEndpoint = 'https://<your Login URL>/ws/oauth/token'
$ClientId      = '<Your Client ID>'
$ClientSecret  = '<Your Client Secret>'
$KbFolderId    = 0        # <-- Replace with your KB folder ID number
$LookbackDays  = 90       # How far back to pull ticket data
$KbArticleName = 'Technician Time Report'   # Name of the KB article
```

**Field reference:**

| Variable | Example | Notes |
|---|---|---|
| `$BaseUrl` | `https://app.ninjarmm.com` | Your NinjaOne login URL. See regional table below. |
| `$TokenEndpoint` | `https://app.ninjarmm.com/ws/oauth/token` | Same as `$BaseUrl` plus `/ws/oauth/token` |
| `$ClientId` | `abc123...` | From the API app you created in Part 1 |
| `$ClientSecret` | `s3cr3t...` | From the API app — shown once at creation |
| `$KbFolderId` | `42` | The number from the KB folder URL (Part 2) |
| `$LookbackDays` | `90` | How many days of ticket history to pull. 90 is the default. |
| `$KbArticleName` | `Technician Time Report` | The article title in the KB. You can change this. |

**Regional URLs:**

| Region | Base URL |
|---|---|
| United States | `https://app.ninjarmm.com` |
| Europe | `https://eu.ninjarmm.com` |
| Oceania | `https://oc.ninjarmm.com` |
| Canada | `https://ca.ninjarmm.com` |

---

## Part 4: Run the Script

1. Open **PowerShell** (does not need to be Administrator for a local run)
2. Navigate to where you saved the script:
   ```powershell
   cd C:\Scripts
   ```
3. Run it:
   ```powershell
   .\Invoke-NinjaTimeReport.ps1
   ```
4. Watch the output — it will tell you at each step what it is doing:
   ```
   [1/6] Authenticating...
   [✓] Authenticated.

   [2/6] Fetching technicians...
   [✓] Found 8 technician(s).

   [3/6] Fetching organizations...
   [✓] Found 45 organization(s).

   [4/6] Fetching tickets updated in the last 90 days (paginated)...
   [✓] Found 312 ticket(s) in the date window.

   [5/6] Fetching time-tracking log entries...
       Processing ticket 25 / 312...
       Processing ticket 50 / 312...
   [✓] Time data aggregated across 312 ticket(s).

   [6/6] Building HTML report and posting to Knowledge Base...
   [✓] Knowledge Base article created successfully.

   [✓] REPORT COMPLETE
       Technicians : 8 with tracked time
       Tickets     : 312 processed
       KB Folder   : ID 42
       Article     : Technician Time Report
       Period      : 2026-03-26 to 2026-06-24
   ```

5. Open NinjaOne and go to **Knowledge Base → your folder → Technician Time Report**

> ℹ️ The first run creates the article. Every subsequent run updates it in place. You will never end up with duplicate articles.

---

## Scheduling the Report

To keep the report automatically up to date, schedule the script to run on a regular cadence.

### Windows Task Scheduler (recommended)

1. Open **Task Scheduler** → **Create Basic Task**
2. Set the trigger (e.g. every Monday at 7:00 AM)
3. Set the action to:
   - **Program:** `powershell.exe`
   - **Arguments:** `-NonInteractive -ExecutionPolicy Bypass -File "C:\Scripts\Invoke-NinjaTimeReport.ps1"`
4. Set **Run whether user is logged on or not**
5. Click **OK**

### Running Manually Whenever You Need It

Just run the script in PowerShell. It takes 1–5 minutes depending on how many tickets exist in the window.

---

## How Time Tracking is Calculated

NinjaOne records time against individual ticket log entries. Each time a technician logs time on a ticket, it creates a log entry with:
- `appUserContactId` — which technician logged the time
- `appUserContactType` — `TECHNICIAN` (the script filters to this type only)
- `timeTracked` — seconds of time logged
- `createTime` — when the log entry was created (used for date filtering)
- `clientId` on the parent ticket — which organization the ticket belongs to

The script:
1. Only processes log entries where `appUserContactType == TECHNICIAN`
2. Only counts entries where `timeTracked > 0`
3. Only counts entries where `createTime` falls within the date window
4. Converts seconds to hours (`/ 3600`) and rounds to 2 decimal places
5. Groups by technician, then by organization within each technician

**Tickets with zero time tracked are skipped entirely** to avoid unnecessary API calls.

---

## Do I Need to Change the Grant Type?

**No.** Client Credentials (the grant type this script uses) works for all endpoints this script calls, including the Knowledge Base. Here is why:

| Grant Type | What It Does | Needed Here? |
|---|---|---|
| Authorization Code | Opens a browser, logs in as a named user | ❌ Not needed |
| Client Credentials | Silent, uses App ID + Secret, no browser | ✅ Used by this script |

The Knowledge Base `POST /v2/knowledgebase/articles` endpoint is covered by the `management` scope, which Client Credentials can request. No grant type change required.

---

## API Endpoints Used

| Method | Endpoint | Purpose |
|---|---|---|
| `POST` | `/ws/oauth/token` | Get access token (Client Credentials) |
| `GET` | `/v2/technicians` | List all technician accounts |
| `GET` | `/v2/organizations` | List all organizations (paginated) |
| `GET` | `/v2/ticketing/ticket` | List all tickets (cursor-paginated) |
| `GET` | `/v2/ticketing/ticket/{id}/log-entry` | Get time-tracking log entries per ticket |
| `GET` | `/v2/knowledgebase/global/articles` | Check if the report article already exists |
| `POST` | `/v2/knowledgebase/articles` | Create the KB article (first run) |
| `PATCH` | `/v2/knowledgebase/article/{id}` | Update the KB article (subsequent runs) |

---

## Troubleshooting

| Error / Symptom | Solution |
|---|---|
| `Authentication failed` | Check `$BaseUrl`, `$ClientId`, `$ClientSecret`. Verify the API app is `API Services (Machine-to-Machine)` with `monitoring` and `management` scopes. |
| `Please fill in BaseUrl` | You left the `<your Login URL>` placeholder in the config — replace it with your actual URL. |
| `Please set KbFolderId` | Change `$KbFolderId = 0` to your actual folder ID number from the KB folder URL. |
| `Ticketing must be enabled` | The script got a 403 or 404 on the ticketing endpoint. NinjaOne Ticketing must be an active feature on your account. |
| `Failed to create KB article (403)` | Your API app is missing the `management` scope. Edit the app in Administration → Apps → API and enable it. |
| `Failed to create KB article (404)` | The `$KbFolderId` does not exist. Double-check it from the folder URL in the KB. |
| Report shows 0 technicians | No tickets with time tracking were found in the lookback window. Try increasing `$LookbackDays` or verify that technicians are logging time on tickets. |
| Report is slow to generate | Normal for large environments. Every ticket with tracked time requires a separate API call. 300+ tickets may take 3–5 minutes. |
| HTML renders as raw text in KB | NinjaOne should render HTML in KB article content. If you see raw code, check if your NinjaOne plan supports HTML KB content. |
| Date filter in report doesn't change data | The date filter in the HTML is a view-only UI hint. To pull data for a different period, change `$LookbackDays` and re-run the script. |

---

## Pre-Flight Checklist

- [ ] NinjaOne System Administrator access confirmed
- [ ] NinjaOne Ticketing module enabled on your account
- [ ] API app created with platform `API Services (Machine-to-Machine)`
- [ ] API app has both `monitoring` and `management` scopes checked
- [ ] Client ID and Client Secret saved somewhere safe
- [ ] Knowledge Base folder created and folder ID noted from the URL
- [ ] All six variables filled in the CONFIGURATION block
- [ ] Test run completed — watched for `[✓] REPORT COMPLETE` in output
- [ ] Opened Knowledge Base in NinjaOne and confirmed the article appeared
- [ ] (Optional) Scheduled with Windows Task Scheduler for automatic refresh
