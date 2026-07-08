# Get-NinjaUptimeReport.ps1

> Generates an interactive HTML uptime report for servers and network devices with per-device uptime bars, org grouping, fleet averages, and instant 30/60/90-day window switching. Saves a timestamped HTML file locally and opens it automatically on completion.

> **Note:** Knowledge Base posting will be added in a future revision. This version generates and saves the report locally only.

---

## Table of Contents

- [What This Script Does](#what-this-script-does)
- [What the Report Shows](#what-the-report-shows)
- [Device Scope](#device-scope)
- [How Uptime Is Calculated](#how-uptime-is-calculated)
- [Offline Device Handling](#offline-device-handling)
- [Important Accuracy Note](#important-accuracy-note)
- [The 30 / 60 / 90 Day Dropdown](#the-30--60--90-day-dropdown)
- [Prerequisites](#prerequisites)
- [Part 1: Create the API App](#part-1-create-the-api-app)
- [Part 2: Configure the Script](#part-2-configure-the-script)
- [Part 3: Run the Script](#part-3-run-the-script)
- [Reading the Report](#reading-the-report)
- [Scheduling the Script](#scheduling-the-script)
- [API Endpoints Used](#api-endpoints-used)
- [Troubleshooting](#troubleshooting)
- [Pre-Flight Checklist](#pre-flight-checklist)

---

## What This Script Does

1. **Authenticates silently** using Client Credentials — no browser, no login prompt
2. **Pulls all servers and network devices** from your NinjaOne instance
3. **Pulls 90 days of activity logs** per device, looking for offline and online status events
4. **Calculates uptime percentage** per device for 30, 60, and 90-day windows
5. **Handles offline devices with no logged events** — if a device is currently offline but has no activity log, downtime is estimated from its last contact time to now rather than assuming 100%
6. **Pre-calculates all three datasets** and embeds them in one self-contained HTML file
7. **Saves the report** as a timestamped HTML file in the script directory and opens it automatically

---

## What the Report Shows

**Fleet summary stat cards** (update live as you filter and switch windows):
- Total devices in view
- Fleet average uptime %
- Currently online count
- Currently offline count
- Devices with no activity data
- Lowest uptime device

**Org-grouped device cards**, each showing:
- Organization name, device count, and org average uptime %
- Per-device row with:
  - Device name and node class
  - Number of outages recorded in the window
  - Visual uptime progress bar (green / yellow / red)
  - Uptime percentage
  - Total downtime in minutes
  - Status badge (Online / Offline / Offline* / No Data)
  - Last seen timestamp

**Filters:**
- 30 / 60 / 90 day window dropdown
- Text search by device name or org name
- Organization dropdown
- Device type (Server / Network)
- Status (All / Online Only / Offline Only)
- Reset button

---

## Device Scope

| Category | Node Classes Included |
|---|---|
| **Servers** | `WINDOWS_SERVER`, `LINUX_SERVER`, `MAC_SERVER` |
| **Network** | `NMS_ROUTER`, `NMS_SWITCH`, `NMS_FIREWALL`, `NMS_OTHER`, `NMS_UNKNOWN`, `NMS_PRINTER`, `NMS_STORAGE` |

Workstations, laptops, and mobile devices are intentionally excluded.

The script prints all node classes found in your instance at runtime alongside the target classes it filters for, so any mismatch is immediately visible in the console output.

---

## How Uptime Is Calculated

For each device the script:

1. Pulls activity log entries from the last 90 days via `GET /v2/device/{id}/activities`
2. Filters to events that indicate the device went **offline** or came back **online**
3. Pairs offline events with the next online event to form outage periods
4. Clips each outage period to the start of the reporting window
5. Sums clipped downtime in seconds per window
6. Derives uptime %:

```
Uptime % = (window seconds - downtime seconds) / window seconds × 100
```

**Unpaired offline events** (no matching online event) are handled as follows:
- If the device is **currently offline** — outage assumed ongoing to right now
- If the device is **currently online** — outage closed at the device's `lastContact` timestamp

---

## Offline Device Handling

A key design decision in this script is how it handles devices that are currently offline but have no logged activity events.

**Previous behaviour:** assume 100% uptime (incorrect — makes a genuinely down device look healthy)

**Current behaviour:** four distinct cases

| Situation | How it's handled | Badge |
|---|---|---|
| Has events, currently online | Calculate from paired offline/online events | 🟢 Online |
| Has events, currently offline | Calculate from events + extend last open outage to now | 🔴 Offline |
| No events, currently **online** | Assume 100% — device has been up with nothing to log | ⚪ No Data |
| No events, currently **offline** | Synthesize outage from `lastContact` → now | 🔴 Offline* |

The **Offline*** badge distinguishes a synthesized estimate from a measured one. A note on the device row reads "No events logged; downtime estimated from last contact" so viewers know it's an approximation.

In the console, a yellow `[i]` line is printed for each device that triggers synthesis so you can see exactly which ones were affected.

---

## Important Accuracy Note

> ⚠️ Uptime figures are calculated from NinjaOne agent and NMS check-in events — **not true network-level packet availability.**

- If a device goes down for 3 minutes between polls, NinjaOne may not log it
- The polling interval in your NinjaOne policies determines detection resolution — a 5-minute poll means outages shorter than 5 minutes may be missed
- These figures are best described as "NinjaOne-detected availability" rather than true uptime

For true packet-level availability a dedicated network monitoring solution is required. These figures are still useful for trend analysis and identifying devices with frequent or extended outages.

---

## The 30 / 60 / 90 Day Dropdown

The script pulls 90 days of data in one run and pre-calculates all three windows. Switching the dropdown in the report is instant — no re-fetch, no reload. The report is fully self-contained and works offline once saved. All stat cards, org groups, device bars, and percentages update immediately when the window changes.

---

## Prerequisites

| Requirement | Notes |
|---|---|
| PowerShell 5.1+ | Built-in on Windows 10/11. Run `$PSVersionTable` to check. |
| NinjaOne System Administrator | Required to create the API app |
| No extra modules | Uses only built-in PowerShell |

---

## Part 1: Create the API App

One-time setup.

1. Log into NinjaOne as a **System Administrator**
2. Go to: **Administration → Apps → API → Client App IDs → Add**
3. Fill in:

   | Field | Value |
   |---|---|
   | **Name** | Any name, e.g. `UptimeReport` |
   | **Platform** | `API Services (Machine-to-Machine)` |
   | **Allowed Scopes** | ✅ `monitoring` |
   | **Redirect URI** | Leave blank |

4. Click **Save** — copy the **Client ID** and **Client Secret** immediately

> ⚠️ The Client Secret is shown only once.

> ℹ️ Only `monitoring` scope is required for this version of the script. The `management` scope will be needed when KB posting is added in a future revision.

---

## Part 2: Configure the Script

Open `Get-NinjaUptimeReport.ps1` and fill in the **CONFIGURATION** block at the top:

```powershell
# ==============================================================================
#  CONFIGURATION -- Fill in ALL values before running
# ==============================================================================

$BaseUrl                 = 'https://<your Login URL>'
$TokenEndpoint           = 'https://<your Login URL>/ws/oauth/token'
$ClientId                = '<Your Client ID>'
$ClientSecret            = '<Your Client Secret>'
$UptimeWarnThreshold     = 99.0
$UptimeCriticalThreshold = 95.0
```

| Variable | Example | Notes |
|---|---|---|
| `$BaseUrl` | `https://app.ninjarmm.com` | NinjaOne login URL — no trailing slash |
| `$TokenEndpoint` | `https://app.ninjarmm.com/ws/oauth/token` | Same as `$BaseUrl` + `/ws/oauth/token` |
| `$ClientId` | `abc123...` | From Part 1 |
| `$ClientSecret` | `s3cr3t...` | From Part 1 — shown once |
| `$UptimeWarnThreshold` | `99.0` | Below this % the bar and percentage show yellow |
| `$UptimeCriticalThreshold` | `95.0` | Below this % the bar and percentage show red |

**Regional URLs:**

| Region | Base URL |
|---|---|
| United States | `https://app.ninjarmm.com` |
| Europe | `https://eu.ninjarmm.com` |
| Oceania | `https://oc.ninjarmm.com` |
| Canada | `https://ca.ninjarmm.com` |

---

## Part 3: Run the Script

```powershell
.\Get-NinjaUptimeReport.ps1
```

Example console output:

```
  ================================================================
  NinjaOne Uptime Report -- Servers and Network Devices
  Warn: <99%  Critical: <95%
  ================================================================

  [1/5] Authenticating...
  [OK] Authenticated.

  [2/5] Loading organizations and devices...
  [OK] 45 org(s)  |  412 total devices  |  87 matched
       Node classes in your instance : LINUX_SERVER, NMS_OTHER, NMS_PRINTER, ...
       Node classes this script uses : WINDOWS_SERVER, LINUX_SERVER, MAC_SERVER, ...

  [3/5] Pulling activity logs and calculating uptime...
       (87 devices -- one API call each)
    [i] Juniper 4500 -- offline with no events, outage assumed from lastContact (2026-06-01 12:15)
    [i] RYANSNYDERBE36 -- offline with no events, outage assumed from lastContact (2025-11-14 13:36)
  [OK] Uptime calculated for 87 device(s).

  [4/5] Building HTML report...

  [5/5] Saving report...
  [OK] Local copy saved: C:\Scripts\NinjaUptimeReport_20260708_143022.html

  ================================================================
  [OK] COMPLETE
       Devices processed : 87
       Fleet avg uptime  : 99.71% (90-day)
       Local file        : C:\Scripts\NinjaUptimeReport_20260708_143022.html
  ================================================================
```

The report opens automatically in your default browser.

---

## Reading the Report

### Window dropdown

The **Last 30 Days / Last 60 Days / Last 90 Days** dropdown switches the entire dataset instantly. Stat cards, org groups, device bars, and percentages all update. The blue badge in the header also updates.

### Uptime bar colours

| Colour | Meaning (default thresholds) |
|---|---|
| 🟢 Green bar | At or above `$UptimeWarnThreshold` (≥ 99%) |
| 🟡 Yellow bar | Between `$UptimeCriticalThreshold` and `$UptimeWarnThreshold` (95–99%) |
| 🔴 Red bar | Below `$UptimeCriticalThreshold` (< 95%) |

### Device status badges

| Badge | Meaning |
|---|---|
| 🟢 Online | Device is currently online; uptime calculated from logged events |
| 🔴 Offline | Device is currently offline; uptime calculated from logged events |
| 🔴 Offline* | Device is currently offline with no logged events; downtime estimated from last contact |
| ⚪ No Data | Device is currently online but had no logged events; 100% assumed |

### Org averages

Each org section header shows the average uptime across all devices in that org for the current window, colour-coded to the same green/yellow/red thresholds.

---

## Scheduling the Script

To generate the report automatically on a schedule:

1. Open **Task Scheduler → Create Basic Task**
2. Set your trigger — e.g. every Monday at 6:00 AM
3. Action:
   - **Program:** `powershell.exe`
   - **Arguments:** `-NonInteractive -ExecutionPolicy Bypass -File "C:\Scripts\Get-NinjaUptimeReport.ps1"`
4. Enable **Run whether user is logged on or not**
5. Click **OK**

Each run saves a new timestamped HTML file. If you want only one file to be kept, add a cleanup step or set a fixed output filename by editing the `$LocalPath` line in the script.

---

## API Endpoints Used

| Method | Endpoint | Purpose |
|---|---|---|
| `POST` | `/ws/oauth/token` | Client Credentials auth |
| `GET` | `/v2/organizations` | Load org names (paginated) |
| `GET` | `/v2/devices` | Load all devices (paginated) |
| `GET` | `/v2/device/{id}/activities` | Pull 90-day activity log per device |

---

## Troubleshooting

| Error / Symptom | Solution |
|---|---|
| `Fill in $BaseUrl` | Placeholder still in config. Replace all `<...>` values. |
| `Authentication failed` | Check `$BaseUrl`, `$ClientId`, `$ClientSecret`. App must be `API Services (Machine-to-Machine)` with `monitoring` scope. |
| `0 matched` devices | Node classes in your instance don't match the target list. Check the console output — it prints all node classes found. Add any missing ones to `$AllTargetClasses` in the script. |
| All devices show `No Data` | Activity logs may not be enabled or status events may not be generated by your policy configuration. Devices that are online will still show as 100%. |
| Some devices show `Offline*` | Those devices are currently offline and had no logged events. Downtime is estimated from `lastContact`. This is intentional — it prevents falsely healthy uptime scores. |
| Fleet average seems too high | Devices with `No Data` that are online are assumed 100% and included in the average. This is accurate — if they're online and have no outage events, they were up. |
| `Offline*` device shows very high downtime | The device has been offline for a long time relative to the window. `lastContact` may be outside the 90-day window, in which case the full window is counted as downtime. |
| Script is slow | One activity log API call per device. 100+ devices may take several minutes. Run off-hours if needed. |
| Report opens but bars are blank / 0% | Check that `$UptimeWarnThreshold` and `$UptimeCriticalThreshold` are decimals, not integers (use `99.0` not `99`). |
| HTML file not found after run | Check `$PSScriptRoot` — in ISE this may be empty. The script falls back to `$env:TEMP`. Check the console for the exact saved path. |

---

## Pre-Flight Checklist

- [ ] PowerShell 5.1+ confirmed (`$PSVersionTable`)
- [ ] NinjaOne System Administrator access confirmed
- [ ] API app created — Platform: `API Services (Machine-to-Machine)`, Scope: `monitoring`
- [ ] Client ID and Client Secret saved securely
- [ ] All config variables filled in — no `<placeholder>` text remaining
- [ ] `$UptimeWarnThreshold` and `$UptimeCriticalThreshold` set to match your SLA
- [ ] Script run — confirmed `[OK] COMPLETE` in console output
- [ ] Console output checked — node classes match, device count looks right
- [ ] Report opened — confirmed 30/60/90 day dropdown switches correctly
- [ ] `Offline*` devices reviewed — confirm last contact timestamps are plausible
- [ ] (Optional) Scheduled with Windows Task Scheduler for recurring runs
