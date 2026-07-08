# Get-NinjaUnassignedTicketReport.ps1

> Generates an interactive HTML report showing every ticket that sat unassigned longer than a configurable threshold — including who it was eventually assigned to, who made the assignment, and how long it took. Also surfaces any tickets that are still unassigned right now.

---

## Table of Contents

- [What This Script Does](#what-this-script-does)
- [What the Report Shows](#what-the-report-shows)
- [Which Tickets Are Included](#which-tickets-are-included)
- [Prerequisites](#prerequisites)
- [Part 1: Create the API App](#part-1-create-the-api-app)
- [Part 2: Configure the Script](#part-2-configure-the-script)
- [Part 3: Run the Script](#part-3-run-the-script)
- [Reading the Report](#reading-the-report)
- [How Wait Time Is Calculated](#how-wait-time-is-calculated)
- [How "Assigned By" Is Determined](#how-assigned-by-is-determined)
- [Changing the Threshold or Lookback Window](#changing-the-threshold-or-lookback-window)
- [API Endpoints Used](#api-endpoints-used)
- [Troubleshooting](#troubleshooting)
- [Pre-Flight Checklist](#pre-flight-checklist)

---

## What This Script Does

For every ticket created within the lookback window (default: last 30 days), the script:

1. Records the ticket creation timestamp
2. Pulls all log entries for that ticket
3. Finds the **first assignment event** in the log
4. Calculates the gap between ticket creation and first assignment
5. Flags the ticket if that gap exceeds `$ThresholdMinutes` (default: 15 minutes)
6. Also flags tickets with **no assignment at all** — still sitting unassigned right now
7. Captures **who the ticket was assigned to** and **who performed the assignment**
8. Builds a self-contained interactive HTML report and opens it automatically

---

## What the Report Shows

**Summary stat cards (update live as you filter):**
- Total tickets shown
- Still unassigned — tickets with no assignee right now (shown in red)
- Eventually assigned — tickets that were assigned after the threshold
- Average wait time across all shown tickets
- Longest single wait time

**Sortable, filterable table with these columns:**

| Column | Description |
|---|---|
| **Ticket #** | NinjaOne ticket ID |
| **Subject** | Ticket subject line |
| **Organization** | The client org the ticket belongs to |
| **Priority** | CRITICAL / HIGH / MEDIUM / LOW / NONE — colour coded |
| **Status** | Current ticket status (OPEN, IN_PROGRESS, RESOLVED, etc.) |
| **Created** | Date and time the ticket was created |
| **Wait Time** | How long the ticket sat unassigned before first assignment |
| **Assigned To** | The technician the ticket was assigned to |
| **Assigned By** | The person who performed the assignment action |

**Wait time badge colour coding:**

| Badge colour | Meaning |
|---|---|
| 🟡 Yellow | Over threshold (e.g. > 15 min) |
| 🟠 Orange | Over 2× threshold (e.g. > 30 min) |
| 🔴 Red | Over 4× threshold (e.g. > 60 min) |
| 🔴 Red italic | Still unassigned — no assignee at all |

**Filter controls:**
- Free-text search by subject or organization name
- Organization dropdown
- Priority dropdown
- Status filter: All / Still Unassigned Only / Eventually Assigned Only

All columns are sortable by clicking the header. Default sort is worst wait time first.

---

## Which Tickets Are Included

The script pulls **all tickets** from the lookback window regardless of their current status — open, in progress, resolved, and closed tickets are all included.

This is intentional. A ticket that was resolved yesterday but sat unassigned for two hours is still evidence of a slow response and should appear in the report. Limiting to only open tickets would miss the historical pattern entirely.

The only tickets excluded are those whose wait time did not exceed the threshold and that currently have an assignee — in other words, tickets that were assigned promptly are not shown.

---

## Prerequisites

| Requirement | Notes |
|---|---|
| PowerShell 5.1+ | Built-in on Windows 10/11. Run `$PSVersionTable` to check. |
| NinjaOne System Administrator | Required to create the API app |
| NinjaOne Ticketing | Must be enabled on your account |
| No extra modules | Uses only built-in PowerShell |

---

## Part 1: Create the API App

This is a **one-time setup**. You are creating silent read-only credentials the script uses to authenticate.

1. Log into NinjaOne as a **System Administrator**
2. Go to: **Administration → Apps → API → Client App IDs → Add**
3. Fill in:

   | Field | Value |
   |---|---|
   | **Name** | Any name, e.g. `UnassignedTicketReport` |
   | **Platform** | `API Services (Machine-to-Machine)` |
   | **Allowed Scopes** | ✅ `monitoring` — read-only is sufficient |
   | **Redirect URI** | Leave blank |

4. Click **Save** — copy the **Client ID** and **Client Secret** immediately

> ⚠️ The Client Secret is shown only once. If you miss it, delete the app and create a new one.

> ℹ️ Only the `monitoring` scope is needed. This script never creates, updates, or deletes anything in NinjaOne.

---

## Part 2: Configure the Script

Open `Get-NinjaUnassignedTicketReport.ps1` in any text editor. Find the **CONFIGURATION** block at the top and fill in all values:

```powershell
# ==============================================================================
#  CONFIGURATION -- Fill in ALL values before running
# ==============================================================================

$BaseUrl          = 'https://<your Login URL>'
$TokenEndpoint    = 'https://<your Login URL>/ws/oauth/token'
$ClientId         = '<Your Client ID>'
$ClientSecret     = '<Your Client Secret>'
$ThresholdMinutes = 15
$LookbackDays     = 30
```

| Variable | Example | Notes |
|---|---|---|
| `$BaseUrl` | `https://app.ninjarmm.com` | Your NinjaOne login URL — no trailing slash |
| `$TokenEndpoint` | `https://app.ninjarmm.com/ws/oauth/token` | Same as `$BaseUrl` + `/ws/oauth/token` |
| `$ClientId` | `abc123...` | From the API app created in Part 1 |
| `$ClientSecret` | `s3cr3t...` | From the API app — shown once at creation |
| `$ThresholdMinutes` | `15` | Tickets assigned faster than this are excluded from the report |
| `$LookbackDays` | `30` | How many days of ticket history to analyse |

**Regional URLs:**

| Region | Base URL |
|---|---|
| United States | `https://app.ninjarmm.com` |
| Europe | `https://eu.ninjarmm.com` |
| Oceania | `https://oc.ninjarmm.com` |
| Canada | `https://ca.ninjarmm.com` |

> ⚠️ Do not commit `$ClientSecret` to source control. Treat it like a password.

---

## Part 3: Run the Script

1. Open **PowerShell** (does not need to be Administrator)
2. Navigate to the folder where the script is saved:
   ```powershell
   cd C:\Scripts
   ```
3. Run it:
   ```powershell
   .\Get-NinjaUnassignedTicketReport.ps1
   ```
4. Watch the progress output:

   ```
   ================================================================
   NinjaOne Unassigned Ticket Report  [Threshold: 15min]
   Lookback: 30 days
   ================================================================

   [1/5] Authenticating...
   [OK] Authenticated.

   [2/5] Loading organizations and technicians...
   [OK] 45 org(s), 12 technician(s).

   [3/5] Fetching tickets from the last 30 days...
   [OK] 312 ticket(s) in the last 30 days.

   [4/5] Analysing assignment timing per ticket...
   (progress bar)

   [5/5] Building HTML report...

   ================================================================
   [OK] REPORT COMPLETE
        Tickets analysed  : 312
        Breached threshold: 47
        Still unassigned  : 3
        Avg wait time     : 38.2 min
        Longest wait      : 247.5 min -- Outlook won't open
        Report saved to   : C:\Scripts\NinjaUnassignedTicketReport_20260701_143022.html
   ================================================================
   ```

5. The HTML report opens automatically in your default browser
6. The report file is saved to the same folder as the script with a timestamp in the filename

> ℹ️ The script can take several minutes on large environments. Each ticket that breached the threshold requires a separate API call to fetch its log entries. A progress bar shows which ticket is being processed.

---

## Reading the Report

### Wait time badges

The colour of the wait time badge tells you at a glance how severe the breach was:

| Badge | Threshold Example (15 min default) |
|---|---|
| 🟡 Yellow — `23 min` | Between 15 and 30 minutes |
| 🟠 Orange — `67 min` | Between 30 and 60 minutes |
| 🔴 Red — `180 min` | Over 60 minutes |
| 🔴 Red italic — `Still Unassigned` | No assignee at all — gap is measured from creation to right now |

### Assigned To vs Assigned By

| Column | Who it represents |
|---|---|
| **Assigned To** | The technician the ticket was routed to — the one who received the ticket |
| **Assigned By** | The person who performed the assignment — the one who clicked assign |

These can be the same person (a technician assigning a ticket to themselves) or different people (a manager or dispatcher assigning work to a team member).

For **still unassigned** tickets, both columns show `—` since no assignment has occurred.

### Priority colour coding

| Colour | Priority |
|---|---|
| 🔴 Bold red | CRITICAL |
| 🟠 Orange | HIGH |
| 🟡 Yellow | MEDIUM |
| 🟢 Green | LOW |
| Gray | NONE |

### Sorting and filtering

- **Click any column header** to sort by that column. Click again to reverse the sort order. Default is worst wait time first.
- **Search box** — filters by subject or organization name as you type
- **Organization dropdown** — narrows to a single client
- **Priority dropdown** — narrows to a single priority level
- **Status dropdown** — switch between All, Still Unassigned Only, or Eventually Assigned Only

All stat cards update automatically to reflect the current filtered view.

---

## How Wait Time Is Calculated

**For tickets that were eventually assigned:**

```
Wait Time = Timestamp of first assignment log entry − Ticket creation timestamp
```

The script looks through the ticket's log entries for the first event with type `ASSIGNMENT`, `TECHNICIAN_CHANGED`, or `ASSIGNED`. If none of those explicit types exist, it falls back to activity entries whose body text contains the word "assign".

**For tickets still unassigned:**

```
Wait Time = Current time − Ticket creation timestamp
```

This value grows continuously as the ticket ages without being assigned.

**Tickets assigned within the threshold** (e.g. within 15 minutes) are not shown in the report at all — they passed the SLA and are not a concern.

---

## How "Assigned By" Is Determined

The script reads the assignment actor from the log entry using three fallback layers:

1. **`createdBy.name`** — the primary source. NinjaOne embeds the actor's name as a nested object `{ id, name }` on the log entry. This is used first when available.

2. **`actorUserId`** — a flat ID field used by some NinjaOne API versions. Cross-referenced against the technician list loaded at startup.

3. **`userId`** — a secondary flat ID field used by older API response formats.

If none of these fields are present on the log entry, the column shows `Unknown`.

---

## Changing the Threshold or Lookback Window

Both values are at the top of the script in the CONFIGURATION block:

```powershell
# How many minutes before a ticket is considered "slow to assign"
$ThresholdMinutes = 15

# How many days back to look for tickets
$LookbackDays = 30
```

**Common configurations:**

| Use case | `$ThresholdMinutes` | `$LookbackDays` |
|---|---|---|
| Strict SLA (15 min) | `15` | `30` |
| Standard SLA (1 hour) | `60` | `30` |
| Weekly review | `30` | `7` |
| Monthly management report | `30` | `90` |
| Identify the worst offenders | `60` | `90` |

> ℹ️ Increasing `$LookbackDays` significantly increases run time, as more tickets means more log entry API calls. For 90 days on a busy instance, expect 5–15 minutes of run time.

---

## API Endpoints Used

| Method | Endpoint | Purpose |
|---|---|---|
| `POST` | `/ws/oauth/token` | Get access token (Client Credentials) |
| `GET` | `/v2/organizations` | Load org names for display (paginated) |
| `GET` | `/v2/technicians` | Load technician names for display |
| `GET` | `/v2/ticketing/ticket` | Fetch all tickets in the lookback window (cursor-paginated) |
| `GET` | `/v2/ticketing/ticket/{id}/log-entry` | Fetch log entries to find assignment events (one call per breached ticket) |

---

## Troubleshooting

| Error / Symptom | Solution |
|---|---|
| `Fill in $BaseUrl` | The `<your Login URL>` placeholder is still in the config block. |
| `Authentication failed` | Check `$BaseUrl`, `$ClientId`, `$ClientSecret`. API app must be `API Services (Machine-to-Machine)` with `monitoring` scope. |
| Report shows 0 tickets | Either no tickets were created in the lookback window, or all tickets were assigned within the threshold. Try increasing `$LookbackDays` or lowering `$ThresholdMinutes`. |
| `Assigned By` shows `Unknown` for all rows | The log entry format on your NinjaOne instance may not include a `createdBy` field on assignment events. This varies by NinjaOne version. The `Assigned To` column is unaffected. |
| `Assigned To` shows `Unknown` | The technician ID on the log entry doesn't match any entry in the technician list. This can happen if the technician account was deleted after the assignment was made. |
| Report is slow to generate | Expected — each ticket that breached the threshold requires its own API call for log entries. Reduce `$LookbackDays` or increase `$ThresholdMinutes` to process fewer tickets. |
| Report opens but shows `No tickets match your filters` | A filter is active. Click **Reset** to clear all filters. |
| Still unassigned count seems wrong | The script checks the current assignee on the ticket object at the time the script runs. If a ticket was assigned after the script started, it may appear as still unassigned. Re-run the script for a fresh snapshot. |
| Ticketing not enabled error | NinjaOne Ticketing must be an active feature on your account. Contact NinjaOne support if the ticketing endpoint returns a 403 or 404. |

---

## Pre-Flight Checklist

- [ ] NinjaOne Ticketing is enabled on your account
- [ ] NinjaOne System Administrator access confirmed
- [ ] API app created — Platform: `API Services (Machine-to-Machine)`, Scope: `monitoring`
- [ ] Client ID and Client Secret saved securely
- [ ] `$BaseUrl`, `$TokenEndpoint`, `$ClientId`, `$ClientSecret` filled in
- [ ] `$ThresholdMinutes` set to match your SLA requirement
- [ ] `$LookbackDays` set to the desired reporting window
- [ ] Script run — confirmed `[OK] REPORT COMPLETE` in the output
- [ ] HTML report opened and reviewed in browser
