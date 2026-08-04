# Set-NinjaTicketDevice.ps1

> Assigns a device to a NinjaOne ticket via the NinjaOne API. Uses the Authorization Code OAuth flow so the action is performed under your technician account — the same way it would be if you did it manually in the NinjaOne UI.

---

## Table of Contents

- [What This Script Does](#what-this-script-does)
- [Why It Works Differently From Other Scripts](#why-it-works-differently-from-other-scripts)
- [Requirements](#requirements)
- [Part 1: Create the API App in NinjaOne](#part-1-create-the-api-app-in-ninjaone)
- [Part 2: Configure the Script](#part-2-configure-the-script)
- [Part 3: Run the Script](#part-3-run-the-script)
- [How Authentication Works](#how-authentication-works)
- [Finding Your Ticket ID and Device ID](#finding-your-ticket-id-and-device-id)
- [Known Requirement: Ticket Must Have an Organization](#known-requirement-ticket-must-have-an-organization)
- [Diagnostic Mode](#diagnostic-mode)
- [Troubleshooting](#troubleshooting)
- [Pre-Flight Checklist](#pre-flight-checklist)

---

## What This Script Does

1. Authenticates to NinjaOne using your technician account via a browser login
2. Looks up the ticket to confirm it exists and shows you its current details
3. Looks up the device to confirm it exists and shows you its name and node class
4. Assigns the device to the ticket — only the device field changes, everything else on the ticket stays exactly as it is
5. Caches the login token so future runs don't need the browser again

---

## Why It Works Differently From Other Scripts

Most NinjaOne API scripts in this repo use **Client Credentials** — a silent machine-to-machine token that never needs a browser login. That approach works for read operations and most management tasks.

However, NinjaOne explicitly **blocks Client Credentials tokens from modifying tickets**. When a ticket is updated, NinjaOne requires a user context — meaning it needs to know which technician made the change. This is enforced at the API level and cannot be bypassed with different scopes or permissions.

This script uses the **Authorization Code** flow instead. You log in once via a browser window, and NinjaOne issues a token tied to your technician account. That token is cached locally so subsequent runs are silent — the browser only opens again if the cached token expires.

---

## Requirements

| Requirement | Notes |
|---|---|
| PowerShell 5.1+ | Built-in on Windows 10/11. Run `$PSVersionTable` to check. |
| PSAuthClient module | Installed automatically by the script if missing |
| NinjaOne technician account | The account you log in with must have permission to edit tickets |
| NinjaOne System Administrator | Required to create the API app (one-time setup only) |
| Internet access | Required to reach NinjaOne and install PSAuthClient from PSGallery |

---

## Part 1: Create the API App in NinjaOne

> ⚠️ This must be a **separate app** from any existing machine-to-machine app. Client Credentials apps cannot be used for ticket updates.

1. Log into NinjaOne as a **System Administrator**
2. Go to: **Administration → Apps → API → Client App IDs → Add**
3. Fill in the following:

   | Field | Value |
   |---|---|
   | **Name** | Any name, e.g. `TicketDeviceAssign` |
   | **Platform** | `Web` |
   | **Allowed Scopes** | ✅ `monitoring` AND ✅ `management` |
   | **Grant Types** | ✅ `Authorization Code` AND ✅ `Refresh Token` |
   | **Redirect URI** | `https://localhost/` — exactly this, with the trailing slash |

4. Click **Save** — copy the **Client ID** and **Client Secret** immediately

> ⚠️ The Client Secret is shown only once. Save it somewhere secure before closing the page.

> ℹ️ The Redirect URI must be `https://localhost/` exactly — including the `https://` and the trailing `/`. PSAuthClient handles the HTTPS localhost callback automatically without requiring any certificates or admin rights.

---

## Part 2: Configure the Script

Open `Set-NinjaTicketDevice.ps1` in any text editor and fill in the **CONFIGURATION** block at the top:

```powershell
# ==============================================================================
#  CONFIGURATION -- Fill in ALL values before running
# ==============================================================================

$BaseUrl       = 'https://<your Login URL>'       # No trailing slash
$ClientId      = '<Your Web App Client ID>'
$ClientSecret  = '<Your Web App Client Secret>'

# The ticket you want to assign a device to
$TicketId      = 0    # <-- Replace with your ticket ID

# The device you want to assign to the ticket
$DeviceId      = 0    # <-- Replace with your device ID

# Set to $true to print the raw ticket JSON and exit without making changes
$DiagnosticMode = $false
```

| Variable | Example | Notes |
|---|---|---|
| `$BaseUrl` | `https://app.ninjarmm.com` | Your NinjaOne login URL — no trailing slash |
| `$ClientId` | `abc123...` | From Part 1 |
| `$ClientSecret` | `s3cr3t...` | From Part 1 — shown once |
| `$TicketId` | `8536` | The numeric ID of the ticket — see below for how to find it |
| `$DeviceId` | `3` | The numeric ID of the device — see below for how to find it |
| `$DiagnosticMode` | `$false` | Set to `$true` to inspect the ticket without making changes |

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
.\Set-NinjaTicketDevice.ps1
```

**First run** — a browser window opens automatically:

```
  ================================================================
  NinjaOne -- Assign Device to Ticket
  Ticket ID: 8536  |  Device ID: 3
  ================================================================

  [1/4] Authenticating...
       Found cached token -- attempting silent refresh...

       (or on first run:)

  [..] Opening browser for NinjaOne login...
       Log in and approve the request in your browser.
       This window will continue automatically after login.
```

After the browser login completes the script continues automatically:

```
  [OK] Login successful. Tokens cached for future runs.

  [2/4] Looking up ticket 8536...
  [OK] Found ticket:
       Subject : Outlook crashes repeatedly
       Status  : OPEN
       Org ID  : 4
       Current node   : (none)

  [3/4] Looking up device 3...
  [OK] Found device:
       Name       : SE-W19E
       Node class : WINDOWS_SERVER
       Org ID     : 4

  [4/4] Assigning device 3 (SE-W19E) to ticket 8536...
  [OK] Device assigned successfully.
       Ticket ID      : 8536
       Ticket subject : Outlook crashes repeatedly
       Node ID        : 3
       Device name    : SE-W19E

  ================================================================
  [OK] COMPLETE
  ================================================================
```

**Subsequent runs** — no browser needed, uses the cached token silently.

---

## How Authentication Works

The script uses the **Authorization Code + Refresh Token** flow:

| Run | What happens |
|---|---|
| **First run** | Browser opens → you log in → NinjaOne issues an access token and refresh token → both are saved to `NinjaTokenCache.json` in the script folder |
| **Subsequent runs** | Script reads `NinjaTokenCache.json` and uses the refresh token to get a new access token silently — no browser |
| **Token expired** | If the refresh token has expired, the browser opens again for a fresh login |
| **Force fresh login** | Delete `NinjaTokenCache.json` from the script folder and run again |

> ⚠️ `NinjaTokenCache.json` contains a refresh token that grants access to your NinjaOne account. Keep this file secure and do not share it or commit it to source control.

---

## Finding Your Ticket ID and Device ID

### Ticket ID

Open the ticket in NinjaOne. The ticket ID is the number at the end of the URL:

```
https://app.ninjarmm.com/#/ticketing/ticket/8536
                                              ^^^^
```

Or it is displayed as the ticket number in the ticket list view.

### Device ID

The easiest way is to open the device in NinjaOne and read the number from the URL:

```
https://app.ninjarmm.com/#/deviceDashboard/3/overview
                                           ^
```

Alternatively you can use the `Get-NinjaUptimeReport.ps1` script which lists all devices with their IDs in the console output, or query the API directly:

```powershell
# List all devices and their IDs (requires a valid Client Credentials token)
Invoke-RestMethod -Uri 'https://app.ninjarmm.com/v2/devices?pageSize=50' `
    -Headers @{ Authorization = "Bearer YOUR_TOKEN" } |
    Select-Object id, systemName, nodeClass | Format-Table
```

---

## Known Requirement: Ticket Must Have an Organization

NinjaOne requires a ticket to be linked to an organization before a device can be assigned to it. If the ticket has no organization set the script will exit with a clear message:

```
  [!] Ticket 8537 has no organization assigned.
      A device cannot be assigned to a ticket that is not linked to an organization.
      Open the ticket in NinjaOne, assign it to an organization, then run this script again.
```

**To fix:** open the ticket in NinjaOne, set the **Organization** field, save the ticket, then run the script again.

---

## Diagnostic Mode

Set `$DiagnosticMode = $true` in the config block to inspect a ticket without making any changes. The script will authenticate, fetch the ticket, print the full raw JSON, and exit before the update step.

```powershell
$DiagnosticMode = $true
```

Example output:

```
  [DIAGNOSTIC] Raw ticket JSON:
{
    "id":  8536,
    "version":  3,
    "subject":  "Outlook crashes repeatedly",
    "status":  {
                   "name":  "OPEN",
                   ...
               },
    "clientId":  4,
    "nodeId":  null,
    ...
}

  [DIAGNOSTIC] Exiting without making changes.
               Set $DiagnosticMode = $false to run normally.
```

Set `$DiagnosticMode = $false` to run normally after reviewing the output.

---

## Troubleshooting

| Error / Symptom | Solution |
|---|---|
| `Fill in $BaseUrl` | Placeholder still in config. Replace all `<...>` values. |
| `Invalid redirect_uri` | The Redirect URI in your NinjaOne app must be exactly `https://localhost/` including the trailing slash. |
| `user_context_required` | You are using a Client Credentials app. This script requires a **Web** platform app with Authorization Code grant type. |
| `Ticket X not found` | The ticket ID doesn't exist or your account doesn't have access to it. |
| `Device X not found` | The device ID doesn't exist or your account doesn't have access to it. |
| `Ticket has no organization assigned` | Open the ticket in NinjaOne, assign it to an organization, then re-run. |
| `ticket_updated_by_another_user` | Someone else modified the ticket between when you ran the script and when it tried to update. Run the script again immediately. |
| `PSAuthClient failed to install` | Run manually: `Install-Module PSAuthClient -Scope CurrentUser -Force` — may require internet access or an unrestricted execution policy. |
| Browser opens but redirects to an error page | Check the Redirect URI in your NinjaOne app matches `https://localhost/` exactly. |
| Token cache keeps expiring | NinjaOne refresh tokens have a limited lifetime. This is normal — log in again when prompted. |
| Script assigns the wrong device | Double-check `$DeviceId` in the config. Use the device URL in NinjaOne to confirm the ID. |

---

## Pre-Flight Checklist

- [ ] PowerShell 5.1+ confirmed (`$PSVersionTable`)
- [ ] NinjaOne System Administrator access confirmed (for app creation)
- [ ] NinjaOne technician account confirmed (for browser login — must have ticket edit permissions)
- [ ] API app created — Platform: `Web`, Grant Types: `Authorization Code` + `Refresh Token`, Redirect URI: `https://localhost/`
- [ ] Client ID and Client Secret saved securely
- [ ] All config variables filled in — no `<placeholder>` text remaining
- [ ] `$TicketId` set to a real ticket ID
- [ ] `$DeviceId` set to a real device ID
- [ ] Target ticket confirmed to have an organization assigned in NinjaOne
- [ ] Script run — browser login completed successfully
- [ ] Console output shows `[OK] COMPLETE`
- [ ] Ticket confirmed in NinjaOne — device now shows in the ticket details
- [ ] `NinjaTokenCache.json` stored securely — not shared or committed to source control
