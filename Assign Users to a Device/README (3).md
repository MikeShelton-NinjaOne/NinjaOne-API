# Set-NinjaDeviceOwner.ps1

> Assign a NinjaOne device owner by hostname, first name, and last name using the NinjaOne Public API v2.

---

## Table of Contents

- [Overview](#overview)
- [Prerequisites](#prerequisites)
- [Setup](#setup)
  - [Install PSAuthClient](#1-install-psauthclient)
  - [Create a NinjaOne API App](#2-create-a-ninjaone-api-app)
  - [Configure the Script](#3-configure-the-script)
- [Usage](#usage)
  - [Parameters](#parameters)
  - [Examples](#examples)
  - [What Happens When You Run It](#what-happens-when-you-run-it)
- [Bulk Assignment via CSV](#bulk-assignment-via-csv)
- [How It Works](#how-it-works)
  - [API Endpoints Used](#api-endpoints-used)
  - [Hostname Resolution](#hostname-resolution)
- [Troubleshooting](#troubleshooting)
- [Security Notes](#security-notes)
- [Pre-Flight Checklist](#pre-flight-checklist)

---

## Overview

`Set-NinjaDeviceOwner.ps1` automates assigning a device owner in NinjaOne via the Public API v2. Instead of manually navigating the portal for each device, you run a single command with a hostname and a name — the script handles OAuth2 authentication, device and contact lookup, ID resolution, and owner assignment automatically.

---

## Prerequisites

| Requirement | Version | Notes |
|---|---|---|
| PowerShell | 5.1+ | Built-in on Windows 10/11. Run `$PSVersionTable` to check. |
| PSAuthClient | Any | Handles the OAuth2 browser flow. One-time install (see below). |
| NinjaOne Account | Any | Must be a **System Administrator** to create API credentials. |
| Network Access | HTTPS | Outbound HTTPS to your NinjaOne URL must be allowed. |

---

## Setup

### 1. Install PSAuthClient

Open PowerShell **as Administrator** and run once:

```powershell
Install-Module PSAuthClient -Confirm:$false
```

If prompted about an untrusted repository, type `Y` and press Enter.

---

### 2. Create a NinjaOne API App

This is a one-time step in your NinjaOne portal.

1. Log in as a **System Administrator**
2. Go to **Administration → Apps → API → Client App IDs → Add**
3. Fill in the form:
   - **Name:** Any descriptive name, e.g. `DeviceOwnerScript`
   - **Platform:** `Web`
   - **Redirect URI:** `https://localhost`
4. Click **Save** — copy the **Client ID** and **Client Secret**

> ⚠️ The Redirect URI must be set to exactly `https://localhost`. The script will fail to authenticate if this does not match.

---

### 3. Configure the Script

Open `Set-NinjaDeviceOwner.ps1` in any text editor and fill in the configuration block near the top:

```powershell
# ===================================================================
#  CONFIGURATION — Fill in your NinjaOne details here
# ===================================================================
$BaseUrl       = 'https://<your Login URL>'
$TokenEndpoint = 'https://<your Login URL>/ws/oauth/token'
$ClientId      = '<Your Client ID>'
$ClientSecret  = '<Your Client Secret>'
$RedirectUri   = 'https://localhost'
```

| Variable | Example | Where to Find It |
|---|---|---|
| `$BaseUrl` | `https://app.ninjarmm.com` | Your NinjaOne login URL |
| `$TokenEndpoint` | `https://app.ninjarmm.com/ws/oauth/token` | Same as `$BaseUrl` + `/ws/oauth/token` |
| `$ClientId` | `abc123...` | Administration → Apps → API after creating the app |
| `$ClientSecret` | `s3cr3t...` | Shown once at app creation. Regenerate from the portal if lost. |
| `$RedirectUri` | `https://localhost` | **Do not change this.** Must match the portal exactly. |

> ℹ️ **Regional URLs:**
> | Region | Base URL |
> |---|---|
> | US | `https://app.ninjarmm.com` |
> | EU | `https://eu.ninjarmm.com` |
> | Oceania | `https://oc.ninjarmm.com` |
> | Canada | `https://ca.ninjarmm.com` |

---

## Usage

### Parameters

| Parameter | Required | Description |
|---|---|---|
| `-Hostname` | ✅ | Exact system name or DNS name of the device in NinjaOne. Case-insensitive. |
| `-FirstName` | ✅ | First name of the contact/end-user to assign. Must match NinjaOne exactly. |
| `-LastName` | ✅ | Last name of the contact/end-user to assign. Must match NinjaOne exactly. |

### Examples

```powershell
# Assign a desktop
.\Set-NinjaDeviceOwner.ps1 -Hostname "DESKTOP-ABC123" -FirstName "Jane" -LastName "Smith"

# Assign a laptop
.\Set-NinjaDeviceOwner.ps1 -Hostname "LAPTOP-MKTG01" -FirstName "Carlos" -LastName "Rivera"

# Assign a server
.\Set-NinjaDeviceOwner.ps1 -Hostname "WIN-SRV-PROD01" -FirstName "Sarah" -LastName "Thompson"
```

### What Happens When You Run It

The script walks through four labeled steps:

```
[1/4] Authenticating with NinjaOne...
      A browser window will open for you to log in.

[✓] Successfully authenticated.

[2/4] Looking up device with hostname: 'DESKTOP-ABC123'...
[✓] Device found — ID: 12345  |  Name: DESKTOP-ABC123

[3/4] Looking up contact: 'Jane Smith'...
[✓] Contact found — ID: 67890  |  Name: Jane Smith

[4/4] Assigning 'Jane Smith' as owner of 'DESKTOP-ABC123'...

  ================================================
  [✓] SUCCESS
      Device : DESKTOP-ABC123  (ID: 12345)
      Owner  : Jane Smith  (ID: 67890)
  ================================================
```

---

## Bulk Assignment via CSV

To assign owners for many devices at once, loop over a CSV file.

**CSV format (`DeviceOwners.csv`):**

```csv
Hostname,FirstName,LastName
DESKTOP-ABC123,Jane,Smith
LAPTOP-MKTG01,Carlos,Rivera
WIN-WS-HR002,Sarah,Thompson
```

**PowerShell bulk loop:**

```powershell
$rows = Import-Csv -Path '.\DeviceOwners.csv'
foreach ($row in $rows) {
    .\Set-NinjaDeviceOwner.ps1 `
        -Hostname  $row.Hostname `
        -FirstName $row.FirstName `
        -LastName  $row.LastName
}
```

> ⚠️ Running in bulk will trigger a browser login once per script execution. For large batches, consider adding a short delay between iterations to avoid API rate limiting.

---

## How It Works

### API Endpoints Used

| Method | Endpoint | Purpose |
|---|---|---|
| `GET` | `/v2/devices` | Retrieves all devices for local hostname filtering |
| `GET` | `/v2/contacts` | Retrieves all contacts/end-users for name filtering |
| `POST` | `/v2/device/{id}/owner` | Assigns the resolved contact ID as the device owner |

### Hostname Resolution

The NinjaOne API does not support querying devices directly by hostname. The script handles this by:

1. Fetching all devices via `GET /v2/devices`
2. Filtering client-side by matching `systemName` or `dnsName` against the provided hostname
3. Extracting the numeric device `id` from the matched result
4. Using that `id` for all subsequent API calls — including the owner assignment

The hostname you provide is only used as a human-friendly lookup key. All API operations use the internal numeric ID.

> ⚠️ The default page size is `1000`. If your environment has more than 1,000 devices, pagination should be added to ensure all devices are retrieved.

---

## Troubleshooting

| Error / Symptom | Solution |
|---|---|
| `PSAuthClient module not found` | Run `Install-Module PSAuthClient -Confirm:$false` as Administrator |
| Browser opens but login fails | Verify `$BaseUrl` matches your NinjaOne region. Check your login credentials. |
| Script closes immediately without browser | Confirm Redirect URI in the NinjaOne portal is exactly `https://localhost` |
| `No device found with hostname` | Check exact spelling in NinjaOne → Administration → Devices. Spelling must be exact (case-insensitive). |
| `No contact found matching name` | User must exist as a Contact or End User under an Organization in NinjaOne. Verify spelling. |
| `Multiple devices matched hostname` | Two devices share the same name. Rename one in NinjaOne to make hostnames unique. |
| HTTP `403` on owner assignment | API app is missing the `management` scope. Recreate the app or contact your NinjaOne admin. |
| HTTP `400` on owner assignment | Contact and device belong to different NinjaOne organizations. |
| HTTP `401` Unauthorized | Access token issue — verify `$ClientSecret` is correct. |
| PowerShell execution policy error | Run `Set-ExecutionPolicy -Scope CurrentUser -ExecutionPolicy RemoteSigned` then retry. |

---

## Security Notes

- **Do not commit your `$ClientSecret` to source control.** Consider using [PowerShell SecretManagement](https://learn.microsoft.com/en-us/powershell/module/microsoft.powershell.secretmanagement/) or prompting for credentials at runtime.
- The script does not write tokens to disk — the access token exists in memory only for the duration of execution.
- Restrict the NinjaOne API app to only the scopes it needs: `monitoring` and `management`.
- Only System Administrators should configure and run this script.

---

## Pre-Flight Checklist

Before your first run:

- [ ] PowerShell 5.1 or later is installed (`$PSVersionTable`)
- [ ] `Install-Module PSAuthClient -Confirm:$false` run as Administrator
- [ ] API app created in NinjaOne with Redirect URI = `https://localhost`
- [ ] Client ID and Client Secret copied from the portal
- [ ] All four variables filled in the script configuration block
- [ ] Target user exists as a Contact or End User in NinjaOne
- [ ] Target device hostname is visible in NinjaOne
- [ ] Tested with one device before running a bulk CSV
