# Set-NinjaDevicePurchaseDate.ps1

> Reads the warranty start date from a NinjaOne device and writes it to the ITAM Purchase Date field on the same device — silently, with no browser login required.

---

## Table of Contents

- [What This Script Does](#what-this-script-does)
- [How the Date Conversion Works](#how-the-date-conversion-works)
- [Prerequisites](#prerequisites)
- [Part 1: Enable ITAM in NinjaOne](#part-1-enable-itam-in-ninjaone)
- [Part 2: Set API Permission on the ITAM Field](#part-2-set-api-permission-on-the-itam-field)
- [Part 3: Enable Warranty Tracking](#part-3-enable-warranty-tracking)
- [Part 4: Create the API App](#part-4-create-the-api-app)
- [Part 5: Configure the Script](#part-5-configure-the-script)
- [Part 6: Run the Script](#part-6-run-the-script)
- [Finding Your Device ID](#finding-your-device-id)
- [API Endpoints Used](#api-endpoints-used)
- [Troubleshooting](#troubleshooting)
- [Pre-Flight Checklist](#pre-flight-checklist)

---

## What This Script Does

1. Authenticates silently to NinjaOne using Client Credentials — no browser, no login prompt
2. Fetches the device details including warranty data using `?expand=warranty`
3. Reads the `warranty.startDate` field from the response
4. Converts it from Unix epoch seconds to milliseconds (required by the ITAM field)
5. Writes the value to the `itamAssetPurchaseDate` field via the custom fields endpoint
6. Prints a confirmation showing the device name, field written, and date value set

The script exits cleanly with a clear message if the device has no warranty start date set, so it is safe to run without worrying about overwriting a purchase date with a blank value.

---

## How the Date Conversion Works

NinjaOne stores warranty dates and ITAM dates in different units:

| Field | Unit | Example |
|---|---|---|
| `warranty.startDate` (GET response) | Unix epoch **seconds** | `1672531200` |
| `itamAssetPurchaseDate` (PATCH body) | Unix epoch **milliseconds** | `1672531200000` |

The script multiplies the warranty start date by 1000 to convert seconds to milliseconds before writing. No manual conversion is needed.

---

## Prerequisites

| Requirement | Notes |
|---|---|
| PowerShell 5.1+ | Built-in on Windows 10/11. Run `$PSVersionTable` to check. |
| NinjaOne System Administrator | Required to create the API app and configure ITAM |
| NinjaOne ITAM enabled | Must be active on your account — see Part 1 |
| Warranty Tracking enabled | Device must have a warranty start date — see Part 3 |
| No extra modules | Uses only built-in PowerShell |

---

## Part 1: Enable ITAM in NinjaOne

NinjaOne ITAM (IT Asset Management) must be enabled before the `itamAssetPurchaseDate` field is available.

1. Log into NinjaOne as a **System Administrator**
2. Go to: **Administration → Apps → IT Asset Management**
3. Enable the ITAM app if it is not already active
4. Click **Save**

> ℹ️ If ITAM is already enabled you will see asset fields on device Overview tabs. Skip to Part 2.

---

## Part 2: Set API Permission on the ITAM Field

The `itamAssetPurchaseDate` field must have API write permission enabled. Without this the PATCH call will return a 403 or silently ignore the field.

1. Go to: **Administration → Devices → Custom Fields**
2. Find the **Purchase Date** field (internal name: `itamAssetPurchaseDate`)
3. Click to edit it
4. Under **API Access** set the permission to **Read/Write**
5. Click **Save**

> ⚠️ If you cannot see this field, ITAM is not enabled. Complete Part 1 first.

---

## Part 3: Enable Warranty Tracking

The script reads `warranty.startDate` from the device. This field is only populated if Warranty Tracking is enabled and the device has a warranty start date set.

### Enable Warranty Tracking globally

1. Go to: **Administration → Apps → Warranty Tracking**
2. Enable the app if not already active
3. Click **Save**

### Set or sync a warranty start date on the device

1. Open the device in NinjaOne
2. Go to the **Overview** tab
3. Find the **Warranty date** field in the Device Information card
4. Either:
   - Click the **circling arrows** icon to automatically sync with the manufacturer (supported manufacturers only)
   - Click the date link to set the start and end dates manually

> ℹ️ NinjaOne supports automatic warranty sync for workstations, servers, and virtual hosts from supported manufacturers. HP Enterprise (HPE) devices are not supported for automatic sync — set dates manually.

> ℹ️ Automatic warranty sync runs every Saturday at 1:00 AM. Run a manual sync if you need current data before the next scheduled run.

---

## Part 4: Create the API App

One-time setup.

1. Log into NinjaOne as a **System Administrator**
2. Go to: **Administration → Apps → API → Client App IDs → Add**
3. Fill in:

   | Field | Value |
   |---|---|
   | **Name** | Any name, e.g. `SetPurchaseDate` |
   | **Platform** | `API Services (Machine-to-Machine)` |
   | **Allowed Scopes** | ✅ `monitoring` AND ✅ `management` |
   | **Redirect URI** | Leave blank |

4. Click **Save** — copy the **Client ID** and **Client Secret** immediately

> ⚠️ The Client Secret is shown only once. Save it somewhere secure before closing the page.

---

## Part 5: Configure the Script

Open `Set-NinjaDevicePurchaseDate.ps1` in any text editor and fill in the **CONFIGURATION** block at the top:

```powershell
# ==============================================================================
#  CONFIGURATION -- Fill in ALL values before running
# ==============================================================================

$BaseUrl       = 'https://<your Login URL>'
$TokenEndpoint = 'https://<your Login URL>/ws/oauth/token'
$ClientId      = '<Your Client ID>'
$ClientSecret  = '<Your Client Secret>'
$DeviceId      = 0
```

| Variable | Example | Notes |
|---|---|---|
| `$BaseUrl` | `https://app.ninjarmm.com` | Your NinjaOne login URL — no trailing slash |
| `$TokenEndpoint` | `https://app.ninjarmm.com/ws/oauth/token` | Same as `$BaseUrl` + `/ws/oauth/token` |
| `$ClientId` | `abc123...` | From Part 4 |
| `$ClientSecret` | `s3cr3t...` | From Part 4 — shown once |
| `$DeviceId` | `3` | Numeric device ID — see below for how to find it |

**Regional URLs:**

| Region | Base URL |
|---|---|
| United States | `https://app.ninjarmm.com` |
| Europe | `https://eu.ninjarmm.com` |
| Oceania | `https://oc.ninjarmm.com` |
| Canada | `https://ca.ninjarmm.com` |

---

## Part 6: Run the Script

```powershell
.\Set-NinjaDevicePurchaseDate.ps1
```

Example output when successful:

```
  ================================================================
  NinjaOne -- Write Warranty Start Date to ITAM Purchase Date
  Device ID: 3
  ================================================================

  [1/4] Authenticating...
  [OK] Authenticated.

  [2/4] Fetching device 3 with warranty data...
  [OK] Device found:
       Name      : SE-W19E
       Org ID    : 4
       Warranty start : 2023-01-15 (epoch: 1673740800 seconds)

  [3/4] Converting warranty start date for ITAM field...
  [OK] Warranty start  : 2023-01-15
       ITAM value (ms) : 1673740800000

  [4/4] Writing Purchase Date to device 3...
  [OK] Purchase Date written successfully.
       Device    : SE-W19E (ID: 3)
       Field     : itamAssetPurchaseDate
       Value set : 2023-01-15

  ================================================================
  [OK] COMPLETE
       Device    : SE-W19E (ID: 3)
       Purchase Date set to : 2023-01-15
  ================================================================
```

Example output when no warranty start date is set:

```
  [2/4] Fetching device 3 with warranty data...
  [OK] Device found:
       Name      : SE-W19E
  [!] No warranty start date found on device 3.
      Ensure Warranty Tracking is enabled and the device has a warranty start date set.
      In NinjaOne: device Overview tab > Warranty date > set start date or Sync with manufacturer.
```

The script exits cleanly without writing anything if no warranty start date is found.

---

## Finding Your Device ID

Open the device in NinjaOne and read the number from the browser URL:

```
https://app.ninjarmm.com/#/deviceDashboard/3/overview
                                           ^
                                    This is the device ID
```

Alternatively, use the Devices search grid — add the **ID** column from the **Columns** menu to see device IDs alongside device names.

---

## API Endpoints Used

| Method | Endpoint | Purpose |
|---|---|---|
| `POST` | `/ws/oauth/token` | Client Credentials auth |
| `GET` | `/v2/device/{id}?expand=warranty` | Fetch device details including warranty dates |
| `PATCH` | `/v2/device/{id}/custom-fields` | Write the ITAM Purchase Date field |

### Why `?expand=warranty`

The warranty dates are not included in the standard `GET /v2/device/{id}` response. The `?expand=warranty` query parameter must be appended to the URL to include the `warranty` object in the response. Without it the warranty fields are absent and the script will exit with a "no warranty start date found" message even if warranty dates are set in NinjaOne.

### Why the custom-fields endpoint for an ITAM field

NinjaOne ITAM core asset fields (Purchase Date, Asset ID, Serial Number, etc.) are exposed through the same `custom-fields` endpoint as custom fields, using internal field names. The internal name for Purchase Date is `itamAssetPurchaseDate`. The body is a flat JSON object keyed by the internal field name:

```json
{ "itamAssetPurchaseDate": 1673740800000 }
```

---

## Troubleshooting

| Error / Symptom | Solution |
|---|---|
| `Fill in $BaseUrl` | Placeholder still in config. Replace all `<...>` values. |
| `Authentication failed` | Check `$BaseUrl`, `$ClientId`, `$ClientSecret`. App must be `API Services (Machine-to-Machine)` with `monitoring` and `management` scopes. |
| `Device X not found` | Device ID doesn't exist or your account doesn't have access to it. Verify the ID from the device URL in NinjaOne. |
| `No warranty start date found` | The device has no warranty start date set. Open the device in NinjaOne, go to Overview tab, and either sync with the manufacturer or set the date manually. |
| `Failed to write Purchase Date (HTTP 403)` | The `itamAssetPurchaseDate` field does not have API Read/Write permission. Complete Part 2. |
| `Failed to write Purchase Date (HTTP 400)` | The value format is wrong. This should not happen if the script is run unmodified — file an issue if you see this. |
| Purchase Date field shows wrong date in NinjaOne | Verify the warranty start date on the device Overview tab is correct. The script writes exactly what the warranty start date contains — if the warranty date is wrong the purchase date will be wrong too. |
| Purchase Date does not appear in NinjaOne UI | ITAM may not be enabled. Complete Part 1. |
| Script completes but field is still blank | The `itamAssetPurchaseDate` field may not have API write permission. Complete Part 2. |
| Wrong date in NinjaOne (off by hours) | The script uses UTC epoch milliseconds which NinjaOne converts to local time for display. The stored value is correct — the display depends on your NinjaOne timezone setting. |

---

## Pre-Flight Checklist

- [ ] PowerShell 5.1+ confirmed (`$PSVersionTable`)
- [ ] NinjaOne ITAM enabled (Administration → Apps → IT Asset Management)
- [ ] `itamAssetPurchaseDate` field API permission set to Read/Write
- [ ] Warranty Tracking enabled (Administration → Apps → Warranty Tracking)
- [ ] Target device has a warranty start date set (Overview tab → Warranty date)
- [ ] API app created — Platform: `API Services (Machine-to-Machine)`, Scopes: `monitoring` + `management`
- [ ] Client ID and Client Secret saved securely
- [ ] All config variables filled in — no `<placeholder>` text remaining
- [ ] `$DeviceId` set to the correct numeric device ID
- [ ] Script run — confirmed `[OK] COMPLETE` in console output
- [ ] Purchase Date confirmed in NinjaOne device Overview tab
