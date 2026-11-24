# NinjaOne Dashboard Monitor

`NinjaOne-ServerDashboard.ps1` reads the NinjaOne REST API and writes a single self-refreshing HTML page. The script re-queries NinjaOne every 30 seconds and rewrites the file. The page reloads itself, so you can leave it open in a browser.

It is read-only: it never changes anything in NinjaOne.

---

## Requirements

| Item | Detail |
|---|---|
| PowerShell | Windows PowerShell 5.1 or PowerShell 7+ |
| Network | HTTPS access to your NinjaOne instance host |
| NinjaOne API client | Client credentials (machine-to-machine) with the **Monitoring** scope |

### Create the API client

1. In NinjaOne go to **Administration > Apps > API > Client app IDs > Add client app**.
2. Set **Application platform** to *API Services (machine-to-machine)*.
3. Tick the **Monitoring** scope. Tick **Management** too if ticketing or other data comes back empty.
4. Set **Allowed grant types** to *Client Credentials*.
5. Save, and copy the **Client ID** and **Client Secret**.

### Instance host

Use the host that matches your tenant: `app.ninjarmm.com`, `us2.ninjarmm.com`, `ca.ninjarmm.com`, `eu.ninjarmm.com` or `oc.ninjarmm.com`.

---

## Running it

```powershell
.\NinjaOne-ServerDashboard.ps1 -Instance app.ninjarmm.com -ClientId 'xxxx' -ClientSecret 'yyyy' -OpenBrowser
```

Using environment variables so the secret stays out of the command line:

```powershell
$env:NINJA_CLIENT_ID     = 'xxxx'
$env:NINJA_CLIENT_SECRET = 'yyyy'
.\NinjaOne-ServerDashboard.ps1 -Instance ca.ninjarmm.com -OpenBrowser
```

The script keeps running and refreshing until you stop it with Ctrl+C. Restart it after replacing the script file to pick up changes.

### Parameters

| Parameter | Default | Description |
|---|---|---|
| `-Instance` | `app.ninjarmm.com` | NinjaOne instance host |
| `-ClientId` | `$env:NINJA_CLIENT_ID` | API client ID |
| `-ClientSecret` | `$env:NINJA_CLIENT_SECRET` | API client secret |
| `-OutputPath` | `NinjaOne-Server-Dashboard.html` next to the script | Where the HTML file is written |
| `-RefreshSeconds` | `30` | Seconds between refreshes |
| `-NodeClasses` | `WINDOWS_SERVER,LINUX_SERVER,MAC_SERVER` | Device classes treated as "servers" |
| `-OpenBrowser` | off | Open the dashboard in your default browser on start |
| `-Once` | off | Generate the file once and exit, with no loop |

The file is written to a temporary file and then moved into place, so the browser never loads a half-written page.

---

## What the dashboard shows

Every section is a card. Each card can be collapsed, resized by dragging its corner, and moved by dragging the grip in its title bar. Layout is remembered in the browser. **Reset layout** returns everything to default.

### Servers table
One row per server, sorted by health: Unknown first, then Unhealthy, Needs attention, Healthy.

| Column | Meaning |
|---|---|
| Server | NinjaOne display name, with the device class underneath |
| Health | NinjaOne's own health status (Healthy, Needs attention, Unhealthy, Unknown) |
| Status | Online (green) or Offline (red) |
| Alerts | Active alert counts by severity, as colored badges |
| Reboot pending | Whether the OS reports a pending reboot |
| Uptime | Time since last boot, or Offline |
| Disk used | Every volume on the server |
| Backup | Backup status and last job result and time |
| Pending patches | Patches not yet installed (not failed or rejected) |
| Failed patches | Patches that failed to install |
| Last patch scan / apply | Time and Success or Failed for each |

A legend explaining the badge colors sits at the bottom of the card.

### Donut cards
| Card | Shows |
|---|---|
| All Servers vs Healthy Servers | The four NinjaOne statuses: Healthy (green), Needs attention (yellow), Unhealthy (red), Unknown (grey) |
| Online vs Offline Servers | Online (green) against Offline (red), with the total |
| OS Patching Compliance | Patch-enabled servers that are fully patched, plus enablement and scan statistics. Rejected patches are ignored |
| OS Patches Installed | Installed (last 90 days), Approved, Failed and Pending operating-system patches |
| Software Patches Installed | The same breakdown for third-party software patches |
| Healthy vs Unhealthy by Threat | Servers with and without active, blocked or quarantined threats |
| NMS Devices Up vs Down | Network-monitoring devices that are up against down |
| Cloud Monitors Up vs Down | Cloud monitors that are up against down |
| BitLocker: Windows Desktops & Laptops | Workstations with BitLocker enabled, disabled, or with no volume data |

### Tables and tiles
| Card | Shows |
|---|---|
| Threats | Active, blocked and quarantined threats, most urgent first |
| Devices in Maintenance Mode | Device, organization, status, start and end time, and the features suppressed |
| Ticketing | Tiles for each top-level ticket status plus Unassigned, then every status with its child statuses and counts |
| NMS Devices Down | NMS devices currently down, with type, organization and last uptime |
| Cloud Monitors in Alert State | Cloud monitors with active alerts: condition health, message and time triggered |
| Backup Overview | Backup jobs (last result), integrity check results, and local and cloud storage usage |
| Active Alerts | Alerts grouped by server, sorted red then yellow then the rest, newest first. Checkboxes filter by severity |

---

## How things are calculated

- **Health**: NinjaOne's own status for the device. If the health query returns nothing, the script derives it: offline, then any critical or major alert means Unhealthy, any other alert or failed patch or pending reboot means Needs attention, otherwise Healthy. Edit `Get-HealthStatus` to change this.
- **Alert severity**: many alerts report severity `NONE`, so the script also uses the condition health status. UNHEALTHY is treated as red and NEEDS_ATTENTION as yellow.
- **Pending patches**: NinjaOne uses several "not installed" statuses (pending, approved, manual and others). The script reads all patch records and classifies them itself, so none are missed.
- **Patch-enabled**: inferred from whether a server has any patch records.
- **Ticket statuses**: read live from the tenant, so custom parent and child statuses appear automatically. Status colors follow the dots NinjaOne shows beside each status, matched by name. Unassigned excludes Resolved and Closed-type tickets.
- **NMS devices**: every device whose class starts with `NMS_`. Last uptime is the last time NinjaOne heard from the device.
- **Cloud monitors**: devices of class `CLOUD_MONITOR_TARGET`.
- **BitLocker**: a Windows workstation counts as enabled when its system drive (C:) reports BitLocker protection on.
- **Backup usage**: summed from NinjaOne's per-organization backup usage, the same source as the Backup > Device overview page.

### Refresh rates

The page refreshes every 30 seconds. Some data is slow to collect, so it is cached and re-read every 5 minutes: ticket counts, backup usage and BitLocker status.

---

## Partial data banner

If one section's API call fails, the dashboard still renders and shows an amber **Partial data** banner naming the section and the error. The rest of the page keeps working. Common causes are a missing API scope or an endpoint your tenant doesn't expose.

---

## Known limitations

| Area | Limitation |
|---|---|
| Backup job counts | The API's job list only returns completed cloud jobs from the last several days. Failed, warning and local jobs come from the activity feed, so counts can be lower than the NinjaOne console shows |
| Boot verification | No API source was found, so the card shows integrity-check results instead |
| Integrity checks | The endpoint spelling varies, so the script tries three paths. If none works, that group reads "unavailable" |
| Ticketing | Uses the ticketing board endpoints and may need the Management scope |
| Antivirus threats | Field names in the threats query were not fully verified. If the Threats card is empty, check the banner |
| Maintenance features | NinjaOne does not store suppressed features on the device. The script reads them from each device's activity history, so older entries may be missed |
| Removed ideas | Next patch scan or apply dates, per-device CPU and memory use, and vulnerability counts are not available through the API and are not shown |
| Scope | Server cards cover the classes in `-NodeClasses`. NMS, cloud monitor, BitLocker and backup overview cards cover their own device sets |

---

## Troubleshooting

| Symptom | Likely cause and fix |
|---|---|
| `403 Forbidden` on the token request | Wrong instance host, or the client is missing the Monitoring scope |
| `404 Not Found` on the token request | Make sure `-Instance` is a bare host with no `https://` |
| `Cannot convert Object[] to Int32` | Fixed in current versions. Re-download the script |
| A card says "unavailable" | Read the amber banner at the top for the failing call |
| Cards overlap or sit in odd places | Click **Reset layout** |
| Ticketing, threats or other cards are empty | Add the Management scope to the API client and restart |
| Numbers differ from the NinjaOne console | See **Known limitations**, especially the backup job counts |

---

## Security notes

- Treat the Client Secret like a password. Prefer environment variables over typing it into a shell history.
- The generated HTML contains server names, alert text and organization names. Store it somewhere appropriate, and do not publish it.
- The script only reads data. Use an API client with the minimum scopes you need.
