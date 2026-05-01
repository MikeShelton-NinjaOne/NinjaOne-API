# Generate-per-organization-patch-status-reports.ps1

A PowerShell script that connects to the **NinjaOne API**, pulls every Windows device across every organization in your tenant, gathers patch data for each device, and produces a folder of self-contained HTML reports — one per organization — plus a top-level `index.html` linking them together.

> ⚠️ **Unsupported example.** The script's own header notes it is not supported by NinjaOne support. Use it as a starting point and adapt as needed.

---

## What the script does

1. Authenticates to the NinjaOne API using OAuth 2.0 Client Credentials.
2. Fetches every organization in your tenant (paginated).
3. Fetches every device in your tenant (paginated).
4. Filters devices down to Windows-only node classes:
   - `WINDOWS_WORKSTATION`
   - `WINDOWS_SERVER`
   - `WINDOWS_LAPTOP`
   - `WINDOWS_DESKTOP`
5. For each Windows device, calls four NinjaOne patch endpoints:
   - `/v2/device/{id}/os-patches` — pending/approved/failed/rejected OS patches
   - `/v2/device/{id}/software-patches` — pending/approved/failed/rejected third-party patches
   - `/v2/device/{id}/os-patch-installs` — installation history for OS patches
   - `/v2/device/{id}/software-patch-installs` — installation history for third-party patches
6. Categorizes each patch as **Pending**, **Approved**, **Installed**, or **Other**.
7. Filters records to the **previous calendar month** (the reporting window — see below).
8. Writes one HTML file per organization, plus an `index.html` that summarizes all orgs.
9. Auto-opens `index.html` in the default browser when finished.

---

## Reporting window

The script always reports on the **previous full calendar month**, calculated in local time. For example, if you run it any time during April, the window is `March 1 00:00:00` through `March 31 23:59:59`. There is no command-line option to change this — if you need a different window, edit the `$Script:ReportWindow` block near the top of the script.

The two open-patch endpoints (`os-patches`, `software-patches`) don't accept date parameters, so the script fetches them in full and filters client-side. The two install endpoints (`os-patch-installs`, `software-patch-installs`) **do** accept `installedAfter` / `installedBefore`, so the script narrows those server-side using epoch-second timestamps converted from the local-time window.

---

## Requirements

| Requirement | Notes |
|---|---|
| PowerShell | Windows PowerShell **5.1** (the script declares `#Requires -Version 5.1`). |
| Modules | None — uses only built-in cmdlets. |
| Network | TLS 1.2 (the script forces this). Outbound HTTPS to your NinjaOne region. |
| Credentials | A NinjaOne API client (Client ID + Client Secret) with the `monitoring` scope. |
| Permissions | The API client must be able to read organizations, devices, and patch data. |

The script forces `[Net.ServicePointManager]::SecurityProtocol = Tls12` because NinjaOne requires TLS 1.2.

---

## Configuration — edit before running

Open the script and edit the `$Script:Config` hashtable near the top. **The placeholders must be replaced** or the script will fail at the auth step.

```powershell
$Script:Config = @{
    BaseUrl          = 'https://<your Login URL>'
    TokenEndpoint    = 'https://<your login URL>/ws/oauth/token'
    ClientId         = '<Your Client ID>'
    ClientSecret     = '<Your Client Secret>'
    Scope            = 'monitoring'
    PageSize         = 1000
    MaxRetries       = 5
    AllowedNodeClass = @(
        'WINDOWS_WORKSTATION',
        'WINDOWS_SERVER',
        'WINDOWS_LAPTOP',
        'WINDOWS_DESKTOP'
    )
}
```

### Field-by-field

- **`BaseUrl`** — Your NinjaOne tenant base URL. This is the same hostname you log in to. Examples: `https://app.ninjarmm.com`, `https://eu.ninjarmm.com`, `https://ca.ninjarmm.com`, `https://oc.ninjarmm.com`. Do **not** include a trailing slash.
- **`TokenEndpoint`** — The OAuth token URL on the same host: `https://<host>/ws/oauth/token`.
- **`ClientId` / `ClientSecret`** — Generated in NinjaOne under **Administration → Apps → API → Client App IDs**. Create a *Client Credentials* app and grant it the `monitoring` scope.
- **`Scope`** — Leave as `monitoring`. Without this scope the install-history endpoints will return empty.
- **`PageSize`** — How many records per page when listing orgs and devices. 1000 is the API maximum.
- **`MaxRetries`** — How many times to retry on `429` (rate limit) or `5xx` (server) errors before giving up. The script uses exponential backoff and honors any `Retry-After` header.
- **`AllowedNodeClass`** — Node-class strings to include. Anything not in this list is filtered out. Edit this list if you want to include other Windows variants.

> 🔒 **Security note.** The script stores credentials inline. The header comments explicitly recommend rotating them after testing and moving them to **Windows Credential Manager**, an environment variable, or a secret vault for production use. Do not commit edited copies of this script to source control with real credentials in place.

---

## Running the script

Open Windows PowerShell, change to the directory containing the script, and run it:

```powershell
PS C:\> cd C:\path\to\script
PS C:\path\to\script> .\Generate-per-organization-patch-status-reports.ps1
```

If your execution policy blocks unsigned local scripts, either sign it or temporarily relax the policy for the current session:

```powershell
Set-ExecutionPolicy -Scope Process -ExecutionPolicy Bypass
```

There are no command-line parameters. Everything is driven by the configuration block.

---

## What you'll see while it runs

The script writes color-coded progress to the host. A typical run looks like this:

```
Reporting window (previous calendar month): 2026-03-01 to 2026-03-31
Requesting OAuth token from https://ca.ninjarmm.com/ws/oauth/token ...
Authentication successful. Token type: Bearer
Fetching organizations ...
  Orgs page 1: 47. Running total: 47
Total organizations: 47
Fetching devices (page size 1000) ...
  Devices page 1: 1000. Running total: 1000
  Devices page 2: 312. Running total: 1312
Total devices retrieved (before filter): 1312
Filtered out 184 non-Windows device(s). Windows devices to process: 1128
[progress bar: Collecting patches]
Writing reports to: C:\path\to\script\NinjaOne-PatchReports-20260501-141233
  Acme Corp                                devices=42   pending=18   approved=7    installed=124
  Beta Industries                          devices=9    pending=2    approved=0    installed=31
  ...
Index written to: C:\...\NinjaOne-PatchReports-20260501-141233\index.html

Install endpoint diagnostics:
  OS patch installs:       API returned 1742, kept after date filter: 1742
  Software patch installs: API returned 893, kept after date filter: 893
```

A `Write-Progress` bar shows the current device being processed (e.g. `[412/1128] Acme Corp :: WIN-DC01`). The script handles `429` rate-limit responses automatically by sleeping for the duration in `Retry-After` (or using exponential backoff if no header is provided), and refreshes the token automatically on a `401`.

---

## Output

Each run creates a new timestamped subfolder in the **current working directory**:

```
NinjaOne-PatchReports-20260501-141233\
├── index.html
├── Acme Corp-PatchReport.html
├── Beta Industries-PatchReport.html
├── Unassigned-PatchReport.html       (only if any devices have no org)
└── ...
```

### `index.html`

The landing page. Contains:

- Reporting window
- Total Windows devices included, non-Windows devices filtered out, total devices seen
- A sortable table listing each organization with its device count and patch totals (Pending / Approved / Installed)
- Each org name links to that org's detail report
- A grand-totals row at the bottom

### Per-organization HTML reports

Each org file is self-contained — no external CSS, JS, or images — so you can email them or archive them as-is. Each report includes:

- Summary cards for the org (Pending / Approved / Installed totals)
- A live filter box that searches devices by hostname
- A collapsible section per device, each showing:
  - Hostname, OS info, last contact time, and badges with that device's patch counts
  - One or more tables listing the actual patches under that device, with status, source endpoint, and the most relevant timestamp

Filenames are sanitized (invalid Windows filename characters become `_`, and Windows reserved names like `CON` or `PRN` get prefixed with `_`). If two organizations would generate the same filename, the second one gets the org ID appended for uniqueness. Devices with no resolvable organization land in `Unassigned-PatchReport.html`.

---

## How patches are categorized

The `Get-PatchCategory` function maps raw API fields to four buckets:

| Bucket | What it means | Source signal |
|---|---|---|
| **Pending** | Patch detected but not yet approved or installed. | `status` starts with `PENDING` on the open-patch endpoints. |
| **Approved** | Patch is approved/scheduled but not yet installed. | `status` starts with `APPROVED`, `MANUAL` (manual approval), or `AUTO` (auto-approved). |
| **Installed** | Patch was installed successfully. | Record came from a `*-patch-installs` endpoint with `INSTALLED` / `SUCCESS` / `SUCCEEDED`, or open-patch records flagged the same way via `installResult`. |
| **Other** | Anything else, including `FAILED` install attempts and unrecognized statuses. | Failed installs are explicitly bucketed here so they don't inflate the Installed count. |

Records from the install endpoints are always treated as install records via a `SourceHint` parameter, so the categorizer doesn't have to guess from ambiguous status fields.

---

## Diagnostics

After writing the reports, the script prints an **install endpoint diagnostics** block:

```
Install endpoint diagnostics:
  OS patch installs:       API returned 1742, kept after date filter: 1742
  Software patch installs: API returned 893, kept after date filter: 893
```

If both numbers are zero, the script warns you with one of two probable causes:

- No patches were actually installed during the reporting window, **or**
- Your API credentials don't have the `monitoring` scope needed to read install history.

If the API returned records but the date filter dropped them all, the script dumps a sample install record so you can see exactly which date field the API is returning. You can use that to tweak the `$candidateFields` list inside `Get-PatchRelevantDate` if your tenant uses a different field name.

---

## Common issues

**`FATAL: Authentication failed.`**
The placeholders in `$Script:Config` weren't replaced, the secret is wrong, the `BaseUrl` host doesn't match where you log in, or the API client doesn't have the `monitoring` scope.

**Reports show zero installed patches.**
Check the diagnostics block at the end of the run. The most common cause is missing `monitoring` scope on the API client. The next most common is that the previous calendar month genuinely had no patch installs (rare in real environments).

**Rate-limit messages.**
You'll see `Rate limited (429). Retrying in N s ...` lines. The script handles these automatically up to `MaxRetries` times. If retries are exhausted the call throws and the device's patch fetch is skipped (with a yellow warning), but the rest of the run continues.

**`Got 401. Refreshing token and retrying.`**
The bearer token expired mid-run. The script refreshes it and retries the call once or twice. This is normal for long runs.

**Some devices land in `Unassigned-PatchReport.html`.**
Those devices have an `organizationId` that didn't resolve to any org name returned by `/v2/organizations`, or no `organizationId` at all. Check those devices in the NinjaOne console.

**Filename collision warning skipped silently.**
If two orgs share a name, the second one gets its org ID appended to the filename automatically. Check the index page for two rows with the same display name.

---

## Customizing the script

Common edits, with rough pointers:

- **Change the reporting window** → edit the `$Script:ReportWindow = & { ... }` block near the top.
- **Include more device types** → add to `AllowedNodeClass` in `$Script:Config`. For example, add `'MAC'` and `'LINUX'`-prefixed classes if you also patch those.
- **Output to a different folder** → edit the `$outputDir = Join-Path ...` line inside `Main`.
- **Skip the auto-open at the end** → comment out or remove the `Start-Process -FilePath $indexPath` line at the end of `Main`.
- **Change the look of the HTML** → edit `Get-ReportCss`. The CSS is embedded inline in every output file, so just rerun the script after editing.
- **Change patch categorization rules** → edit `Get-PatchCategory`.

---

## File summary

| Section in script | Purpose |
|---|---|
| `$Script:Config` | All tenant-specific settings. |
| `$Script:ReportWindow` | Computes previous-month start/end in both local time and epoch seconds. |
| `ConvertTo-SafeValue`, `Format-HtmlValue`, `Format-EpochOrIso`, `ConvertTo-DateTimeOrNull` | Small helpers for null-safe formatting. |
| `Get-NinjaToken` | OAuth 2.0 Client Credentials authentication. |
| `Invoke-NinjaApi` | Generic GET wrapper with retry/backoff for `429`, `401`, and `5xx`. |
| `Get-AllOrganizations`, `Get-AllDevices` | Paginated list calls. |
| `Get-DevicePatches` | Calls all four patch endpoints for one device, applies the date window. |
| `Get-PatchCategory` | Maps raw status fields to Pending/Approved/Installed/Other. |
| `Get-ReportCss`, `Get-FilterJs` | Inline CSS and the JS hostname filter shared by every HTML report. |
| `New-OrganizationReport` | Builds one org's HTML file. |
| `New-IndexReport` | Builds the top-level `index.html`. |
| `Main` | Orchestrates everything. |
