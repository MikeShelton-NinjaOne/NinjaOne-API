# NinjaOne Ticket Import Script

A PowerShell utility for importing tickets from an Excel spreadsheet into NinjaOne Ticketing via the public API. Designed for new customers migrating ticket data from a previous ticketing system.

---

## ⚠️ IMPORTANT — NOT SUPPORTED BY NINJAONE

> **This script is provided as-is and is NOT supported by NinjaOne or NinjaOne Support.**
>
> - This is a **community / customer-built resource**, not an official NinjaOne product.
> - **NinjaOne Support will not troubleshoot, debug, or assist with this script** under any circumstances.
> - Please do **NOT** open a support ticket with NinjaOne about issues related to this script, its behavior, errors it produces, or the data it imports.
> - There is **no warranty, guarantee, or service-level agreement** of any kind. Use of this script is entirely at your own risk.
> - You are solely responsible for the data you import, the API credentials you create, and any changes this script makes to your NinjaOne instance.
> - Always run the script in **dry-run mode (`DryRun = $true`) first** and verify the results before performing a real import. Once tickets are imported, undoing them is your responsibility.

If you need help, please refer to the [User Guide](./Ticket-Import-User-Guide.md), the included log file (`ticket-import-log.txt`), or reach out to whoever shared this script with you.

---

## What's in this repository

| File | Purpose |
|---|---|
| `Import-Tickets.ps1` | The PowerShell import script. Edit the configuration block at the top before running. |
| `Sample-Tickets.xlsx` | A sample spreadsheet with example tickets, including a couple of intentionally invalid rows so you can see how validation works. |
| `Ticket-Import-User-Guide.md` | Full step-by-step user guide — start here if you've never used the script before. |
| `Ticket-Import-Prompt.md` | The original prompt used to generate the script, in case you want to regenerate or customize it. |

## Quick start

1. Read the [User Guide](./Ticket-Import-User-Guide.md) — it walks through everything from creating your API client app in NinjaOne to running the real import.
2. Install the prerequisite module (one time):
   ```powershell
   Install-Module ImportExcel -Scope CurrentUser
   ```
3. Open `Import-Tickets.ps1` in Notepad or VS Code and fill in the `# === CONFIGURATION ===` block.
4. Run a dry run first:
   ```powershell
   .\Import-Tickets.ps1
   ```
5. When the dry run looks good, set `DryRun = $false` and run it again for the real import.

## Requirements

- Windows 10 or 11 with PowerShell 5.1+ (or PowerShell 7+)
- The `ImportExcel` PowerShell module
- An active NinjaOne account with administrator access (to create the API client app)
- A NinjaOne API Client App configured with:
  - Application Platform: `API Services (machine-to-machine)`
  - Redirect URI: `https://localhost`
  - Scopes: `Management`
  - Allowed Grant Types: `Client Credentials`

## Disclaimer

This script is shared as a starting point for ticket migration. It has not been audited, certified, or approved by NinjaOne. Review the code yourself before running it in any environment you care about, and **always test with `DryRun = $true` before doing a real import**.

NinjaOne, NinjaRMM, and related marks are trademarks of NinjaOne, LLC. This project is not affiliated with, endorsed by, or sponsored by NinjaOne.
