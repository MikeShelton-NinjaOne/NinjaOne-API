# InsightVM → NinjaOne Vulnerability Sync

A PowerShell script that pulls CVE data from Rapid7 InsightVM and uploads it into a NinjaOne Vulnerability Management scan group, using NinjaOne's Rapid7 Vulnerability Importer.

Each time the script runs, it uploads the full current list of qualifying CVEs. NinjaOne always displays the data from the most recent upload, so nothing piles up or duplicates between runs — each run simply replaces the last.

## What you need before you start

- **PowerShell 7 or newer** — not the older PowerShell 5.1 that comes built into Windows. Check your version by opening PowerShell and running:
  ```powershell
  $PSVersionTable.PSVersion
  ```
  If the first number is less than 7, download PowerShell 7 from Microsoft and run this script with that instead.
- Network access from wherever you run this script to both your InsightVM console and NinjaOne.
- Admin access to InsightVM and to your NinjaOne console, for the one-time setup below.

## One-time setup

Do all of this once, before running the script for the first time.

### 1. Create an InsightVM account for this script

Use an existing InsightVM login, or create a dedicated one under **Administration > API Keys**. You'll need its username and password.

### 2. Create a NinjaOne API app

1. In NinjaOne, go to **Administration > Apps > API > Client App IDs > Add**.
2. Choose **API Services (machine-to-machine)** as the platform.
3. Under **Scopes**, check both **Monitoring** and **Management**.
4. Save it, then copy the **Client ID** and **Client Secret**. The Secret is only shown once — copy it somewhere safe immediately.

### 3. Enable the Rapid7 Vulnerability Importer in NinjaOne

Go to **Administration > Apps > Installed** (or **Add Apps** if you don't see it listed), find **Rapid7**, and click **Enable**.

### 4. Create the scan group

Inside the Rapid7 app in NinjaOne, open the **Scan Groups** tab and click **Create scan group**. Give it a name (e.g. `Rapid7 - All Servers`) and complete the setup wizard.

Note the **Scan Group ID** shown for it — you'll need that number in the script's config. This is the only manual, one-time step; after this, the script updates the same scan group via API on every run.

## Configuring the script

Open `Sync-InsightVM-To-NinjaOne.ps1` in a text editor (Notepad works fine). All the settings you need to change live in one clearly marked block near the top of the file, under:

```
YOUR SETTINGS - This is the ONLY part of the file you should need to change.
```

Fill in:

| Setting | What it is |
|---|---|
| `$InsightVMConsoleURL` | Your InsightVM console address, including the port (usually `:3780`) |
| `$InsightVMUsername` / `$InsightVMPassword` | The InsightVM login from setup step 1 |
| `$NinjaOneBaseURL` | Which NinjaOne cloud region your account is on (US, US2, EU, CA, or Oceania — pick the matching line in the file) |
| `$NinjaOneClientID` / `$NinjaOneClientSecret` | From the NinjaOne API app created in setup step 2 |
| `$NinjaOneScanGroupID` | The Scan Group ID number from setup step 4 |
| `$MinimumCVSSSeverity` | Only CVEs at or above this CVSS score get synced (e.g. `7.0`). Set to `0` to sync everything |
| `$DeviceIdField` | `"Hostname"` or `"IPAddress"` — must match how your NinjaOne scan group identifies devices. Most people should leave this as `"Hostname"` |
| `$LogFilePath` | Where the run log gets written. Defaults to `sync-log.txt` next to the script |

Save the file after editing. Everything below the "DO NOT EDIT" line handles the actual work and shouldn't need changes for normal use.

## Running it

From a PowerShell 7 prompt, in the folder where the script is saved:

```powershell
.\Sync-InsightVM-To-NinjaOne.ps1
```

The script will print progress to the screen and also write it to the log file. A successful run ends with a line confirming how many rows were sent to NinjaOne.

## Scheduling it to run automatically

Once a manual run works, you can set it up as a recurring scheduled task (Windows Task Scheduler, cron on Linux/macOS via `pwsh`, or your RMM's own script scheduler) so it stays in sync with your InsightVM scan cadence without you needing to run it by hand each time.

## Troubleshooting

The script tries to explain problems in plain language rather than raw error output. Common issues:

- **"COULD NOT CONNECT TO INSIGHTVM"** — check that the console URL is correct and reachable, and that the InsightVM username/password are correct.
- **"COULD NOT LOG IN TO NINJAONE"** — check that the region/base URL matches your account, and that the Client ID/Secret are correct and haven't been revoked.
- **"THE UPLOAD TO NINJAONE FAILED"** — check that the Scan Group ID actually exists, and that the NinjaOne API app has both the Monitoring and Management scopes enabled.
- **Device ID column looks empty in the CSV** — InsightVM's asset field names can vary slightly by version. If hostnames or IPs aren't showing up, that's the first place to check in the script.

## Security notes

- This script stores your InsightVM and NinjaOne credentials in plain text at the top of the file. Treat the saved script as sensitive: don't commit it to a public repository with real credentials filled in, don't email it, and restrict who can read it on disk.
- If you're publishing this repository, keep the script in its template form (with placeholder values) and have each user fill in their own copy locally rather than committing real credentials.

## What this script does not do

- It does not create the NinjaOne scan group for you — that's a one-time manual step (see setup step 4).
- It does not alert on new CVEs or send notifications — it only keeps the scan group's data current for whatever NinjaOne dashboards or workflows already consume it.
- It does not filter by asset group, site, or tag within InsightVM — it pulls from all assets the InsightVM account has access to.
