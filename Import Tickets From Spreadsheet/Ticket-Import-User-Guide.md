# NinjaOne Ticket Import Script — User Guide

A simple, step-by-step guide for migrating tickets into your new NinjaOne ticketing system.

> **About authentication:** NinjaOne requires the **Authorization Code** OAuth grant type for any operation that creates or modifies data (including creating tickets). On your first run, the script will open a browser window so you can log in to NinjaOne and approve the application. After that, it remembers your login automatically using an encrypted refresh token cache, so you won't be prompted again until the saved login expires.

---

## What this script does

This script takes a spreadsheet of tickets from your old ticketing system and creates them in your new NinjaOne ticketing system automatically. Instead of typing tickets in by hand, you fill out a spreadsheet and let the script do the work.

You do **NOT** need to be a programmer to use it. If you can edit a text file and run a command, you can use this script.

### What you will need

- A Windows computer with PowerShell (already included with Windows 10 and 11).
- An active NinjaOne account with administrator access.
- A web browser to handle the first-time login.
- A spreadsheet (`.xlsx` file) with your tickets in it.
- About 20 minutes the first time. After that, each import takes only a few minutes.

---

## Table of contents

1. [Part 1: One-time setup](#part-1-one-time-setup)
2. [Part 2: Create your API client app in NinjaOne](#part-2-create-your-api-client-app-in-ninjaone)
3. [Part 3: Find your NinjaOne BaseUrl](#part-3-find-your-ninjaone-baseurl)
4. [Part 4: Set up the script](#part-4-set-up-the-script)
5. [Part 5: Prepare your spreadsheet](#part-5-prepare-your-spreadsheet)
6. [Part 6: Do a test run (dry run)](#part-6-do-a-test-run-dry-run)
7. [Part 7: Do the real import](#part-7-do-the-real-import)
8. [How the saved login works](#how-the-saved-login-works)
9. [Troubleshooting](#troubleshooting)
10. [Quick reference card](#quick-reference-card)

---

## Part 1: One-time setup

You only need to do this part the first time you ever run the script on your computer.

### Step 1. Open PowerShell

1. Click the Windows Start button.
2. Type **PowerShell**.
3. Click **Windows PowerShell** in the search results.

A blue (or black) window with white text will open. This is PowerShell.

### Step 2. Install the two helper modules

The script uses two free PowerShell modules: **ImportExcel** (for reading spreadsheets) and **PSAuthClient** (for the NinjaOne login flow). Copy and paste these two lines into the PowerShell window, pressing **Enter** after each:

```powershell
Install-Module ImportExcel -Scope CurrentUser
Install-Module PSAuthClient -Scope CurrentUser
```

> ⚠️ **If you see a yellow message asking about an "untrusted repository":**
> Type `Y` and press **Enter** to continue. This is normal — it just means PowerShell is asking permission to download from the official module gallery.

Wait until each command finishes (you'll see the prompt return — the line ending with `>`). The modules are now installed for good — you will never need to do this step again.

---

## Part 2: Create your API client app in NinjaOne

Before the script can talk to NinjaOne, you need to create an **API Client App** inside NinjaOne. This gives the script permission to log in on your behalf.

> 💡 **You only do this once.** After the app is created, you reuse the same Client ID and Secret every time you run the script.

### Step 3. Log in to NinjaOne as an administrator

Open your web browser and go to your NinjaOne portal (whichever URL you normally use to log in). You must log in with an account that has administrator permissions.

### Step 4. Open the API settings page

1. In the left-hand menu, click **Administration** (the gear icon ⚙).
2. Under the **Apps** section, click **API**.
3. At the top of the API page, click the **Client App IDs** tab.

### Step 5. Add a new client app

1. Click the **Add** button (sometimes labeled **Add client app**).
2. The **Application Configuration** form opens.

Fill out the form using these exact values:

| Field | Value | Notes |
|---|---|---|
| **Application Platform** | `API Services (machine-to-machine)` | This is the platform NinjaOne expects for back-end integrations like this one. It supports the Authorization Code and Refresh Token grant types we need. |
| **Name** | `Ticket Import Script` | This is just a friendly label so you can find the app later. Pick anything you want. |
| **Redirect URIs** | `https://localhost` | Type it **exactly** like that. **Do NOT add a port number, slash, or path.** |
| **Scopes** | Check **Monitoring**, **Management**, AND **offline_access** | Management is required to create tickets. Monitoring is required to read account information. **offline_access** is what tells NinjaOne to give the script a refresh token — without it, you would have to log in through the browser every single time. |
| **Allowed Grant Types** | Check **Authorization Code** AND **Refresh Token** | Both are required. Authorization Code handles your initial login; Refresh Token lets the script log in silently on later runs. |

> ⚠️ **About the Redirect URI:**
> The redirect URI must match **EXACTLY** between NinjaOne and the script. The script uses `https://localhost` — no port, no slash, no path. If you type anything different here (for example `https://localhost:8080` or `https://localhost/`), the login will fail with a "redirect_uri_mismatch" error.

> 💡 **Why three scopes?**
> Without **offline_access**, NinjaOne will not issue a refresh token, and you would have to open the browser and log in every single time the script runs. With it, you log in once and the script remembers your login automatically.

### Step 6. Save the app and copy the Client Secret

1. Click **Save** at the top of the form.
2. NinjaOne will display the **Client Secret**. **This is the ONE AND ONLY time you will ever see it.** Copy it immediately and paste it somewhere safe (a password manager is best, or a temporary text file you can delete later).
3. Click **Close** to return to the Client App IDs list.

> 🚨 **If you lose the Client Secret, you cannot recover it.** You will have to delete the app and create a new one. Treat the secret like a password — keep it private.

### Step 7. Copy the Client ID

Back on the **Client App IDs** list, find the app you just created. The **Client ID** is shown next to it. Copy this value too — you'll paste both the Client ID and Client Secret into the script in Part 4.

You should now have these three pieces of information saved somewhere:

- ✅ **Client ID** (a long string of letters/numbers)
- ✅ **Client Secret** (another long string — this is the one you can't recover)
- ✅ The redirect URI you used: `https://localhost`

---

## Part 3: Find your NinjaOne BaseUrl

NinjaOne hosts customers in several different regions, and each region has its own URL. The script needs the right one or it will not be able to connect.

### Step 8. Find your region's URL

The easiest way to find your BaseUrl is to **look at the address bar of your browser when you are logged into NinjaOne.**

1. Log in to NinjaOne in your browser.
2. Look at the top of the browser window where the web address is shown.
3. Copy everything from `https://` up to (but **not including**) the first `/` after the domain name.

For example, if your browser shows:

```
https://app.ninjarmm.com/#/dashboard
```

Then your BaseUrl is:

```
https://app.ninjarmm.com
```

### Common NinjaOne regional URLs

If you are not sure, here are the standard NinjaOne regional URLs:

| Region | BaseUrl |
|---|---|
| United States (primary) | `https://app.ninjarmm.com` |
| United States (secondary) | `https://us2.ninjarmm.com` |
| Canada | `https://ca.ninjarmm.com` |
| Europe | `https://eu.ninjarmm.com` |
| Oceania (Australia) | `https://oc.ninjarmm.com` |

> 💡 **Some customers have a custom domain** (for example `https://yourcompany.rmmservice.eu`). If your portal lives on a custom domain, just use whatever shows in your browser when you log in.

---

## Part 4: Set up the script

### Step 9. Save the files in one folder

Create a folder somewhere easy to find, like your Desktop or Documents. Put these two files inside it:

- `Import-Tickets.ps1` — the script itself
- Your tickets spreadsheet, for example `MyTickets.xlsx`

A good example folder name: `C:\TicketMigration`

> 💡 The script will create a third file (`ninja-token-cache.xml`) automatically after your first login, in the same folder. That file is encrypted — leave it alone and don't delete it unless you want to be prompted to log in again.

### Step 10. Open the script in Notepad

1. Find `Import-Tickets.ps1` in your folder.
2. **Right-click** on it.
3. Choose **Open with → Notepad** (or use VS Code if you have it).

> ⚠️ **Important:** Do NOT double-click the file yet — that would run the script before you have finished setting it up. Always open it with Notepad first to edit.

### Step 11. Fill in your information

Near the top of the file you will see a section that looks like this:

```powershell
# === CONFIGURATION ===
$Config = @{
    BaseUrl       = 'https://<your NinjaOne URL>'
    ClientId      = '<Your Client ID>'
    ClientSecret  = '<Your Client Secret>'
    RedirectUri   = 'https://localhost'
    SpreadsheetPath = '<Full path to your .xlsx file>'
    WorksheetName = 'Tickets'
    LogFilePath   = '.\ticket-import-log.txt'
    TokenCachePath = '.\ninja-token-cache.xml'
    DryRun        = $true
}
```

Replace each value wrapped in angle brackets `< >` with your real information. **Keep the single quotes around each value.** Here is what to fill in for each field:

| Field | What to put in it | Where to find it |
|---|---|---|
| **BaseUrl** | Your NinjaOne portal URL — see Part 3. Example: `https://app.ninjarmm.com` | The address bar of your browser when logged into NinjaOne. |
| **ClientId** | The Client ID from the app you created in Part 2. | NinjaOne → Administration → Apps → API → Client App IDs. |
| **ClientSecret** | The Client Secret you saved when you created the app. | The value you copied immediately after saving the app. |
| **RedirectUri** | Leave this exactly as `https://localhost`. | Already filled in. Must match what you put in NinjaOne. |
| **SpreadsheetPath** | The full file path to your tickets spreadsheet. | Right-click your `.xlsx` file → **Properties** → check the Location, then add a backslash and the file name. |
| **WorksheetName** | The name of the tab inside the spreadsheet that has your tickets. | Open the spreadsheet — the tab name is at the bottom. The default is `Tickets`. |
| **LogFilePath** | Where to save the log file. The default is fine. | Default saves a file called `ticket-import-log.txt` in the same folder as the script. |
| **TokenCachePath** | Where to save the encrypted login cache. The default is fine. | Default saves an encrypted file called `ninja-token-cache.xml` in the same folder. |
| **DryRun** | Leave as `$true` for your first run. | Do **NOT** change this yet — it protects you from sending bad data. |

> 💡 **Tip — finding your spreadsheet path:**
> Hold **Shift**, right-click your `.xlsx` file in File Explorer, and choose **"Copy as path"**. Then paste it between the single quotes.
> Example: `'C:\TicketMigration\MyTickets.xlsx'`

### Step 12. Save the file

In Notepad, click **File → Save** (or press **Ctrl + S**). Close Notepad.

---

## Part 5: Prepare your spreadsheet

Your spreadsheet must have these columns in the first row (the header row). Column names are not case-sensitive, but they must be spelled exactly as shown.

| Column name | Required? | What goes in it |
|---|---|---|
| **TicketID** | Recommended | The original ticket number from your old system. **Must be numbers only** (e.g., `1001`, `45782`). No letters, dashes, or prefixes — any row with non-numeric TicketIDs will be skipped. |
| **Subject** | ✅ **YES** | A short summary of the ticket — what the user typed in the subject line. |
| **Description** | ✅ **YES** | The full description of the issue or request. |
| **Status** | Optional | `Open`, `In Progress`, `Resolved`, or `Closed`. |
| **Priority** | Optional | `Low`, `Medium`, `High`, or `Critical`. |
| **RequesterEmail** | ✅ **YES** | Email address of the person who reported the ticket. |
| **AssigneeEmail** | Optional | Email address of the person assigned to work the ticket. |
| **CreatedDate** | Optional | When the ticket was originally created. ISO format works best (`2024-11-03T09:14:00Z`). |
| **Category** | Optional | A category or queue name (e.g., `Hardware`, `Software`, `Network`). |
| **Tags** | Optional | Comma-separated tags (e.g., `printer,office`). |

> ⚠️ **Required fields:** `Subject`, `Description`, and `RequesterEmail` are required for every ticket. Rows missing any of these will be skipped (but the rest of your import will still run).

> ⚠️ **TicketID must be numeric:** NinjaOne stores ticket IDs as numbers. If your old ticket numbers had letters or prefixes (like `LEG-1001` or `INC0042`), strip the letters and dashes before importing — leave only the digits (`1001`, `0042`). Rows with non-numeric TicketIDs will be skipped during validation.

---

## Part 6: Do a test run (dry run)

A **dry run** means the script reads your spreadsheet and pretends to import it, but does **NOT** actually send anything to NinjaOne. This lets you catch mistakes before they cause real problems.

### Step 13. Run the script in PowerShell

1. Open PowerShell (Start menu → type **PowerShell** → Enter).
2. Navigate to your folder by typing `cd` followed by the folder path. For example:

   ```powershell
   cd C:\TicketMigration
   ```

3. Run the script by typing this and pressing **Enter**:

   ```powershell
   .\Import-Tickets.ps1
   ```

> 🌐 **First run only — your browser will open.**
> Because this is your first time running the script, a browser window will pop up showing the NinjaOne login page. Sign in normally and click **Allow** (or **Approve**) when NinjaOne asks if your "Ticket Import Script" application can access your account. The browser will then redirect to a `https://localhost` page that may show a "site can't be reached" or similar message — **that's fine**. The script captures what it needs from the URL and continues automatically. You can close that browser tab.

> ⚠️ **If you see a red error about "running scripts is disabled":**
> This is a Windows security setting. Run this command once, then try the script again:
>
> ```powershell
> Set-ExecutionPolicy -Scope CurrentUser -ExecutionPolicy RemoteSigned
> ```
>
> Type `Y` and press **Enter** when prompted.

### Step 14. Read the results

When the script finishes you will see a summary at the bottom that looks something like this:

```
============================================================
                    IMPORT SUMMARY
============================================================
DRY RUN — no data was actually sent to the API.
Total rows read:        15
Imported successfully:  13
Skipped (validation):   2
Failed (API errors):    0

Skipped rows:
  Row 14: Missing required field(s): Subject
  Row 15: Missing required field(s): RequesterEmail
```

What the colors mean:

- 🟢 **Green** = success — the row is good to go.
- 🟡 **Yellow** = skipped — the row is missing required information. Fix it in the spreadsheet.
- 🔴 **Red** = error — the script could not finish. Read the message and fix the cause.
- 🔵 **Cyan** = informational progress messages.

### Step 15. Fix any problems

If any rows were skipped, open your spreadsheet and fix them. The summary tells you the row number and what is missing.

If you got a configuration error (something in red about a field still containing `<` or `>`), go back to **Step 11** and make sure you replaced ALL the placeholder values.

Save the spreadsheet and run the dry run again until the summary shows zero skipped rows (or only rows you intentionally want skipped).

> ✅ **Good news:** Subsequent runs will NOT open the browser. Your login is now saved (encrypted) in `ninja-token-cache.xml`.

---

## Part 7: Do the real import

Once your dry run looks good, you are ready to do the actual import.

### Step 16. Turn off DryRun

1. Open `Import-Tickets.ps1` in Notepad again.
2. Find this line in the configuration block:

   ```powershell
   DryRun        = $true
   ```

3. Change `$true` to `$false`:

   ```powershell
   DryRun        = $false
   ```

4. Save and close the file.

### Step 17. Run the import for real

Back in PowerShell (still in your folder), run the script again:

```powershell
.\Import-Tickets.ps1
```

This time the script will use your saved login (no browser needed), authenticate with NinjaOne, and create each ticket. You will see green `SUCCESS` lines as each ticket is imported. The summary at the bottom will show how many succeeded.

> ✅ **You did it!** Your tickets are now in NinjaOne. Open the ticketing portal in your browser and confirm a few of them look right.

---

## How the saved login works

After your first successful login, the script saves a NinjaOne **refresh token** to an encrypted file (default: `.\ninja-token-cache.xml`).

- The encryption uses Windows DPAPI tied to your current Windows user account.
- This means **only the same Windows user on the same computer can decrypt the file**. Even another user on the same machine cannot read it.
- The script uses this saved token to log in silently on every future run, so you won't see the browser prompt again.
- Refresh tokens eventually expire (typically several months). When that happens, the script will automatically open the browser and ask you to log in again.
- If you delete the cache file, run the script as a different Windows user, or move it to a different computer, you'll be prompted to log in again — that is expected and safe.

> 🔒 **The cache file is encrypted, but treat it as sensitive anyway.** Don't email it or upload it to shared storage. If you suspect it has been exposed, delete it and revoke/rotate the API client app's secret in NinjaOne.

---

## Troubleshooting

Below are the most common issues and how to fix them.

### "The following configuration field(s) still contain placeholder values"

You forgot to replace one of the values in the configuration block. The error message tells you which field. Open the script in Notepad, find that field, and replace the value between the angle brackets.

### "Cannot find the spreadsheet file at..."

The path in `SpreadsheetPath` is wrong, or the file has been moved. Right-click the spreadsheet, choose **Properties**, and check the Location. Make sure `SpreadsheetPath` has the full path including the file name and `.xlsx` extension.

### "The required PowerShell module 'ImportExcel' is not installed" (or 'PSAuthClient')

You skipped Step 2. Open PowerShell and run whichever line is missing:

```powershell
Install-Module ImportExcel -Scope CurrentUser
Install-Module PSAuthClient -Scope CurrentUser
```

### "redirect_uri_mismatch" or "Invalid redirect URI"

The redirect URI in the script does not match the one you set in the NinjaOne API client app. Both must be **exactly** `https://localhost` — no port number, no trailing slash, no path. Open the API client app in NinjaOne and confirm the redirect URI is exactly `https://localhost`.

### Browser opened but the page just says "This site can't be reached" / "ERR_CONNECTION_REFUSED"

That's actually expected. NinjaOne redirects your browser to `https://localhost` after login, and there's no real web server there to respond — but the script captures the authorization code from the URL before that error appears. Just close the tab and look back at PowerShell, where the script should have continued.

If the script did NOT continue and is still waiting, the redirect URI on the NinjaOne app probably doesn't match. See the previous troubleshooting entry.

### "401 Unauthorized" during login

One of your credentials is wrong, or the API client app is configured incorrectly. Check these in order:

1. Open the script and double-check `ClientId` and `ClientSecret`. If you copied and pasted, make sure no extra spaces snuck in at the start or end.
2. Make sure your `BaseUrl` matches your region (Part 3). A US user pointing at `eu.ninjarmm.com` will fail.
3. Verify the API client app in NinjaOne has **Authorization Code** AND **Refresh Token** checked under Allowed Grant Types, and **Monitoring**, **Management**, AND **offline_access** checked under Scopes.
4. If the Client Secret was lost or you're not sure it's correct, delete the app in NinjaOne and create a new one (Part 2). The Client Secret is only displayed once at creation time.

### "invalid_grant" or "Your saved login has expired"

This is normal — refresh tokens expire after a few months. The script will automatically open the browser to let you log in again. Just sign in and approve the application.

### "403 Forbidden" when creating a ticket

Your API client app is missing the **Management** scope. Go back to NinjaOne → Administration → Apps → API → Client App IDs, edit the app, check **Management** under Scopes, and save. Then delete `ninja-token-cache.xml` and run the script again.

### "Cannot connect to host" or DNS errors

Your `BaseUrl` is probably wrong for your region. Log in to NinjaOne in your browser and check the address bar. Update the `BaseUrl` in the script to match.

### "running scripts is disabled on this system"

Windows is blocking PowerShell scripts by default. Run this command once, then try again:

```powershell
Set-ExecutionPolicy -Scope CurrentUser -ExecutionPolicy RemoteSigned
```

### Some rows say "Skipped"

Those rows are missing one of the three required fields (`Subject`, `Description`, or `RequesterEmail`), or the TicketID contains letters or other non-numeric characters. Find the row in your spreadsheet, fix the issue, and run the script again.

### "TicketID must contain only digits" or skipped rows due to ticket ID format

Your spreadsheet has TicketIDs with letters, dashes, or other non-numeric characters (like `LEG-1001` or `INC0042`). NinjaOne requires numeric ticket IDs. Open the spreadsheet and either:

- Strip the letters and dashes so only the number remains (`LEG-1001` → `1001`), or
- Clear the TicketID column entirely if you don't need the reference (the import will still work).

Save and re-run.

### Some rows say "Failed"

The API rejected those tickets. The reason is in the summary — common ones include:

- **400 Bad Request** — a field value is invalid (like a Status that NinjaOne does not recognize).
- **403 Forbidden** — your API client app is missing the **Management** scope. Edit the app to enable it.
- **409 Conflict** — a ticket with that ID already exists.

Fix the row in the spreadsheet (or the app's scopes) and re-run.

### I want to force the script to ask me to log in again

Delete the file `ninja-token-cache.xml` from the script folder. The next run will open the browser as if it were the first time.

### Where is the log file?

Look in the same folder as the script for `ticket-import-log.txt`. It contains a complete record of every action the script took, with timestamps. Send this file to whoever shared the script with you if you need help.

---

## Quick reference card

### Every-time checklist

1. Open PowerShell.
2. `cd` to your script folder.
3. Edit `Import-Tickets.ps1` if anything has changed (spreadsheet path, etc.).
4. Run with `DryRun = $true` first.
5. Review the summary, fix any skipped rows.
6. Set `DryRun = $false`.
7. Run again for the real import.

### NinjaOne API client app — required settings

| Setting | Value |
|---|---|
| Application Platform | `API Services (machine-to-machine)` |
| Redirect URI | `https://localhost` |
| Scopes | `Monitoring` + `Management` + `offline_access` |
| Allowed Grant Types | `Authorization Code` + `Refresh Token` |

### Commands you will use

| What it does | Command |
|---|---|
| Install ImportExcel module (one time) | `Install-Module ImportExcel -Scope CurrentUser` |
| Install PSAuthClient module (one time) | `Install-Module PSAuthClient -Scope CurrentUser` |
| Allow scripts to run (one time, if needed) | `Set-ExecutionPolicy -Scope CurrentUser -ExecutionPolicy RemoteSigned` |
| Move into your script folder | `cd C:\TicketMigration` |
| Run the import script | `.\Import-Tickets.ps1` |
| Force a fresh browser login | Delete `ninja-token-cache.xml`, then run the script |

---

*Need more help? Check the log file (`ticket-import-log.txt`) and contact whoever shared this script with you.*
