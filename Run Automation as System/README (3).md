# NinjaOne — Run a Script as SYSTEM via PowerShell

This script lets you run a PowerShell command on a remote device through NinjaOne, running as the **SYSTEM** account — the highest-privilege local account on a Windows machine. You do not need to be logged into the device, and no agent interaction is required.

---

## What You Will Need

Before you start, make sure you have the following:

- A **NinjaOne account** with administrator access
- **Windows PowerShell** (already installed on any modern Windows PC — no download needed)
- The `.ps1` script file from this repository downloaded to your computer

---

## Step 1 — Create an API App in NinjaOne

This step gives the script permission to talk to NinjaOne on your behalf. You only need to do this once.

1. Log in to your NinjaOne portal.
2. In the left sidebar, click **Administration**.
3. Go to **Apps** → **API** → **Client App**.
4. Click **Add** to create a new app (or open an existing one if you already have one set up).
5. Fill in the following:
   - **Name:** Give it any name you like, e.g. `Run Script Tool`
   - **Redirect URI:** Enter exactly the following — do not add anything extra:
     ```
     https://localhost
     ```
6. Save the app. NinjaOne will display a **Client ID** and **Client Secret** — copy both and keep them somewhere safe. You will need them in the next step.

> **Note:** The Client Secret is only shown once. If you lose it, you will need to generate a new one.

---

## Step 2 — Find Your Device ID

The script needs to know which device to run the command on. NinjaOne identifies devices by a number called the **Device ID**.

1. In NinjaOne, go to **Devices** and click on the device you want to target.
2. Look at the URL in your browser's address bar. It will look something like this:
   ```
   https://app.ninjarmm.com/#/deviceDashboard/12345/overview
   ```
3. The number near the end of the URL (`12345` in the example above) is your **Device ID**. Write it down.

---

## Step 3 — Edit the Script

Open the `.ps1` script file in a text editor. **Notepad works fine** — right-click the file and choose **Edit** or **Open with → Notepad**.

Near the top of the file you will see a section that looks like this:

```powershell
$BaseUrl         = 'https://<your login URL>'
$TokenEndpoint   = 'https://<your login URL>/ws/oauth/token'
$ClientId        = '<Your Client ID>'
$ClientSecret    = '<Your Client Secret>'

$DeviceId        = '<Your Device ID>'
```

Replace each placeholder with your real values:

| Placeholder | What to put here |
|---|---|
| `https://<your login URL>` | The URL you use to log in to NinjaOne, e.g. `https://app.ninjarmm.com` |
| `https://<your login URL>/ws/oauth/token` | Same URL as above, with `/ws/oauth/token` added to the end |
| `<Your Client ID>` | The Client ID from Step 1 |
| `<Your Client Secret>` | The Client Secret from Step 1 |
| `<Your Device ID>` | The number you found in Step 2 |

**Example of a filled-in config block:**

```powershell
$BaseUrl         = 'https://app.ninjarmm.com'
$TokenEndpoint   = 'https://app.ninjarmm.com/ws/oauth/token'
$ClientId        = 'abc123def456'
$ClientSecret    = 'supersecretvalue'

$DeviceId        = '12345'
```

---

## Step 4 — Enter the Commands You Want to Run

Still in the script, find the `$ScriptBody` block just below the config section:

```powershell
$ScriptBody      = @'
Write-Output "Hello from SYSTEM context"
whoami
'@
```

Replace the example commands between `@'` and `'@` with whatever PowerShell commands you want to run on the device. The commands will run as SYSTEM.

**Do not delete or modify the `@'` and `'@` lines themselves** — they are just markers that tell PowerShell where the script starts and ends.

Save the file when you are done.

---

## Step 5 — Run the Script

1. Open **PowerShell** on your computer.
   - Press the **Windows key**, type `PowerShell`, and press **Enter**.
2. Navigate to the folder where you saved the script. For example, if it is in your Downloads folder:
   ```powershell
   cd C:\Users\YourName\Downloads
   ```
3. Run the script:
   ```powershell
   .\NinjaOne-RunScriptAsSystem.ps1
   ```

> **If you see an error about "running scripts is disabled"**, run this command first and then try again:
> ```powershell
> Set-ExecutionPolicy -Scope Process -ExecutionPolicy Bypass
> ```

---

## Step 6 — Log In and Authorize

The script will automatically open your **browser** and take you to a NinjaOne login page.

1. Log in with your NinjaOne credentials.
2. Click **Authorize** (or **Allow**) when prompted.
3. Your browser will go to a **blank page** or show a "connection refused" error — **this is completely normal**. It means the login worked.
4. Look at the **address bar** at the top of your browser. The URL will look something like this:
   ```
   https://localhost/?code=abc123xyz&state=somethinglong
   ```
5. **Copy that entire URL** from the address bar.
6. Go back to the PowerShell window. It will be waiting and will show:
   ```
   Paste URL here:
   ```
7. Paste the URL you copied and press **Enter**.

---

## Step 7 — Wait for the Result

The script will now connect to NinjaOne, send your commands to the device, and wait for them to finish. You will see progress messages in the PowerShell window.

When complete, the output will look something like this:

```
===============================================================
 Script completed with status: SUCCESS
===============================================================

--- Output ---
nt authority\system
```

If the script takes longer than 2 minutes, a yellow message will appear with an **Activity ID** you can use to check the result manually in NinjaOne under the device's activity log.

---

## Troubleshooting

**"The URL does not look right (state mismatch)"**
You may have copied the URL from the wrong tab, or the page refreshed. Re-run the script and try again, making sure to copy the URL immediately after the blank page appears.

**"No authorization code found in the URL"**
The URL you pasted does not contain a `code=` parameter. Make sure you copied the full URL from the address bar, not just part of it.

**"Could not get an access token"**
Double-check your `$ClientId`, `$ClientSecret`, and `$TokenEndpoint` values in the script. Make sure there are no extra spaces or quotes around the values.

**"Failed to send the script"**
Make sure the `$DeviceId` is correct and that the target device is currently online in NinjaOne.

**The browser did not open automatically**
Copy the URL that appears in the PowerShell window and paste it into your browser manually.

---

## Security Notes

- Your **Client Secret** is sensitive — treat it like a password. Do not share the script with the secret already filled in.
- The script runs commands as **SYSTEM**, which has full control over the local machine. Double-check your `$ScriptBody` commands before running.
- Access tokens obtained by this script are short-lived and are not saved anywhere.
