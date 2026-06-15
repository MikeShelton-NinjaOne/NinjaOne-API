# NinjaOne — Device Detail Export (Expanded Fields)

This script connects to NinjaOne and exports a full CSV of every device in your environment, including the **expanded** Organization, Location, and Warranty fields that are not included in a standard device list pull.

After the first time you log in, the script saves a refresh token so it can renew its own access on every future run — no browser login needed after the first time.

---

## What Gets Exported

Each row in the CSV represents one device. The columns are grouped into four sections:

### Core Device Info
| Column | Description |
|---|---|
| Device ID | NinjaOne's internal numeric ID for the device |
| UID | The device's unique identifier (UUID format) |
| Display Name | The name shown in the NinjaOne interface |
| System Name | The actual computer/hostname name |
| DNS Name | The fully qualified domain name, if available |
| Node Class | The type of device (e.g. WINDOWS_WORKSTATION, MAC, LINUX_SERVER) |
| Node Role ID | The ID of the role assigned to the device in NinjaOne |
| Online | Whether the device is currently online (True/False) |
| Last Contact | The date and time NinjaOne last heard from the device |
| Created | The date and time the device was first added to NinjaOne |
| IP Addresses | All known local IP addresses |
| MAC Addresses | All known MAC addresses |
| Public IP | The device's public-facing IP address |
| Agent Version | The version of the NinjaOne agent installed on the device |

### System & OS Info
| Column | Description |
|---|---|
| Manufacturer | Hardware manufacturer (e.g. Dell, HP, Lenovo) |
| Model | Hardware model name |
| BIOS Serial Number | The serial number from the device's BIOS |
| OS Name | Operating system name (e.g. Windows 11 Pro) |
| OS Build Number | OS build number |
| OS Version | Full OS version string |

### Organization (Expanded)
| Column | Description |
|---|---|
| Organization ID | NinjaOne's internal ID for the organization |
| Organization Name | The name of the organization the device belongs to |
| Organization Description | The description set on the organization, if any |
| Organization Website | The website set on the organization, if any |

### Location (Expanded)
| Column | Description |
|---|---|
| Location ID | NinjaOne's internal ID for the location |
| Location Name | The name of the location the device is assigned to |
| Location Address | Street address of the location |
| Location City | City |
| Location State | State or province |
| Location Zip | Zip or postal code |
| Location Country | Country |

### Warranty (Expanded)
| Column | Description |
|---|---|
| Warranty Start Date | The date the warranty began |
| Warranty End Date | The date the warranty expires |
| Warranty Mfr Fulfillment Date | The date the manufacturer fulfilled/registered the warranty |

> **Note:** Warranty fields will only be populated for devices where warranty tracking has been set up in NinjaOne. Devices without warranty data will have blank warranty columns.

---

## What You Will Need

- A **NinjaOne account** with administrator access
- **Windows PowerShell** (already installed on any modern Windows PC)
- The `.ps1` script file downloaded to your computer

---

## Step 1 — Create an API App in NinjaOne

You only need to do this once.

1. Log in to your NinjaOne portal.
2. In the left sidebar click **Administration**.
3. Go to **Apps** → **API** → **Client App**.
4. Click **Add** to create a new app (or open an existing one).
5. Set the following:
   - **Name:** Any name you like, e.g. `Device Export Tool`
   - **Redirect URI:** Enter exactly the following — nothing extra:
     ```
     https://localhost
     ```
   - **Scopes / Allowed Grants:** Make sure **Refresh Token** (sometimes labeled `offline_access`) is enabled.
6. Save the app. Copy the **Client ID** and **Client Secret** — keep them somewhere safe.

> **Note:** The Client Secret is only shown once. If you lose it, you will need to generate a new one.

---

## Step 2 — Edit the Script

Open the `.ps1` file in a text editor. **Notepad works fine** — right-click the file and choose **Edit** or **Open with → Notepad**.

Find the configuration section near the top:

```powershell
$BaseUrl         = 'https://<your login URL>'
$TokenEndpoint   = 'https://<your login URL>/ws/oauth/token'
$ClientId        = '<Your Client ID>'
$ClientSecret    = '<Your Client Secret>'
```

Replace each placeholder with your real values:

| Placeholder | What to put here |
|---|---|
| `https://<your login URL>` | The URL you use to log in to NinjaOne, e.g. `https://app.ninjarmm.com` |
| `https://<your login URL>/ws/oauth/token` | Same URL with `/ws/oauth/token` added to the end |
| `<Your Client ID>` | The Client ID from Step 1 |
| `<Your Client Secret>` | The Client Secret from Step 1 |

**Example of a filled-in config block:**

```powershell
$BaseUrl         = 'https://app.ninjarmm.com'
$TokenEndpoint   = 'https://app.ninjarmm.com/ws/oauth/token'
$ClientId        = 'abc123def456'
$ClientSecret    = 'supersecretvalue'
```

The `$CsvOutputPath` and `$TokenFile` lines below the config control where files are saved. The defaults put everything in the same folder as the script, which works for most people. You can change those paths if you prefer a different location.

Save the file when done.

---

## Step 3 — Run the Script

1. Open **PowerShell** on your computer.
   - Press the **Windows key**, type `PowerShell`, and press **Enter**.
2. Navigate to the folder where you saved the script:
   ```powershell
   cd C:\Users\YourName\Downloads
   ```
3. Run the script:
   ```powershell
   .\NinjaOne-DeviceExport-Expanded.ps1
   ```

> **If you see an error about "running scripts is disabled"**, run this first and then try again:
> ```powershell
> Set-ExecutionPolicy -Scope Process -ExecutionPolicy Bypass
> ```

---

## First Run — Log In and Authorize

The first time you run the script, it will open your browser.

1. Log in with your NinjaOne credentials and click **Authorize**.
2. Your browser will land on a **blank page** — this is normal.
3. Copy the **entire URL** from the browser address bar. It will look like:
   ```
   https://localhost/?code=abc123xyz&state=somethinglong
   ```
4. Go back to the PowerShell window, paste the URL when prompted, and press **Enter**.

The script will log in, save a refresh token, pull all your devices, and export the CSV.

---

## Every Run After That

No browser login needed. The script will print:

```
Found saved refresh token. Renewing access...
Access token renewed successfully.
```

Then it goes straight to pulling devices and building the CSV.

---

## What Gets Created

After the script runs you will find two files in the same folder as the script:

| File | What it is |
|---|---|
| `NinjaOne-Devices-YYYY-MM-DD.csv` | The device export, named with today's date |
| `ninja_refresh_token.txt` | The saved login token (keep this safe — see Security Notes) |

A new dated CSV is created each time you run the script, so previous exports are not overwritten.

---

## Large Environments

The script automatically handles environments with more than 1,000 devices by paging through the results. You will see a running count in the PowerShell window as it pulls each page:

```
Retrieved 1000 devices so far...
Retrieved 2000 devices so far...
Total devices retrieved: 2000
```

No configuration is needed for this — it works automatically regardless of how many devices you have.

---

## Troubleshooting

**"Found saved refresh token" but then it asks me to log in again**
The refresh token has expired. Log in again and a new one will be saved automatically.

**"WARNING: No refresh token returned"**
The `offline_access` / Refresh Token grant is not enabled on your API app. Go to **Administration → Apps → API → Client App**, open your app, and enable it.

**"Failed to retrieve devices"**
Check that your `$BaseUrl` is correct, that you are connected to the internet or VPN if required, and that your NinjaOne account has permission to view devices.

**Warranty columns are all blank**
Warranty tracking must be set up in NinjaOne before data appears here. Go to **Administration → Endpoint Management → Warranty** to enable it for your device manufacturers.

**"The URL does not look right (state mismatch)"**
Re-run the script and copy the URL immediately after the blank page appears, without navigating away first.

**"No authorization code found in the URL"**
Make sure you copied the full URL, including everything after the `?`.

---

## Security Notes

- **`ninja_refresh_token.txt`** grants API access to your NinjaOne account — treat it like a password. Do not share it or commit it to source control.
- Do not share the script with the `$ClientId` and `$ClientSecret` already filled in.
- Add these to your `.gitignore` if you are storing this in a Git repository:
  ```
  ninja_refresh_token.txt
  NinjaOne-Devices-*.csv
  ```
- To fully revoke access, delete `ninja_refresh_token.txt` **and** regenerate the Client Secret in NinjaOne under **Administration → Apps → API**.
