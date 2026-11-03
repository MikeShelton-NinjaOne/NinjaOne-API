# NinjaOne CSV Custom Field Import

A PowerShell script that reads a CSV file and writes its data into a **device custom field** in NinjaOne using the NinjaOne Public API. Devices are matched by their NinjaOne **Device ID**.

- **Script:** `Simple-NinjaImport.ps1`
- **Auth method:** OAuth 2.0 Client Credentials (no browser sign-in)
- **Field type tested:** Multi-line text

---

## Contents

1. [What the script does](#what-the-script-does)
2. [Requirements](#requirements)
3. [Step 1: Create the NinjaOne API app](#step-1-create-the-ninjaone-api-app)
4. [Step 2: Prepare the custom field](#step-2-prepare-the-custom-field)
5. [Step 3: Prepare the CSV](#step-3-prepare-the-csv)
6. [Step 4: Configure the script](#step-4-configure-the-script)
7. [Step 5: Run the script](#step-5-run-the-script)
8. [Understanding the results](#understanding-the-results)
9. [Troubleshooting](#troubleshooting)
10. [Security notes](#security-notes)

---

## What the script does

1. Reads your CSV file.
2. Logs in to the NinjaOne API using your API app's Client ID and Client Secret.
3. Retrieves your device list and matches each CSV row to a device by Device ID.
4. Writes the mapped CSV values to the chosen device custom field(s).
5. Prints a color-coded result for each row and saves a results CSV next to your import file.

Behavior to be aware of:

- **Empty cells are skipped.** Existing values in Ninja are not cleared.
- **Existing values are overwritten** when the CSV cell has data.
- **Line breaks** in cells are normalized so they display correctly in multi-line fields.
- **There is no dry run.** The script writes to NinjaOne every time it runs. Always test with a small CSV first.

---

## Requirements

- Windows with PowerShell 5.1 or later
- A NinjaOne account with permission to create API apps and custom fields
- Network access to your NinjaOne instance (for example `app.ninjarmm.com`)
- A CSV file containing NinjaOne Device IDs and the data to import

---

## Step 1: Create the NinjaOne API app

In NinjaOne, go to **Administration > Apps > API > Client app IDs** and click **Add**. Menu wording can vary slightly between NinjaOne versions.

| Setting | Value |
|---|---|
| Application platform | **API Services (machine-to-machine)** |
| Allowed grant types | **Client Credentials** only |
| Allowed scopes | **Monitoring** and **Management** |
| Redirect URI | Not needed. Leave blank (if a value is required, a placeholder such as `https://localhost` is fine) |

Save the app, then copy the **Client ID** and **Client Secret**. The secret is typically shown only once, so store it in a password manager.

> Authorization Code flow is **not** used by this script. If the app is created as a web app type, the login will fail.

---

## Step 2: Prepare the custom field

Go to **Administration > Devices** and open **Global Custom Fields** or **Role Custom Fields**.

1. Create or open the custom field you want to fill. Type should be **Multi-line text**.
2. Set the field's **API permission** to **Write** (or Read/Write). Without this, saves will fail.
3. Note the field's **Name** (the machine name, no spaces). This is **not** the display Label.

If you use a **Role** custom field, only devices assigned to that role can receive the value. A **Global** custom field applies to all devices.

Multi-line text fields have a length limit of roughly 10,000 characters. Confirm the current limit in the NinjaOne documentation for your tenant.

---

## Step 3: Prepare the CSV

The CSV needs:

- A column containing the numeric **NinjaOne Device ID**
- One column for each custom field you want to fill

Example:

```csv
DeviceID,Notes
1234,"First line
Second line"
5678,Another note
```

**Finding a Device ID:** open a device in NinjaOne and look at the URL. The number in it is the Device ID, for example `.../#/deviceDashboard/1234/overview`.

Tips:

- Wrap any cell that contains line breaks or commas in double quotes.
- If you edit the CSV in Excel, make sure the Device ID column isn't reformatted as a decimal or scientific notation.
- Save as **CSV (Comma delimited)**.

---

## Step 4: Configure the script

Open `Simple-NinjaImport.ps1` in a text editor and edit the lines marked `# <<< EDIT` in the **SETTINGS** section.

| # | Setting | What to enter | Example |
|---|---|---|---|
| 1 | `$CsvPath` | Full path to your CSV file | `"C:\Temp\import.csv"` |
| 2 | `$DeviceIdColumn` | CSV column header that holds the Device IDs | `"DeviceID"` |
| 3 | `$FieldMap` | CSV column mapped to the custom field machine name | `"Notes" = "deviceNotes"` |
| 4 | `$ClientId` | Client ID from your API app | `"AbC123xyz"` |
| 5 | `$Region` | NinjaOne region: `app`, `eu`, `oc`, `ca`, `us2`, or `fedgov` | `"app"` |

### Field mapping

`$FieldMap` uses the format `"CSV Column Header" = "ninjaFieldMachineName"`. Add one line per field:

```powershell
$FieldMap = @{
    "Notes"     = "deviceNotes"
    "Asset Tag" = "assetTag"
}
```

- Left side: the CSV header, exactly as written (including capitalization and spaces).
- Right side: the custom field's machine name from NinjaOne.

### Region

Set `$Region` to match the address you sign in to NinjaOne with:

| Sign-in address | Region value |
|---|---|
| `app.ninjarmm.com` | `app` |
| `us2.ninjarmm.com` | `us2` |
| `eu.ninjarmm.com` | `eu` |
| `ca.ninjarmm.com` | `ca` |
| `oc.ninjarmm.com` | `oc` |
| `fedgov.rmmservice.com` | `fedgov` |

API apps created on one instance do not work on another.

---

## Step 5: Run the script

1. **Test first.** Create a small CSV with two or three rows and point `$CsvPath` at it.
2. Right-click `Simple-NinjaImport.ps1` and choose **Run with PowerShell**, or run it from a PowerShell window:

   ```powershell
   powershell.exe -File ".\Simple-NinjaImport.ps1"
   ```

3. Enter your **Client Secret** when prompted. The input is hidden and the secret is never saved.
4. Check the results on screen, then open one of the test devices in NinjaOne and confirm the custom field contains the expected text.
5. Point `$CsvPath` at the full CSV and run again.

Running the script again is safe. It overwrites the field with the CSV value and does not create duplicates.

### If Windows blocks the script

If you see "running scripts is disabled on this system", use one of these:

- **Unblock the file:** right-click the file > **Properties** > tick **Unblock** > OK.
- **Allow scripts for your user:**

  ```powershell
  Set-ExecutionPolicy -Scope CurrentUser RemoteSigned
  ```

- **Bypass for one run only:**

  ```powershell
  powershell.exe -ExecutionPolicy Bypass -File ".\Simple-NinjaImport.ps1"
  ```

If your organization enforces the policy through Group Policy, you may need IT to allow or sign the script.

---

## Understanding the results

Each row is shown with a status, the Device ID, the device name from Ninja, and a note.

| Status | Color | Meaning |
|---|---|---|
| `UPDATED` | Green | The custom field was written successfully |
| `SKIPPED` | Gray | The row had no Device ID or all mapped cells were empty |
| `NOT FOUND` | Yellow | No device with that ID exists in Ninja |
| `ERROR` | Red | The API rejected the update (the error text is shown in the note) |

A results file named `NinjaImportResults_<date>_<time>.csv` is saved in the same folder as your input CSV.

---

## Troubleshooting

| Problem | Likely cause and fix |
|---|---|
| **"Could not log in"** | Check the Client ID, Client Secret, and Region. Read the yellow `Details from NinjaOne` line for the exact reason. |
| `invalid_client` | Client ID or Secret doesn't match, or the Region is wrong. Re-copy the credentials, or generate a new secret. |
| `unsupported_grant_type` / `unauthorized_client` | **Client Credentials** isn't enabled on the API app. |
| `invalid_scope` | **Monitoring** and **Management** scopes aren't both enabled on the API app. |
| **400 Bad Request at login** | Usually an API app setup problem (grant type, scopes, or platform type). Confirm the settings in [Step 1](#step-1-create-the-ninjaone-api-app). |
| **"Column '...' was not found in the CSV"** | A name in `$DeviceIdColumn` or `$FieldMap` doesn't match your CSV header. The script lists the columns it found. |
| **`NOT FOUND` rows** | The Device ID isn't in Ninja, or Excel changed the number format. Verify the IDs. |
| **`ERROR` with a 403 or permission message** | The custom field's API permission isn't set to Write, or the API app is missing the Management scope. |
| **`ERROR` with a 400 message on a row** | The custom field machine name in `$FieldMap` is wrong, or the field doesn't apply to that device (for example, a Role field on a device with a different role). |
| **"Running scripts is disabled"** | See [If Windows blocks the script](#if-windows-blocks-the-script). |

---

## Security notes

- **Never commit your Client Secret** to Git. The script prompts for it at run time specifically so it isn't stored in the file.
- Your Client ID is less sensitive but should still not be shared publicly. Consider leaving the placeholder value in the repository copy.
- Don't commit your import CSVs or results files if they contain internal data. Consider adding this to `.gitignore`:

  ```
  *.csv
  NinjaImportResults_*.csv
  ```

- The API app acts as itself, so changes in NinjaOne's audit trail are attributed to the API app's name. Give it a clear name such as "CSV Custom Field Import".
- If the secret is ever exposed, generate a new one in the API app and update wherever it's stored.

---

## API endpoints used

| Purpose | Method and path |
|---|---|
| Log in | `POST /ws/oauth/token` |
| Get devices | `GET /v2/devices` |
| Update custom fields | `PATCH /v2/device/{id}/custom-fields` |

Reference: NinjaOne Public API documentation for your region, for example `https://app.ninjarmm.com/apidocs`.
