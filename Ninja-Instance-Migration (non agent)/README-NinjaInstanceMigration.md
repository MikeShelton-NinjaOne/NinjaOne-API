# NinjaOne Instance Config Migration

> Two PowerShell scripts that export automation scripts, custom fields, device roles, and device policy shells from one NinjaOne instance and import them into another. Custom fields, roles, and policy shells are fully automated. Automation scripts and policy conditions require manual steps — the scripts generate a complete checklist so nothing gets missed.

---

## The Two Scripts

| Script | Purpose |
|---|---|
| `Export-NinjaInstanceConfig.ps1` | Pulls config from the **source** instance and saves a JSON package |
| `Import-NinjaInstanceConfig.ps1` | Reads the JSON and creates matching config in the **destination** instance |

Run Export first, then Import.

---

## Table of Contents

- [What Is and Isn't Automated](#what-is-and-isnt-automated)
- [Critical Prerequisite — Custom Field API Permissions](#critical-prerequisite--custom-field-api-permissions)
- [Prerequisites](#prerequisites)
- [Part 1: API App Setup — Source Instance](#part-1-api-app-setup--source-instance)
- [Part 2: API App Setup — Destination Instance](#part-2-api-app-setup--destination-instance)
- [Part 3: Run the Export Script](#part-3-run-the-export-script)
- [Part 4: Run the Import Script](#part-4-run-the-import-script)
- [What the Export JSON Contains](#what-the-export-json-contains)
- [The Manual Steps Checklist](#the-manual-steps-checklist)
- [What Gets Skipped](#what-gets-skipped)
- [API Endpoints Used](#api-endpoints-used)
- [Troubleshooting](#troubleshooting)
- [Pre-Flight Checklist](#pre-flight-checklist)

---

## What Is and Isn't Automated

Understanding this table before you start will save you time:

| Item | Exported | Imported Automatically | Manual Steps Required |
|---|---|---|---|
| **Global Custom Fields** | ✅ Full definition | ✅ Yes — created in destination | None |
| **Role Custom Fields** | ✅ Full definition | ✅ Yes — created in destination | Link fields to roles in portal |
| **Device Roles** | ✅ Name, class, description | ✅ Yes — created in destination | Assign policies and fields to roles |
| **Device Policies** | ✅ Shell only (name, class, enabled) | ✅ Yes — shell created in destination | Configure all conditions/rules manually |
| **Automation Scripts** | ✅ Full script body in JSON | ❌ No API endpoint exists | Paste each script manually in portal |
| **Policy conditions/rules** | ❌ Not exposed by API | ❌ N/A | Rebuild manually in destination portal |
| **Policy-to-role mappings** | ❌ Not exposed by API | ❌ N/A | Set manually after import |

The import script generates a **formatted checklist file** after running that lists every manual step, every script to recreate, and every policy to configure — so you have a complete to-do list.

---

## Critical Prerequisite — Custom Field API Permissions

> ⚠️ **This is the most important step. Do it before running the export script or you will get empty results.**

The NinjaOne API silently excludes any custom field whose API permission is set to `None`. If you run the export and get zero custom fields back, this is why.

**Before running the export, set API permission on every custom field you want to migrate:**

### For Global Custom Fields:
1. Go to: **Administration → Devices → Global Custom Fields**
2. Click **Edit** on each field
3. Under **Permissions**, find the **API** row
4. Set it to **Read Only** or **Read/Write** (either works for export)
5. Click **Save**

### For Role Custom Fields:
1. Go to: **Administration → Devices → Roles**
2. Click **Edit** on each role
3. Click the **Custom Fields** tab
4. For each field, click **Edit**
5. Set **API** permission to **Read Only** or **Read/Write**
6. Click **Save**

> ℹ️ You do not need to change Script or Technician permissions for the export — only API permission matters for the export script to see the field.

---

## Prerequisites

| Requirement | Notes |
|---|---|
| PowerShell 5.1+ | Built-in on Windows 10/11. Run `$PSVersionTable` to verify. |
| NinjaOne System Administrator | Required on **both** instances to create API apps |
| No extra modules | Uses only built-in PowerShell |
| Custom field API permissions set | See section above — must be done before export |

---

## Part 1: API App Setup — Source Instance

The export script needs **read-only** access to the source instance.

1. Log into the **source** NinjaOne portal as a System Administrator
2. Go to: **Administration → Apps → API → Client App IDs → Add**
3. Fill in:

   | Field | Value |
   |---|---|
   | **Name** | `InstanceExportScript` |
   | **Platform** | `API Services (Machine-to-Machine)` |
   | **Allowed Scopes** | ✅ `monitoring` only — read-only is sufficient |
   | **Redirect URI** | Leave blank |

4. Click **Save** — copy the **Client ID** and **Client Secret**

---

## Part 2: API App Setup — Destination Instance

The import script needs **write** access to the destination instance.

1. Log into the **destination** NinjaOne portal as a System Administrator
2. Go to: **Administration → Apps → API → Client App IDs → Add**
3. Fill in:

   | Field | Value |
   |---|---|
   | **Name** | `InstanceImportScript` |
   | **Platform** | `API Services (Machine-to-Machine)` |
   | **Allowed Scopes** | ✅ `monitoring` AND ✅ `management` |
   | **Redirect URI** | Leave blank |

4. Click **Save** — copy the **Client ID** and **Client Secret**

---

## Part 3: Run the Export Script

Open `Export-NinjaInstanceConfig.ps1` in any text editor and fill in the **CONFIGURATION** block at the top:

```powershell
# ==============================================================================
#  CONFIGURATION -- Fill in ALL values before running
# ==============================================================================

$SourceBaseUrl       = 'https://<source Login URL>'
$SourceTokenEndpoint = 'https://<source Login URL>/ws/oauth/token'
$SourceClientId      = '<Source Client ID>'
$SourceClientSecret  = '<Source Client Secret>'
$OutputFolder        = ''   # Leave blank to save next to the script
```

| Variable | Example | Notes |
|---|---|---|
| `$SourceBaseUrl` | `https://app.ninjarmm.com` | Source instance login URL |
| `$SourceTokenEndpoint` | `https://app.ninjarmm.com/ws/oauth/token` | Same URL + `/ws/oauth/token` |
| `$SourceClientId` | `abc123...` | From Part 1 |
| `$SourceClientSecret` | `s3cr3t...` | From Part 1 — shown once |
| `$OutputFolder` | `C:\Exports\` | Where to save the JSON. Leave blank for script folder. |

**Regional URLs:**

| Region | Base URL |
|---|---|
| United States | `https://app.ninjarmm.com` |
| Europe | `https://eu.ninjarmm.com` |
| Oceania | `https://oc.ninjarmm.com` |
| Canada | `https://ca.ninjarmm.com` |

**Run it:**
```powershell
.\Export-NinjaInstanceConfig.ps1
```

**Output files produced:**
- `NinjaExport_<instance>_<timestamp>.json` — the full export package (give this to the import script)
- `NinjaExport_<instance>_<timestamp>_Summary.txt` — human-readable summary of everything captured

**Example console output:**
```
  [1/6] Authenticating to source instance...
  [OK] Authenticated.

  [2/6] Exporting automation scripts...
  [OK] 12 custom automation script(s) exported.

  [3/6] Exporting custom fields...
  [OK] 8 global field(s), 4 role field(s) exported.

  [4/6] Exporting device roles...
  [OK] 5 device role(s) exported.

  [5/6] Exporting device policies...
  [OK] 9 policy/policies exported.

  [6/6] Saving export package...
  [OK] Export complete.

  JSON export : C:\Scripts\NinjaExport_app_ninjarmm_com_20260701_143022.json
  Summary     : C:\Scripts\NinjaExport_app_ninjarmm_com_20260701_143022_Summary.txt

  NEXT STEP: Run Import-NinjaInstanceConfig.ps1
  and point it at the JSON file above.
```

> ⚠️ If `global field(s)` shows 0 and you know fields exist, go back and set API permissions on those fields. See the [Critical Prerequisite](#critical-prerequisite--custom-field-api-permissions) section.

---

## Part 4: Run the Import Script

Open `Import-NinjaInstanceConfig.ps1` and fill in the **CONFIGURATION** block:

```powershell
# ==============================================================================
#  CONFIGURATION -- Fill in ALL values before running
# ==============================================================================

$DestBaseUrl       = 'https://<destination Login URL>'
$DestTokenEndpoint = 'https://<destination Login URL>/ws/oauth/token'
$DestClientId      = '<Destination Client ID>'
$DestClientSecret  = '<Destination Client Secret>'
$ImportJsonPath    = '<Path to NinjaExport_..._.json>'
```

| Variable | Example | Notes |
|---|---|---|
| `$DestBaseUrl` | `https://eu.ninjarmm.com` | Destination instance URL |
| `$DestTokenEndpoint` | `https://eu.ninjarmm.com/ws/oauth/token` | Same URL + `/ws/oauth/token` |
| `$DestClientId` | `xyz789...` | From Part 2 |
| `$DestClientSecret` | `s3cr3t...` | From Part 2 |
| `$ImportJsonPath` | `C:\Exports\NinjaExport_app_ninjarmm_com_20260701_143022.json` | Path to the JSON from the export step |

**Run it:**
```powershell
.\Import-NinjaInstanceConfig.ps1
```

The script works through 7 steps:
1. Loads and validates the JSON export
2. Authenticates to the destination
3. Loads existing config from destination (for duplicate detection)
4. Creates global custom fields
5. Creates role custom fields
6. Creates device roles
7. Creates device policy shells

**Nothing is ever overwritten or deleted.** If an item already exists in the destination (matched by name), it is skipped and logged as `Skipped (exists)`.

**Output files produced:**
- `NinjaImport_Checklist_<timestamp>.txt` — your complete manual steps to-do list

---

## What the Export JSON Contains

The JSON file is a single compressed object with these top-level keys:

```json
{
  "exportedAt": "2026-07-01 14:30:22",
  "sourceInstance": "https://app.ninjarmm.com",
  "scriptCount": 12,
  "globalFieldCount": 8,
  "roleFieldCount": 4,
  "roleCount": 5,
  "policyCount": 9,
  "automationScripts": [ ... ],
  "globalCustomFields": [ ... ],
  "roleCustomFields": [ ... ],
  "deviceRoles": [ ... ],
  "devicePolicies": [ ... ]
}
```

Each `automationScripts` entry contains the full `scriptBody` — the actual script text — so nothing is lost even though it must be manually pasted into the destination.

---

## The Manual Steps Checklist

After the import script runs, it saves a checklist file next to the JSON. It contains:

- A checkbox list of every automation script to recreate, with name, language, OS, and Run As values
- A checkbox list of every policy to configure with its conditions and rules
- Step-by-step instructions for assigning policies to roles
- Step-by-step instructions for linking role custom fields to roles

**Use this file as your working to-do list** until the destination instance matches the source. Check off each item as you complete it.

---

## What Gets Skipped

**Built-in NinjaOne scripts** are automatically excluded from the export. Only custom scripts you created appear in the export. This is by design — built-in scripts already exist in every instance.

**Existing items** in the destination are never touched. The script matches by name — if a custom field, role, or policy with the same name already exists, it is skipped and noted in the results table.

**Default policies** (the NinjaOne built-in defaults) are exported for reference but may already exist in the destination. They will be skipped if found.

---

## API Endpoints Used

### Export Script

| Method | Endpoint | Purpose |
|---|---|---|
| `POST` | `/ws/oauth/token` | Authenticate to source (Client Credentials) |
| `GET` | `/v2/automation/scripts` | List all automation scripts |
| `GET` | `/v2/custom-fields` | List all custom field definitions |
| `GET` | `/v2/roles` | List all device roles |
| `GET` | `/v2/policies` | List all device policies (paginated) |

### Import Script

| Method | Endpoint | Purpose |
|---|---|---|
| `POST` | `/ws/oauth/token` | Authenticate to destination (Client Credentials) |
| `GET` | `/v2/custom-fields` | Check for existing fields (duplicate detection) |
| `GET` | `/v2/roles` | Check for existing roles |
| `GET` | `/v2/policies` | Check for existing policies |
| `POST` | `/v2/custom-fields` | Create each custom field |
| `POST` | `/v2/roles` | Create each device role |
| `POST` | `/v2/policies` | Create each policy shell |

---

## Troubleshooting

| Error / Symptom | Solution |
|---|---|
| `Fill in $SourceBaseUrl` or similar | The `<placeholder>` text is still in the config block. Replace it. |
| Export returns 0 custom fields | API permissions on fields are set to `None`. See [Critical Prerequisite](#critical-prerequisite--custom-field-api-permissions). |
| Export returns 0 scripts | You may have no custom scripts (only built-ins). Built-ins are intentionally excluded. |
| `Authentication failed` on source | Check `$SourceClientId` and `$SourceClientSecret`. API app must be `monitoring` scope, Machine-to-Machine platform. |
| `Authentication failed` on destination | Check `$DestClientId` and `$DestClientSecret`. App must have both `monitoring` and `management` scopes. |
| `Import JSON not found` | The path in `$ImportJsonPath` is wrong. Use the full absolute path to the JSON file. |
| `HTTP 400` on field creation | The `fieldType` value from the source may not be valid in the destination's version. Check the field type in the source portal and recreate manually if needed. |
| `HTTP 403` on import | The destination API app is missing the `management` scope. Edit it in Administration → Apps → API. |
| Policy created but empty | Expected — policy shells are created without conditions. Use the checklist to configure each one manually. |
| Role field not appearing on devices | Role custom fields must be assigned to roles in the portal after creation. See the checklist Step 4. |
| Script body looks garbled in JSON | Some scripts contain special characters. Open the JSON in VS Code or Notepad++ with UTF-8 encoding for accurate display. |

---

## Pre-Flight Checklist

### Before Export
- [ ] Logged into source NinjaOne as System Administrator
- [ ] API permissions set to `Read Only` or `Read/Write` on **all** custom fields (global and role)
- [ ] API app created on source instance — scope: `monitoring`
- [ ] Source Client ID and Client Secret saved
- [ ] `$SourceBaseUrl`, `$SourceTokenEndpoint`, `$SourceClientId`, `$SourceClientSecret` filled in
- [ ] Export script run — confirmed counts in output match expectations
- [ ] Opened the Summary .txt and reviewed what was captured

### Before Import
- [ ] API app created on destination instance — scopes: `monitoring` AND `management`
- [ ] Destination Client ID and Client Secret saved
- [ ] `$DestBaseUrl`, `$DestTokenEndpoint`, `$DestClientId`, `$DestClientSecret` filled in
- [ ] `$ImportJsonPath` points to the correct JSON file
- [ ] Import script run — reviewed per-item results table

### After Import — Manual Steps
- [ ] Opened `NinjaImport_Checklist_<timestamp>.txt`
- [ ] Recreated each automation script in destination portal (Administration → Library → Automation)
- [ ] Configured conditions and rules on each policy shell (Administration → Policies)
- [ ] Assigned policies to device roles (Administration → Devices → Roles)
- [ ] Linked role custom fields to device roles (Administration → Devices → Roles → Custom Fields tab)
- [ ] Verified destination matches source by spot-checking a few devices, roles, and policies
