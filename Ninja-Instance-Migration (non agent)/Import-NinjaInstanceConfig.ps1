#Requires -Version 5.1
<#
.SYNOPSIS
    SCRIPT 2 OF 2 -- IMPORT
    Reads the JSON export produced by Export-NinjaInstanceConfig.ps1 and
    imports custom fields, device roles, and device policies into a destination
    NinjaOne instance. Generates a checklist of manual steps for items the
    API cannot create (automation scripts, policy conditions).

.DESCRIPTION
    Uses the NinjaOne Public API v2 with Client Credentials (silent, no browser).

    What is imported automatically:
      - Global custom fields    : created via POST /v2/custom-fields
      - Role custom fields      : created via POST /v2/custom-fields
      - Device roles            : created via POST /v2/roles
      - Device policies (shell) : created via POST /v2/policies

    What requires manual steps (checklist generated automatically):
      - Automation scripts      : no create endpoint in public API
      - Policy conditions/rules : not exposed via API
      - Role-to-policy mappings : must be set in the portal after import

    Existing items (matched by name) are skipped -- the script never
    overwrites or deletes anything in the destination instance.

    API APP SETUP (one-time on DESTINATION instance):
    Administration > Apps > API > Client App IDs > Add
      Platform      : API Services (Machine-to-Machine)
      Allowed Scopes: management  (write access required for import)
      Redirect URI  : Leave blank

    REGIONAL URLS:
    US: https://app.ninjarmm.com  |  EU: https://eu.ninjarmm.com
    OC: https://oc.ninjarmm.com   |  CA: https://ca.ninjarmm.com
#>

# ==============================================================================
#  CONFIGURATION -- Fill in ALL values before running
# ==============================================================================

# Destination instance login URL (no trailing slash)
$DestBaseUrl       = 'https://<destination Login URL>'

# Destination instance token endpoint
$DestTokenEndpoint = 'https://<destination Login URL>/ws/oauth/token'

# API credentials for the DESTINATION instance
$DestClientId      = '<Destination Client ID>'
$DestClientSecret  = '<Destination Client Secret>'

# Full path to the JSON file produced by Export-NinjaInstanceConfig.ps1
$ImportJsonPath    = '<Path to NinjaExport_..._.json>'

# ==============================================================================
#  DO NOT EDIT BELOW THIS LINE
# ==============================================================================

try {
    [Net.ServicePointManager]::SecurityProtocol = `
        [Net.ServicePointManager]::SecurityProtocol -bor `
        [Net.SecurityProtocolType]::Tls12
} catch {}

try { [Console]::OutputEncoding = [System.Text.Encoding]::UTF8 } catch {}

$ErrorActionPreference = 'Continue'

# -- Validate config ----------------------------------------------------------
$Errs = @()
if ($DestBaseUrl       -like '*<*') { $Errs += 'Fill in $DestBaseUrl' }
if ($DestClientId      -like '*<*') { $Errs += 'Fill in $DestClientId' }
if ($DestClientSecret  -like '*<*') { $Errs += 'Fill in $DestClientSecret' }
if ($ImportJsonPath    -like '*<*') { $Errs += 'Fill in $ImportJsonPath' }
if ($Errs.Count -gt 0) {
    Write-Host ""
    Write-Host "  [!] Configuration errors -- please fix:" -ForegroundColor Red
    $Errs | ForEach-Object { Write-Host "      - $_" -ForegroundColor Red }
    exit 1
}

$ResolvedJson = if ([System.IO.Path]::IsPathRooted($ImportJsonPath)) {
    $ImportJsonPath
} elseif ($PSScriptRoot) {
    Join-Path $PSScriptRoot $ImportJsonPath
} else {
    Join-Path $PWD.Path $ImportJsonPath
}

if (-not (Test-Path -LiteralPath $ResolvedJson)) {
    Write-Host "  [!] Import JSON not found: $ResolvedJson" -ForegroundColor Red
    exit 1
}

# -- Safe property helper -----------------------------------------------------
function Get-Prop {
    param([object]$Obj, [string]$Name, [object]$Default = $null)
    if ($null -eq $Obj) { return $Default }
    $p = $Obj.PSObject.Properties[$Name]
    if ($null -eq $p -or $null -eq $p.Value) { return $Default }
    return $p.Value
}

# -- API helper with retry ----------------------------------------------------
function Invoke-NinjaApi {
    param(
        [string]$Path,
        [string]$Method  = 'GET',
        [string]$Body    = $null,
        [int]   $Retries = 3
    )
    $Attempt = 0
    while ($true) {
        $Attempt++
        $Params = @{
            Uri     = "$script:DestBaseUrl/v2/$Path"
            Method  = $Method
            Headers = $script:DestHeaders
        }
        if ($Body) {
            $Params.Body        = $Body
            $Params.ContentType = 'application/json'
        }
        try {
            return Invoke-RestMethod @Params
        } catch {
            $sc = $null; try { $sc = [int]$_.Exception.Response.StatusCode } catch {}
            if ($sc -ge 400 -and $sc -lt 500 -and $sc -ne 429) {
                throw "HTTP $sc on $Method /v2/$Path -- $_"
            }
            if ($Attempt -gt $Retries) {
                throw "Failed after $Retries retries on $Method /v2/$Path -- $_"
            }
            Write-Host "    [~] Retry $Attempt/$Retries (HTTP $sc)..." -ForegroundColor Yellow
            Start-Sleep -Seconds 5
        }
    }
}

Write-Host ""
Write-Host "  ================================================================" -ForegroundColor Cyan
Write-Host "  NinjaOne Instance Config Importer" -ForegroundColor Cyan
Write-Host "  Destination: $DestBaseUrl" -ForegroundColor Cyan
Write-Host "  ================================================================" -ForegroundColor Cyan
Write-Host ""

# =============================================================================
#  STEP 1: Load export package
# =============================================================================
Write-Host "  [1/7] Loading export package..." -ForegroundColor Cyan
try {
    $JsonContent   = Get-Content -LiteralPath $ResolvedJson -Raw -Encoding UTF8
    $ExportPackage = $JsonContent | ConvertFrom-Json
} catch {
    Write-Host "  [!] Failed to read/parse JSON: $_" -ForegroundColor Red
    exit 1
}

$SrcScripts      = @($ExportPackage.automationScripts)
$SrcGlobalFields = @($ExportPackage.globalCustomFields)
$SrcRoleFields   = @($ExportPackage.roleCustomFields)
$SrcRoles        = @($ExportPackage.deviceRoles)
$SrcPolicies     = @($ExportPackage.devicePolicies)

Write-Host "  [OK] Loaded export from: $(Get-Prop $ExportPackage 'sourceInstance')" -ForegroundColor Green
Write-Host "       Exported at: $(Get-Prop $ExportPackage 'exportedAt')" -ForegroundColor Gray
Write-Host "       Scripts: $($SrcScripts.Count)  GlobalFields: $($SrcGlobalFields.Count)  RoleFields: $($SrcRoleFields.Count)  Roles: $($SrcRoles.Count)  Policies: $($SrcPolicies.Count)" -ForegroundColor Gray

# =============================================================================
#  STEP 2: Authenticate to destination
# =============================================================================
Write-Host ""
Write-Host "  [2/7] Authenticating to destination instance..." -ForegroundColor Cyan
try {
    $Token = Invoke-RestMethod -Uri $DestTokenEndpoint -Method POST `
        -ContentType 'application/x-www-form-urlencoded' `
        -Body @{
            grant_type    = 'client_credentials'
            client_id     = $DestClientId
            client_secret = $DestClientSecret
            scope         = 'monitoring management'
        }
    $script:DestHeaders = @{
        Authorization  = "Bearer $($Token.access_token)"
        Accept         = 'application/json'
        'Content-Type' = 'application/json'
    }
} catch {
    Write-Host "  [!] Authentication to destination failed: $_" -ForegroundColor Red
    exit 1
}
Write-Host "  [OK] Authenticated to destination." -ForegroundColor Green

# =============================================================================
#  STEP 3: Load existing destination data (for duplicate detection)
# =============================================================================
Write-Host ""
Write-Host "  [3/7] Loading existing destination config (duplicate detection)..." -ForegroundColor Cyan

$ExistingFields   = @{}
$ExistingRoles    = @{}
$ExistingPolicies = @{}

try {
    $DestFields = Invoke-NinjaApi -Path 'custom-fields'
    $FieldArr   = if ($DestFields -is [array]) { $DestFields } else { @($DestFields) }
    foreach ($f in $FieldArr) {
        $n = Get-Prop $f 'name'
        if ($n) { $ExistingFields[$n.ToLower()] = $true }
    }
} catch { Write-Host "  [i] Could not load existing custom fields -- skipping dupe check." -ForegroundColor Gray }

try {
    $DestRoles = Invoke-NinjaApi -Path 'roles'
    $RoleArr   = if ($DestRoles -is [array]) { $DestRoles } else { @($DestRoles) }
    foreach ($r in $RoleArr) {
        $n = Get-Prop $r 'name'
        if ($n) { $ExistingRoles[$n.ToLower()] = $true }
    }
} catch { Write-Host "  [i] Could not load existing roles -- skipping dupe check." -ForegroundColor Gray }

try {
    $DestPolicies = Invoke-NinjaApi -Path 'policies'
    $PolArr       = if ($DestPolicies -is [array]) { $DestPolicies } else { @($DestPolicies) }
    foreach ($p in $PolArr) {
        $n = Get-Prop $p 'name'
        if ($n) { $ExistingPolicies[$n.ToLower()] = $true }
    }
} catch { Write-Host "  [i] Could not load existing policies -- skipping dupe check." -ForegroundColor Gray }

Write-Host "  [OK] Destination has: $($ExistingFields.Count) fields, $($ExistingRoles.Count) roles, $($ExistingPolicies.Count) policies." -ForegroundColor Green

# Track results
$Results = New-Object System.Collections.ArrayList

# =============================================================================
#  STEP 4: Import global custom fields
# =============================================================================
Write-Host ""
Write-Host "  [4/7] Importing global custom fields ($($SrcGlobalFields.Count))..." -ForegroundColor Cyan

$FieldsCreated = 0; $FieldsSkipped = 0

foreach ($f in $SrcGlobalFields) {
    $fname = Get-Prop $f 'name'
    if (-not $fname) { continue }

    if ($ExistingFields.ContainsKey($fname.ToLower())) {
        Write-Host "    [~] Field '$fname' already exists -- skipping." -ForegroundColor Gray
        $FieldsSkipped++
        [void]$Results.Add([PSCustomObject]@{ Type='GlobalField'; Name=$fname; Status='Skipped (exists)' })
        continue
    }

    try {
        $Body = [ordered]@{
            name             = $fname
            label            = Get-Prop $f 'label'
            description      = Get-Prop $f 'description'
            fieldType        = Get-Prop $f 'fieldType'
            definitionScopes = @('NODE')
            technicianPermission = Get-Prop $f 'technicianPermission' -Default 'READ_ONLY'
            scriptPermission     = Get-Prop $f 'scriptPermission'     -Default 'READ_WRITE'
            apiPermission        = Get-Prop $f 'apiPermission'        -Default 'READ_WRITE'
        } | ConvertTo-Json -Compress -Depth 3

        Invoke-NinjaApi -Path 'custom-fields' -Method POST -Body $Body | Out-Null
        Write-Host "    [OK] Field '$fname' created." -ForegroundColor Green
        $FieldsCreated++
        $ExistingFields[$fname.ToLower()] = $true
        [void]$Results.Add([PSCustomObject]@{ Type='GlobalField'; Name=$fname; Status='Created' })
        Start-Sleep -Milliseconds 300
    } catch {
        Write-Host "    [!] Failed to create field '$fname': $_" -ForegroundColor Red
        [void]$Results.Add([PSCustomObject]@{ Type='GlobalField'; Name=$fname; Status="Error: $_" })
    }
}
Write-Host "  [OK] Global fields: $FieldsCreated created, $FieldsSkipped skipped." -ForegroundColor Green

# =============================================================================
#  STEP 5: Import role custom fields
# =============================================================================
Write-Host ""
Write-Host "  [5/7] Importing role custom fields ($($SrcRoleFields.Count))..." -ForegroundColor Cyan

$RoleFieldsCreated = 0; $RoleFieldsSkipped = 0

foreach ($f in $SrcRoleFields) {
    $fname = Get-Prop $f 'name'
    if (-not $fname) { continue }

    if ($ExistingFields.ContainsKey($fname.ToLower())) {
        Write-Host "    [~] Role field '$fname' already exists -- skipping." -ForegroundColor Gray
        $RoleFieldsSkipped++
        [void]$Results.Add([PSCustomObject]@{ Type='RoleField'; Name=$fname; Status='Skipped (exists)' })
        continue
    }

    try {
        $scopes = Get-Prop $f 'definitionScopes'
        if (-not $scopes) { $scopes = @('NODE_ROLE') }

        $Body = [ordered]@{
            name             = $fname
            label            = Get-Prop $f 'label'
            description      = Get-Prop $f 'description'
            fieldType        = Get-Prop $f 'fieldType'
            definitionScopes = $scopes
            technicianPermission = Get-Prop $f 'technicianPermission' -Default 'READ_ONLY'
            scriptPermission     = Get-Prop $f 'scriptPermission'     -Default 'READ_WRITE'
            apiPermission        = Get-Prop $f 'apiPermission'        -Default 'READ_WRITE'
        } | ConvertTo-Json -Compress -Depth 3

        Invoke-NinjaApi -Path 'custom-fields' -Method POST -Body $Body | Out-Null
        Write-Host "    [OK] Role field '$fname' created." -ForegroundColor Green
        $RoleFieldsCreated++
        $ExistingFields[$fname.ToLower()] = $true
        [void]$Results.Add([PSCustomObject]@{ Type='RoleField'; Name=$fname; Status='Created' })
        Start-Sleep -Milliseconds 300
    } catch {
        Write-Host "    [!] Failed to create role field '$fname': $_" -ForegroundColor Red
        [void]$Results.Add([PSCustomObject]@{ Type='RoleField'; Name=$fname; Status="Error: $_" })
    }
}
Write-Host "  [OK] Role fields: $RoleFieldsCreated created, $RoleFieldsSkipped skipped." -ForegroundColor Green

# =============================================================================
#  STEP 6: Import device roles
# =============================================================================
Write-Host ""
Write-Host "  [6/7] Importing device roles ($($SrcRoles.Count))..." -ForegroundColor Cyan

$RolesCreated = 0; $RolesSkipped = 0

foreach ($r in $SrcRoles) {
    $rname = Get-Prop $r 'name'
    if (-not $rname) { continue }

    if ($ExistingRoles.ContainsKey($rname.ToLower())) {
        Write-Host "    [~] Role '$rname' already exists -- skipping." -ForegroundColor Gray
        $RolesSkipped++
        [void]$Results.Add([PSCustomObject]@{ Type='Role'; Name=$rname; Status='Skipped (exists)' })
        continue
    }

    try {
        $Body = [ordered]@{
            name        = $rname
            description = Get-Prop $r 'description'
            nodeClass   = Get-Prop $r 'nodeClass'
        } | ConvertTo-Json -Compress -Depth 2

        Invoke-NinjaApi -Path 'roles' -Method POST -Body $Body | Out-Null
        Write-Host "    [OK] Role '$rname' created." -ForegroundColor Green
        $RolesCreated++
        $ExistingRoles[$rname.ToLower()] = $true
        [void]$Results.Add([PSCustomObject]@{ Type='Role'; Name=$rname; Status='Created' })
        Start-Sleep -Milliseconds 300
    } catch {
        Write-Host "    [!] Failed to create role '$rname': $_" -ForegroundColor Red
        [void]$Results.Add([PSCustomObject]@{ Type='Role'; Name=$rname; Status="Error: $_" })
    }
}
Write-Host "  [OK] Roles: $RolesCreated created, $RolesSkipped skipped." -ForegroundColor Green

# =============================================================================
#  STEP 7: Import device policies (shells only)
# =============================================================================
Write-Host ""
Write-Host "  [7/7] Importing device policy shells ($($SrcPolicies.Count))..." -ForegroundColor Cyan

$PoliciesCreated = 0; $PoliciesSkipped = 0

foreach ($p in $SrcPolicies) {
    $pname = Get-Prop $p 'name'
    if (-not $pname) { continue }

    if ($ExistingPolicies.ContainsKey($pname.ToLower())) {
        Write-Host "    [~] Policy '$pname' already exists -- skipping." -ForegroundColor Gray
        $PoliciesSkipped++
        [void]$Results.Add([PSCustomObject]@{ Type='Policy'; Name=$pname; Status='Skipped (exists)' })
        continue
    }

    try {
        $Body = [ordered]@{
            name        = $pname
            description = Get-Prop $p 'description'
            nodeClass   = Get-Prop $p 'nodeClass'
            enabled     = if ($null -ne (Get-Prop $p 'enabled')) { [bool](Get-Prop $p 'enabled') } else { $true }
        } | ConvertTo-Json -Compress -Depth 2

        Invoke-NinjaApi -Path 'policies' -Method POST -Body $Body | Out-Null
        Write-Host "    [OK] Policy '$pname' [$($p.nodeClass)] created." -ForegroundColor Green
        $PoliciesCreated++
        $ExistingPolicies[$pname.ToLower()] = $true
        [void]$Results.Add([PSCustomObject]@{ Type='Policy'; Name=$pname; Status='Created' })
        Start-Sleep -Milliseconds 300
    } catch {
        Write-Host "    [!] Failed to create policy '$pname': $_" -ForegroundColor Red
        [void]$Results.Add([PSCustomObject]@{ Type='Policy'; Name=$pname; Status="Error: $_" })
    }
}
Write-Host "  [OK] Policies: $PoliciesCreated created, $PoliciesSkipped skipped." -ForegroundColor Green

# =============================================================================
#  GENERATE MANUAL CHECKLIST
# =============================================================================
$Timestamp    = Get-Date -Format 'yyyyMMdd_HHmmss'
$ChecklistDir = [System.IO.Path]::GetDirectoryName([System.IO.Path]::GetFullPath($ResolvedJson))
$ChecklistPath = Join-Path $ChecklistDir "NinjaImport_Checklist_$Timestamp.txt"

$Checklist = @"
NinjaOne Import -- Manual Steps Checklist
==========================================
Generated   : $(Get-Date -Format 'yyyy-MM-dd HH:mm:ss')
Destination : $DestBaseUrl
Source JSON : $ResolvedJson

AUTOMATED RESULTS
-----------------
Global Custom Fields : $FieldsCreated created, $FieldsSkipped skipped
Role Custom Fields   : $RoleFieldsCreated created, $RoleFieldsSkipped skipped
Device Roles         : $RolesCreated created, $RolesSkipped skipped
Device Policies      : $PoliciesCreated created, $PoliciesSkipped skipped (shells only)

============================================================
MANUAL STEPS -- Complete these in the destination portal
============================================================

STEP 1: AUTOMATION SCRIPTS (must be created manually)
-------------------------------------------------------
The NinjaOne API has no endpoint to create automation scripts programmatically.
Each script must be recreated manually. Open the export JSON and find the
"automationScripts" array. For each entry:

  a) Go to: Administration > Library > Automation > Add > New Script
  b) Set the Name, Language, OS, Architecture, and Run As fields
  c) Paste the script body from the JSON (automationScripts[].scriptBody)
  d) Click Save

Scripts to recreate ($($SrcScripts.Count)):
$(if ($SrcScripts.Count -eq 0) { '  (none)' } else {
    ($SrcScripts | ForEach-Object {
        "  [ ] $($_.name)`n       Language: $($_.language)  OS: $($_.operatingSystem)  RunAs: $($_.runAs)"
    }) -join "`n"
})

STEP 2: POLICY CONDITIONS AND RULES (must be configured manually)
------------------------------------------------------------------
The policy shells have been created automatically. However, the internal
configuration of each policy (patch schedules, monitoring conditions,
software management, automation triggers, etc.) cannot be transferred
via the API. Each policy must be fully configured in the destination portal.

Policies to configure ($($SrcPolicies.Count)):
$(if ($SrcPolicies.Count -eq 0) { '  (none)' } else {
    ($SrcPolicies | ForEach-Object {
        "  [ ] $($_.name) [$($_.nodeClass)]"
    }) -join "`n"
})

For each policy:
  a) Go to: Administration > Policies
  b) Open the policy and configure all tabs:
       - General, Patching, Monitoring, Scripting, Tray Icon, etc.
  c) Compare against the source instance to match the configuration

STEP 3: ROLE-TO-POLICY MAPPINGS
---------------------------------
After policies are configured, assign them to device roles:
  a) Go to: Administration > Devices > Roles
  b) Edit each role
  c) Assign the appropriate policy on the Policy tab

Role-to-policy mappings from source:
  (Open source NinjaOne portal to view current mappings)

STEP 4: ROLE CUSTOM FIELD ASSIGNMENTS
---------------------------------------
Role custom fields were created but must be manually linked to roles:
  a) Go to: Administration > Devices > Roles
  b) Edit each role
  c) Add the relevant custom fields on the Custom Fields tab

============================================================
END OF CHECKLIST
============================================================
"@

$Checklist | Set-Content -LiteralPath $ChecklistPath -Encoding UTF8

# =============================================================================
#  FINAL SUMMARY
# =============================================================================
Write-Host ""
Write-Host "  ================================================================" -ForegroundColor Green
Write-Host "  IMPORT COMPLETE" -ForegroundColor Green
Write-Host "    Global Fields : $FieldsCreated created, $FieldsSkipped skipped" -ForegroundColor Green
Write-Host "    Role Fields   : $RoleFieldsCreated created, $RoleFieldsSkipped skipped" -ForegroundColor Green
Write-Host "    Roles         : $RolesCreated created, $RolesSkipped skipped" -ForegroundColor Green
Write-Host "    Policies      : $PoliciesCreated created, $PoliciesSkipped skipped" -ForegroundColor Green
Write-Host "  ================================================================" -ForegroundColor Green
Write-Host ""
Write-Host "  Checklist saved: $ChecklistPath" -ForegroundColor Cyan
Write-Host ""
Write-Host "  IMPORTANT -- Review the checklist for manual steps:" -ForegroundColor Yellow
Write-Host "  - Recreate $($SrcScripts.Count) automation script(s) manually in the portal" -ForegroundColor Yellow
Write-Host "  - Configure conditions/rules on $PoliciesCreated policy shell(s)" -ForegroundColor Yellow
Write-Host "  - Assign policies to roles and link role custom fields" -ForegroundColor Yellow
Write-Host ""

if ($Results.Count -gt 0) {
    Write-Host "  Per-item results:" -ForegroundColor Cyan
    $Results | Format-Table -AutoSize Type, Name, Status
}
