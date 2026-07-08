#Requires -Version 5.1
<#
.SYNOPSIS
    SCRIPT 1 OF 2 — EXPORT
    Exports automation scripts, role custom fields, global custom fields, device
    roles, and device policies from a source NinjaOne instance and saves them
    to a single compressed JSON file for import into a destination instance.

.DESCRIPTION
    Uses the NinjaOne Public API v2 with Client Credentials (silent, no browser).

    What is exported:
      - Automation scripts       : name, language, OS, architecture, runAs, script content
      - Global custom fields     : name, label, type, scope, permissions
      - Role custom fields       : name, label, type, scope, permissions, associated role
      - Device roles             : name, description, nodeClass
      - Device policies          : name, description, nodeClass, parentPolicyId, enabled

    What CANNOT be exported via the API (must be recreated manually):
      - Policy internal conditions, rules, patch schedules, monitoring thresholds
      - Script category assignments
      - Policy-to-role mappings beyond basic metadata

    The export JSON and a human-readable summary report are saved to the folder
    you configure below.

    IMPORTANT -- BEFORE RUNNING:
    Every custom field you want exported MUST have API Permission set to
    "Read Only" or "Read/Write" in NinjaOne. Fields with API Permission = "None"
    will be silently excluded from the API response.

    To check/set this:
      Administration > Devices > Global Custom Fields > Edit each field
      Administration > Devices > Roles > Edit each role > Custom Fields tab
      Set "API" permission to at least "Read Only" on every field.

    API APP SETUP (one-time on SOURCE instance):
    Administration > Apps > API > Client App IDs > Add
      Platform      : API Services (Machine-to-Machine)
      Allowed Scopes: monitoring  (read-only is sufficient for export)
      Redirect URI  : Leave blank

    REGIONAL URLS:
    US: https://app.ninjarmm.com  |  EU: https://eu.ninjarmm.com
    OC: https://oc.ninjarmm.com   |  CA: https://ca.ninjarmm.com
#>

# ==============================================================================
#  CONFIGURATION -- Fill in ALL values before running
# ==============================================================================

# Source instance login URL (no trailing slash)
$SourceBaseUrl       = 'https://<source Login URL>'

# Source instance token endpoint
$SourceTokenEndpoint = 'https://<source Login URL>/ws/oauth/token'

# API credentials for the SOURCE instance
$SourceClientId      = '<Source Client ID>'
$SourceClientSecret  = '<Source Client Secret>'

# Folder where the export JSON and summary report will be saved
# Leave blank to save in the same folder as this script
$OutputFolder        = ''

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
if ($SourceBaseUrl       -like '*<*') { $Errs += 'Fill in $SourceBaseUrl' }
if ($SourceClientId      -like '*<*') { $Errs += 'Fill in $SourceClientId' }
if ($SourceClientSecret  -like '*<*') { $Errs += 'Fill in $SourceClientSecret' }
if ($Errs.Count -gt 0) {
    Write-Host ""
    Write-Host "  [!] Configuration errors -- please fix before running:" -ForegroundColor Red
    $Errs | ForEach-Object { Write-Host "      - $_" -ForegroundColor Red }
    exit 1
}

# -- Resolve output folder ----------------------------------------------------
if ([string]::IsNullOrWhiteSpace($OutputFolder)) {
    $OutputFolder = if ($PSScriptRoot) { $PSScriptRoot } else { $PWD.Path }
}
if (-not (Test-Path -LiteralPath $OutputFolder)) {
    try { New-Item -ItemType Directory -Path $OutputFolder -Force | Out-Null }
    catch { Write-Host "  [!] Cannot create output folder: $OutputFolder" -ForegroundColor Red; exit 1 }
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
    param([string]$BaseUrl, [hashtable]$Headers, [string]$Path, [int]$Retries = 3)
    $Attempt = 0
    while ($true) {
        $Attempt++
        try {
            return Invoke-RestMethod -Uri "$BaseUrl/v2/$Path" -Method GET -Headers $Headers
        } catch {
            $sc = $null; try { $sc = [int]$_.Exception.Response.StatusCode } catch {}
            if ($sc -ge 400 -and $sc -lt 500) { throw "HTTP $sc on GET /v2/$Path -- $_" }
            if ($Attempt -gt $Retries) { throw "Failed after $Retries retries on GET /v2/$Path -- $_" }
            Write-Host "    [~] Retry $Attempt/$Retries on /v2/$Path (HTTP $sc)..." -ForegroundColor Yellow
            Start-Sleep -Seconds 5
        }
    }
}

# -- Paginated GET helper -----------------------------------------------------
function Get-AllPages {
    param([string]$BaseUrl, [hashtable]$Headers, [string]$Path, [int]$PageSize = 200)
    $All   = New-Object System.Collections.ArrayList
    $After = $null
    do {
        $QS   = if ($Path -match '\?') { "${Path}&pageSize=$PageSize" } else { "${Path}?pageSize=$PageSize" }
        if ($After) { $QS += "&after=$After" }
        $Page  = Invoke-NinjaApi -BaseUrl $BaseUrl -Headers $Headers -Path $QS.TrimStart('/')
        $Items = if ($Page -is [array]) { $Page }
                 elseif ($Page.PSObject.Properties['results']) { $Page.results }
                 else { @($Page) }
        foreach ($i in $Items) { [void]$All.Add($i) }
        $After = if ($Items.Count -eq $PageSize) { Get-Prop $Items[-1] 'id' } else { $null }
    } while ($After)
    return $All
}

Write-Host ""
Write-Host "  ================================================================" -ForegroundColor Cyan
Write-Host "  NinjaOne Instance Config Exporter" -ForegroundColor Cyan
Write-Host "  Source: $SourceBaseUrl" -ForegroundColor Cyan
Write-Host "  ================================================================" -ForegroundColor Cyan
Write-Host ""

# =============================================================================
#  STEP 1: Authenticate to source
# =============================================================================
Write-Host "  [1/6] Authenticating to source instance..." -ForegroundColor Cyan
try {
    $Token = Invoke-RestMethod -Uri $SourceTokenEndpoint -Method POST `
        -ContentType 'application/x-www-form-urlencoded' `
        -Body @{
            grant_type    = 'client_credentials'
            client_id     = $SourceClientId
            client_secret = $SourceClientSecret
            scope         = 'monitoring'
        }
    $SrcHeaders = @{
        Authorization = "Bearer $($Token.access_token)"
        Accept        = 'application/json'
    }
} catch {
    Write-Host "  [!] Authentication failed: $_" -ForegroundColor Red
    exit 1
}
Write-Host "  [OK] Authenticated." -ForegroundColor Green

# =============================================================================
#  STEP 2: Export automation scripts
# =============================================================================
Write-Host ""
Write-Host "  [2/6] Exporting automation scripts..." -ForegroundColor Cyan
$Scripts = @()
try {
    $RawScripts = Invoke-NinjaApi -BaseUrl $SourceBaseUrl -Headers $SrcHeaders -Path 'automation/scripts'
    $ScriptArr  = if ($RawScripts -is [array]) { $RawScripts } else { @($RawScripts) }

    # Filter to custom scripts only (exclude NinjaOne built-ins)
    $Scripts = $ScriptArr | Where-Object {
        (Get-Prop $_ 'scope') -ne 'BUILTIN' -and
        (Get-Prop $_ 'scriptSource') -ne 'BUILTIN'
    } | ForEach-Object {
        [PSCustomObject]@{
            name             = Get-Prop $_ 'name'
            language         = Get-Prop $_ 'language'
            operatingSystem  = Get-Prop $_ 'operatingSystem'
            architecture     = Get-Prop $_ 'architecture'
            runAs            = Get-Prop $_ 'runAs'
            description      = Get-Prop $_ 'description'
            scriptBody       = Get-Prop $_ 'script'
            parameters       = Get-Prop $_ 'parameters'
        }
    }
    Write-Host "  [OK] $($Scripts.Count) custom automation script(s) exported." -ForegroundColor Green
} catch {
    Write-Host "  [!] Could not export automation scripts: $_" -ForegroundColor Yellow
    Write-Host "      Continuing with remaining exports..." -ForegroundColor Yellow
}

# =============================================================================
#  STEP 3: Export custom fields (global and role-scoped)
# =============================================================================
Write-Host ""
Write-Host "  [3/6] Exporting custom fields..." -ForegroundColor Cyan
$GlobalFields = @()
$RoleFields   = @()
try {
    $AllFields = Invoke-NinjaApi -BaseUrl $SourceBaseUrl -Headers $SrcHeaders -Path 'custom-fields'
    $FieldArr  = if ($AllFields -is [array]) { $AllFields } else { @($AllFields) }

    foreach ($f in $FieldArr) {
        $fieldObj = [PSCustomObject]@{
            name                 = Get-Prop $f 'name'
            label                = Get-Prop $f 'label'
            description          = Get-Prop $f 'description'
            fieldType            = Get-Prop $f 'fieldType'
            definitionScopes     = Get-Prop $f 'definitionScopes'
            technicianPermission = Get-Prop $f 'technicianPermission'
            scriptPermission     = Get-Prop $f 'scriptPermission'
            apiPermission        = Get-Prop $f 'apiPermission'
            required             = Get-Prop $f 'required' -Default $false
        }
        $scopes = Get-Prop $f 'definitionScopes'
        if ($scopes -and ($scopes | Where-Object { $_ -eq 'NODE' -or $_ -eq 'DEVICE' })) {
            $GlobalFields += $fieldObj
        } else {
            $RoleFields += $fieldObj
        }
    }
    Write-Host "  [OK] $($GlobalFields.Count) global field(s), $($RoleFields.Count) role field(s) exported." -ForegroundColor Green
    if ($FieldArr.Count -eq 0) {
        Write-Host "  [!] No fields returned. Ensure custom fields have API Permission" -ForegroundColor Yellow
        Write-Host "      set to 'Read Only' or 'Read/Write' before running this script." -ForegroundColor Yellow
    }
} catch {
    Write-Host "  [!] Could not export custom fields: $_" -ForegroundColor Yellow
    Write-Host "      Continuing..." -ForegroundColor Yellow
}

# =============================================================================
#  STEP 4: Export device roles
# =============================================================================
Write-Host ""
Write-Host "  [4/6] Exporting device roles..." -ForegroundColor Cyan
$Roles = @()
try {
    $RawRoles = Invoke-NinjaApi -BaseUrl $SourceBaseUrl -Headers $SrcHeaders -Path 'roles'
    $RoleArr  = if ($RawRoles -is [array]) { $RawRoles } else { @($RawRoles) }
    $Roles = $RoleArr | ForEach-Object {
        [PSCustomObject]@{
            name        = Get-Prop $_ 'name'
            description = Get-Prop $_ 'description'
            nodeClass   = Get-Prop $_ 'nodeClass'
            custom      = Get-Prop $_ 'custom' -Default $true
        }
    }
    Write-Host "  [OK] $($Roles.Count) device role(s) exported." -ForegroundColor Green
} catch {
    Write-Host "  [!] Could not export device roles: $_" -ForegroundColor Yellow
    Write-Host "      Continuing..." -ForegroundColor Yellow
}

# =============================================================================
#  STEP 5: Export device policies
# =============================================================================
Write-Host ""
Write-Host "  [5/6] Exporting device policies..." -ForegroundColor Cyan
$Policies = @()
try {
    $RawPolicies = Get-AllPages -BaseUrl $SourceBaseUrl -Headers $SrcHeaders -Path 'policies'
    $Policies = $RawPolicies | ForEach-Object {
        [PSCustomObject]@{
            name          = Get-Prop $_ 'name'
            description   = Get-Prop $_ 'description'
            nodeClass     = Get-Prop $_ 'nodeClass'
            parentPolicyId = Get-Prop $_ 'parentPolicyId'  # stored for reference only
            enabled       = Get-Prop $_ 'enabled' -Default $true
            isDefault     = Get-Prop $_ 'default' -Default $false
        }
    }
    Write-Host "  [OK] $($Policies.Count) policy/policies exported." -ForegroundColor Green
    Write-Host "  [i] Note: Policy internal conditions/rules cannot be exported via API." -ForegroundColor Gray
    Write-Host "      Only policy shells (name, class, enabled state) are captured." -ForegroundColor Gray
} catch {
    Write-Host "  [!] Could not export policies: $_" -ForegroundColor Yellow
    Write-Host "      Continuing..." -ForegroundColor Yellow
}

# =============================================================================
#  STEP 6: Save export package
# =============================================================================
Write-Host ""
Write-Host "  [6/6] Saving export package..." -ForegroundColor Cyan

$Timestamp = Get-Date -Format 'yyyyMMdd_HHmmss'
$SourceHost = ([Uri]$SourceBaseUrl).Host -replace '\.', '_'

$ExportPackage = [PSCustomObject]@{
    exportedAt      = (Get-Date -Format 'yyyy-MM-dd HH:mm:ss')
    sourceInstance  = $SourceBaseUrl
    scriptCount     = $Scripts.Count
    globalFieldCount = $GlobalFields.Count
    roleFieldCount  = $RoleFields.Count
    roleCount       = $Roles.Count
    policyCount     = $Policies.Count
    automationScripts = $Scripts
    globalCustomFields = $GlobalFields
    roleCustomFields   = $RoleFields
    deviceRoles        = $Roles
    devicePolicies     = $Policies
}

$JsonPath     = Join-Path $OutputFolder "NinjaExport_${SourceHost}_${Timestamp}.json"
$ReportPath   = Join-Path $OutputFolder "NinjaExport_${SourceHost}_${Timestamp}_Summary.txt"

# Save compressed JSON
$ExportPackage | ConvertTo-Json -Depth 10 -Compress | `
    Set-Content -LiteralPath $JsonPath -Encoding UTF8

# Save human-readable summary report
$Report = @"
NinjaOne Instance Config Export Summary
========================================
Exported At   : $($ExportPackage.exportedAt)
Source        : $SourceBaseUrl

COUNTS
------
Automation Scripts  : $($Scripts.Count)
Global Custom Fields: $($GlobalFields.Count)
Role Custom Fields  : $($RoleFields.Count)
Device Roles        : $($Roles.Count)
Device Policies     : $($Policies.Count)

AUTOMATION SCRIPTS
------------------
$(if ($Scripts.Count -eq 0) { '(none)' } else { ($Scripts | ForEach-Object { "  - $($_.name) [$($_.language)] OS:$($_.operatingSystem) RunAs:$($_.runAs)" }) -join "`n" })

GLOBAL CUSTOM FIELDS (auto-importable)
---------------------------------------
$(if ($GlobalFields.Count -eq 0) { '(none -- check API permissions on your custom fields)' } else { ($GlobalFields | ForEach-Object { "  - $($_.name) [$($_.fieldType)] API:$($_.apiPermission)" }) -join "`n" })

ROLE CUSTOM FIELDS (auto-importable)
--------------------------------------
$(if ($RoleFields.Count -eq 0) { '(none)' } else { ($RoleFields | ForEach-Object { "  - $($_.name) [$($_.fieldType)] API:$($_.apiPermission)" }) -join "`n" })

DEVICE ROLES (auto-importable)
--------------------------------
$(if ($Roles.Count -eq 0) { '(none)' } else { ($Roles | ForEach-Object { "  - $($_.name) [$($_.nodeClass)]" }) -join "`n" })

DEVICE POLICIES (shell only -- conditions must be set manually)
----------------------------------------------------------------
$(if ($Policies.Count -eq 0) { '(none)' } else { ($Policies | ForEach-Object { "  - $($_.name) [$($_.nodeClass)] Enabled:$($_.enabled)" }) -join "`n" })

MANUAL STEPS REQUIRED IN DESTINATION INSTANCE
----------------------------------------------
1. AUTOMATION SCRIPTS:
   The API has no create-script endpoint. For each script above:
   Administration > Library > Automation > Add > New Script
   Copy the script body from the JSON file (automationScripts[].scriptBody).

2. DEVICE POLICIES (conditions/rules):
   Policy shells will be created automatically by the import script.
   However, internal conditions, patch schedules, monitoring thresholds,
   and automation triggers CANNOT be transferred via API. Each policy must
   be configured manually in the destination portal after import.

3. ROLE-TO-POLICY MAPPINGS:
   After policies are created, assign them to device roles manually:
   Administration > Devices > Roles > Edit > Policy tab.

EXPORT FILE
-----------
JSON : $JsonPath
"@

$Report | Set-Content -LiteralPath $ReportPath -Encoding UTF8

Write-Host "  [OK] Export complete." -ForegroundColor Green
Write-Host ""
Write-Host "  ================================================================" -ForegroundColor Green
Write-Host "  EXPORT COMPLETE" -ForegroundColor Green
Write-Host "    Scripts      : $($Scripts.Count)" -ForegroundColor Green
Write-Host "    Global Fields: $($GlobalFields.Count)" -ForegroundColor Green
Write-Host "    Role Fields  : $($RoleFields.Count)" -ForegroundColor Green
Write-Host "    Roles        : $($Roles.Count)" -ForegroundColor Green
Write-Host "    Policies     : $($Policies.Count)" -ForegroundColor Green
Write-Host "  ================================================================" -ForegroundColor Green
Write-Host ""
Write-Host "  JSON export : $JsonPath" -ForegroundColor Cyan
Write-Host "  Summary     : $ReportPath" -ForegroundColor Cyan
Write-Host ""
Write-Host "  NEXT STEP: Run Import-NinjaInstanceConfig.ps1" -ForegroundColor Yellow
Write-Host "  and point it at the JSON file above." -ForegroundColor Yellow
Write-Host ""
