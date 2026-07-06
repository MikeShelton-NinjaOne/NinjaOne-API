#Requires -Version 5.1
<#
.SYNOPSIS
    Reads a CSV exported from LogMeIn, then for each row:
      - Finds the device in NinjaOne by hostname
      - Creates the target organization if it doesn't exist
      - Creates the target location inside that org if it doesn't exist
      - Moves the device to that org and location
      - Sets the device display name to the "Computer description" column value
      - Writes the "Notes" column value to a custom field named "logmeinNotes"

    Stops processing after encountering 3 consecutive empty rows.
    Row 1 is treated as the header and skipped.

.NOTES
    CSV COLUMNS USED:
      "Computer description" -- becomes the device display name
      "Host Name"            -- used to find the device in NinjaOne
      "Organization"         -- target org (created if missing)
      "Location"             -- target location inside that org (created if missing)
      "Notes"                -- written to the "logmeinNotes" custom field

    CUSTOM FIELD PREREQUISITE:
    The "logmeinNotes" custom field must already exist in NinjaOne before running.
    Create it at: Administration > Devices > Global Custom Fields > Add
      Label              : LogMeIn Notes
      Name               : logmeinNotes
      Type               : Text / Multi-line
      Script Permission  : Read/Write
      API Permission     : Read/Write

    API APP SETUP (one-time):
    Go to: Administration > Apps > API > Client App IDs > Add
      Platform      : API Services (Machine-to-Machine)
      Allowed Scopes: monitoring  AND  management
      Redirect URI  : Leave blank
    Click Save. Copy the Client ID and Client Secret.

    REGIONAL URLS:
    US: https://app.ninjarmm.com
    EU: https://eu.ninjarmm.com
    OC: https://oc.ninjarmm.com
    CA: https://ca.ninjarmm.com
#>

# ==============================================================================
#  CONFIGURATION -- Fill in ALL values in this block before running
# ==============================================================================

# Your NinjaOne login URL (no trailing slash)
$BaseUrl       = 'https://<your Login URL>'

# Same URL with /ws/oauth/token appended
$TokenEndpoint = 'https://<your Login URL>/ws/oauth/token'

# From Administration > Apps > API > Client App IDs
$ClientId      = '<Your Client ID>'

# From Administration > Apps > API > Client App IDs (shown once at creation)
$ClientSecret  = '<Your Client Secret>'

# Full path to the CSV file exported from LogMeIn
# Examples:
#   Windows : 'C:\Users\You\Downloads\LMI_X_Ninja.csv'
#   Same dir: Join-Path $PSScriptRoot 'LMI_X_Ninja.csv'
$CsvPath       = '<Path to your CSV file>'

# How many consecutive empty rows before the script stops reading
$EmptyRowLimit = 3

# How many times to retry a failed API call before giving up on that row
$ApiRetryCount = 3

# Seconds to wait between retries
$ApiRetryDelay = 5

# ==============================================================================
#  DO NOT EDIT BELOW THIS LINE
# ==============================================================================

# -- Force TLS 1.2 globally ---------------------------------------------------
# Required for NinjaOne API. PS 5.1 on older Windows defaults to TLS 1.0
# which NinjaOne rejects. This must be set before any web request is made.
try {
    [Net.ServicePointManager]::SecurityProtocol = `
        [Net.ServicePointManager]::SecurityProtocol -bor `
        [Net.SecurityProtocolType]::Tls12
} catch {
    Write-Warning "Could not set TLS 1.2 -- connection to NinjaOne may fail on older OS versions."
}

# -- Set console output encoding to UTF-8 -------------------------------------
# Prevents ? characters on machines with non-UTF-8 console codepages
try {
    [Console]::OutputEncoding = [System.Text.Encoding]::UTF8
    $OutputEncoding            = [System.Text.Encoding]::UTF8
} catch {}

# -- ErrorActionPreference scoped carefully -----------------------------------
# We do NOT set this globally to Stop because that would kill the entire script
# on any non-critical error. Each sensitive block has its own try/catch instead.
$ErrorActionPreference = 'Continue'

# -- Validate config ----------------------------------------------------------
$ConfigErrors = @()
if ($BaseUrl      -like '*<*') { $ConfigErrors += 'Fill in $BaseUrl' }
if ($ClientId     -like '*<*') { $ConfigErrors += 'Fill in $ClientId' }
if ($ClientSecret -like '*<*') { $ConfigErrors += 'Fill in $ClientSecret' }
if ($CsvPath      -like '*<*') { $ConfigErrors += 'Fill in $CsvPath' }
if ($ConfigErrors.Count -gt 0) {
    Write-Host ""
    Write-Host "  [!] CONFIGURATION ERRORS -- please fix before running:" -ForegroundColor Red
    foreach ($e in $ConfigErrors) { Write-Host "      - $e" -ForegroundColor Red }
    Write-Host ""
    exit 1
}

# Resolve CSV path relative to script location when possible
if (-not [System.IO.Path]::IsPathRooted($CsvPath) -and $PSScriptRoot) {
    $CsvPath = Join-Path $PSScriptRoot $CsvPath
}
if (-not (Test-Path -LiteralPath $CsvPath)) {
    Write-Host "  [!] CSV file not found at: $CsvPath" -ForegroundColor Red
    exit 1
}

# -- Safe property helper -----------------------------------------------------
# Reads a property from an API PSObject without throwing if absent
function Get-Prop {
    param([object]$Obj, [string]$Name, [object]$Default = $null)
    if ($null -eq $Obj) { return $Default }
    $p = $Obj.PSObject.Properties[$Name]
    if ($null -eq $p -or $null -eq $p.Value -or
        ($p.Value -is [string] -and $p.Value.Trim() -eq '')) {
        return $Default
    }
    return $p.Value
}

# -- API helper with retry ----------------------------------------------------
function Invoke-NinjaApi {
    param(
        [string]$Path,
        [string]$Method  = 'GET',
        [string]$Body    = $null,
        [int]   $Retries = $script:ApiRetryCount,
        [int]   $Delay   = $script:ApiRetryDelay
    )
    $Attempt = 0
    while ($true) {
        $Attempt++
        $Params = @{
            Uri     = "$script:BaseUrl/v2/$Path"
            Method  = $Method
            Headers = $script:Headers
        }
        if ($Body) {
            $Params.Body        = $Body
            $Params.ContentType = 'application/json'
        }
        try {
            return Invoke-RestMethod @Params
        } catch {
            $sc = $null
            try { $sc = [int]$_.Exception.Response.StatusCode } catch {}

            # Do not retry on definitive client errors (4xx except 429)
            if ($sc -ge 400 -and $sc -lt 500 -and $sc -ne 429) {
                throw "API [$Method /v2/$Path] HTTP $sc -- $_"
            }

            if ($Attempt -gt $Retries) {
                throw "API [$Method /v2/$Path] failed after $Retries retries. Last error: $_"
            }

            Write-Host "    [~] API call failed (HTTP $sc), retrying in ${Delay}s (attempt $Attempt/$Retries)..." -ForegroundColor Yellow
            Start-Sleep -Seconds $Delay
        }
    }
}

# -- Banner -------------------------------------------------------------------
Write-Host ""
Write-Host "  ================================================================" -ForegroundColor Cyan
Write-Host "  NinjaOne LogMeIn Migration Script" -ForegroundColor Cyan
Write-Host "  CSV: $CsvPath" -ForegroundColor Cyan
Write-Host "  ================================================================" -ForegroundColor Cyan
Write-Host ""

# =============================================================================
#  STEP 1: Authenticate
# =============================================================================
Write-Host "  [1/4] Authenticating..." -ForegroundColor Cyan

$script:Headers = $null
try {
    $Token = Invoke-RestMethod `
        -Uri         $TokenEndpoint `
        -Method      POST `
        -ContentType 'application/x-www-form-urlencoded' `
        -Body        @{
            grant_type    = 'client_credentials'
            client_id     = $ClientId
            client_secret = $ClientSecret
            scope         = 'monitoring management'
        }
    $script:Headers = @{
        Authorization  = "Bearer $($Token.access_token)"
        Accept         = 'application/json'
        'Content-Type' = 'application/json'
    }
} catch {
    Write-Host "  [!] Authentication failed: $_" -ForegroundColor Red
    Write-Host "      Check BaseUrl, ClientId, ClientSecret." -ForegroundColor Yellow
    Write-Host "      API app must be API Services (Machine-to-Machine)" -ForegroundColor Yellow
    Write-Host "      with monitoring AND management scopes." -ForegroundColor Yellow
    exit 1
}
Write-Host "  [OK] Authenticated." -ForegroundColor Green

# =============================================================================
#  STEP 2: Load orgs, devices, and locations into memory
# =============================================================================
Write-Host ""
Write-Host "  [2/4] Loading organizations and devices from NinjaOne..." -ForegroundColor Cyan

# -- Paginate orgs ------------------------------------------------------------
$AllOrgs = New-Object System.Collections.ArrayList
$After   = $null
do {
    $QS   = "organizations?pageSize=200"
    if ($After) { $QS += "&after=$After" }
    try {
        $Page  = Invoke-NinjaApi -Path $QS
        $Items = if ($Page -is [array]) { $Page } else { @($Page) }
        foreach ($i in $Items) { [void]$AllOrgs.Add($i) }
        $After = if ($Items.Count -eq 200) { Get-Prop $Items[-1] 'id' } else { $null }
    } catch {
        Write-Host "  [!] Failed to load organizations: $_" -ForegroundColor Red
        exit 1
    }
} while ($After)

# -- Paginate devices ---------------------------------------------------------
$AllDevices = New-Object System.Collections.ArrayList
$After      = $null
do {
    $QS   = "devices?pageSize=200"
    if ($After) { $QS += "&after=$After" }
    try {
        $Page  = Invoke-NinjaApi -Path $QS
        $Items = if ($Page -is [array]) { $Page } else { @($Page) }
        foreach ($i in $Items) { [void]$AllDevices.Add($i) }
        $After = if ($Items.Count -eq 200) { Get-Prop $Items[-1] 'id' } else { $null }
    } catch {
        Write-Host "  [!] Failed to load devices: $_" -ForegroundColor Red
        exit 1
    }
} while ($After)

Write-Host "  [OK] Loaded $($AllOrgs.Count) org(s) and $($AllDevices.Count) device(s)." -ForegroundColor Green

# -- Build lookup dictionaries ------------------------------------------------
$OrgByName    = @{}   # orgName (lowercase) -> org object
$OrgLocations = @{}   # orgId -> hashtable of locationName (lowercase) -> location object
$DevByHost    = @{}   # hostname (lowercase) -> device object

foreach ($o in $AllOrgs) {
    $n = Get-Prop $o 'name'
    if ($n) { $OrgByName[$n.ToLower()] = $o }
}

foreach ($d in $AllDevices) {
    $sn = Get-Prop $d 'systemName'
    $dn = Get-Prop $d 'dnsName'
    if ($sn) { $DevByHost[$sn.ToLower()] = $d }
    if ($dn -and -not $DevByHost.ContainsKey($dn.ToLower())) {
        $DevByHost[$dn.ToLower()] = $d
    }
}

# -- Pre-load locations for known orgs ----------------------------------------
foreach ($o in $AllOrgs) {
    $oid = Get-Prop $o 'id'
    if ($null -eq $oid) { continue }
    $OrgLocations[$oid] = @{}
    try {
        $Locs = Invoke-NinjaApi -Path "organization/$oid/locations"
        if ($Locs) {
            $LocArr = if ($Locs -is [array]) { $Locs } else { @($Locs) }
            foreach ($l in $LocArr) {
                $ln = Get-Prop $l 'name'
                if ($ln) { $OrgLocations[$oid][$ln.ToLower()] = $l }
            }
        }
    } catch {
        # Non-fatal -- some orgs may return 404 on locations
    }
}

# =============================================================================
#  STEP 3: Process CSV rows
# =============================================================================
Write-Host ""
Write-Host "  [3/4] Processing CSV rows..." -ForegroundColor Cyan
Write-Host ""

# Read with UTF-8 encoding explicitly so BOM and special chars in org names
# (accents, apostrophes) are handled correctly
$RawLines = Get-Content -LiteralPath $CsvPath -Encoding UTF8

$EmptyCount  = 0
$RowNum      = 0
$Processed   = 0
$Skipped     = 0
$Errors      = 0
$Results     = New-Object System.Collections.ArrayList

foreach ($Line in $RawLines) {
    $RowNum++

    # Skip header
    if ($RowNum -eq 1) { continue }

    # -- Detect empty rows ----------------------------------------------------
    $IsEmpty = [string]::IsNullOrWhiteSpace($Line) -or
               ($Line -replace '[,\s]', '') -eq ''

    if ($IsEmpty) {
        $EmptyCount++
        Write-Host "  [i] Row $RowNum empty ($EmptyCount of $EmptyRowLimit)." -ForegroundColor Gray
        if ($EmptyCount -ge $EmptyRowLimit) {
            Write-Host "  [i] $EmptyRowLimit consecutive empty rows reached -- stopping." -ForegroundColor Yellow
            break
        }
        continue
    }
    $EmptyCount = 0

    # -- Parse CSV line -------------------------------------------------------
    # Use a proper state-machine parser: handles quoted commas, doubled quotes
    # ("" inside a quoted field), and both LF and CRLF line endings
    $Fields   = New-Object System.Collections.ArrayList
    $InQuote  = $false
    $Current  = New-Object System.Text.StringBuilder
    $CleanLine = $Line.TrimEnd("`r")   # strip carriage return if CRLF

    $Chars = $CleanLine.ToCharArray()
    for ($ci = 0; $ci -lt $Chars.Count; $ci++) {
        $ch = $Chars[$ci]
        if ($ch -eq '"') {
            # Check for escaped quote: "" inside a quoted field
            if ($InQuote -and ($ci + 1) -lt $Chars.Count -and $Chars[$ci + 1] -eq '"') {
                [void]$Current.Append('"')
                $ci++
            } else {
                $InQuote = -not $InQuote
            }
        } elseif ($ch -eq ',' -and -not $InQuote) {
            [void]$Fields.Add($Current.ToString())
            [void]$Current.Clear()
        } else {
            [void]$Current.Append($ch)
        }
    }
    [void]$Fields.Add($Current.ToString())

    # -- Map columns by position ----------------------------------------------
    # Header order: Computer description | Host Name | Organization | Location |
    #               secure name | username | password | Notes
    $ComputerDesc = if ($Fields.Count -gt 0) { $Fields[0].Trim() } else { '' }
    $Hostname     = if ($Fields.Count -gt 1) { $Fields[1].Trim() } else { '' }
    $OrgName      = if ($Fields.Count -gt 2) { $Fields[2].Trim() } else { '' }
    $LocationName = if ($Fields.Count -gt 3) { $Fields[3].Trim() } else { '' }
    $Notes        = if ($Fields.Count -gt 7) { $Fields[7].Trim() } else { '' }

    if ([string]::IsNullOrWhiteSpace($Hostname)) {
        Write-Host "  [~] Row $RowNum -- skipped (no hostname)." -ForegroundColor Gray
        $Skipped++
        continue
    }

    Write-Host "  -- Row $RowNum : $Hostname" -ForegroundColor White

    $RowStatus = 'OK'
    $RowNotes  = New-Object System.Collections.ArrayList

    try {
        # -- Find device ------------------------------------------------------
        $Device = $null
        $HLower = $Hostname.ToLower()

        if ($DevByHost.ContainsKey($HLower)) {
            $Device = $DevByHost[$HLower]
        }
        # Fallback: strip domain suffix (host.domain.com -> host)
        if (-not $Device -and $HLower.Contains('.')) {
            $Short = $HLower.Split('.')[0]
            if ($DevByHost.ContainsKey($Short)) {
                $Device = $DevByHost[$Short]
            }
        }

        if (-not $Device) {
            Write-Host "    [!] '$Hostname' not found in NinjaOne -- skipping." -ForegroundColor Red
            [void]$RowNotes.Add('Device not found')
            $RowStatus = 'NOT FOUND'
            $Skipped++
            [void]$Results.Add([PSCustomObject]@{
                Row = $RowNum; Hostname = $Hostname; Org = $OrgName
                Location = $LocationName; Status = $RowStatus
                Notes = ($RowNotes -join '; ')
            })
            continue
        }

        $DeviceId = Get-Prop $Device 'id'
        Write-Host "    [OK] Device found: $(Get-Prop $Device 'systemName') (ID: $DeviceId)" -ForegroundColor Green

        # -- Resolve or create org --------------------------------------------
        $OrgKey    = $OrgName.ToLower()
        $TargetOrg = if ($OrgByName.ContainsKey($OrgKey)) { $OrgByName[$OrgKey] } else { $null }

        if (-not $TargetOrg) {
            Write-Host "    [i] Org '$OrgName' not found -- creating..." -ForegroundColor Yellow
            $NewOrgBody = @{ name = $OrgName } | ConvertTo-Json -Compress -Depth 2
            $TargetOrg  = Invoke-NinjaApi -Path 'organizations' -Method POST -Body $NewOrgBody
            $OrgByName[$OrgKey]                           = $TargetOrg
            $OrgLocations[(Get-Prop $TargetOrg 'id')]     = @{}
            [void]$RowNotes.Add('Org created')
            Write-Host "    [OK] Org created (ID: $(Get-Prop $TargetOrg 'id'))." -ForegroundColor Green
        } else {
            Write-Host "    [OK] Org: $OrgName (ID: $(Get-Prop $TargetOrg 'id'))." -ForegroundColor Green
        }
        $TargetOrgId = Get-Prop $TargetOrg 'id'

        # -- Resolve or create location ---------------------------------------
        $TargetLocationId = $null
        if (-not [string]::IsNullOrWhiteSpace($LocationName)) {
            $LocKey = $LocationName.ToLower()

            if (-not $OrgLocations.ContainsKey($TargetOrgId)) {
                $OrgLocations[$TargetOrgId] = @{}
            }

            $TargetLoc = if ($OrgLocations[$TargetOrgId].ContainsKey($LocKey)) {
                $OrgLocations[$TargetOrgId][$LocKey]
            } else { $null }

            if (-not $TargetLoc) {
                Write-Host "    [i] Location '$LocationName' not found -- creating..." -ForegroundColor Yellow
                $NewLocBody = @{ name = $LocationName } | ConvertTo-Json -Compress -Depth 2
                $TargetLoc  = Invoke-NinjaApi -Path "organization/$TargetOrgId/locations" -Method POST -Body $NewLocBody
                $OrgLocations[$TargetOrgId][$LocKey] = $TargetLoc
                [void]$RowNotes.Add('Location created')
                Write-Host "    [OK] Location created (ID: $(Get-Prop $TargetLoc 'id'))." -ForegroundColor Green
            } else {
                Write-Host "    [OK] Location: $LocationName (ID: $(Get-Prop $TargetLoc 'id'))." -ForegroundColor Green
            }
            $TargetLocationId = Get-Prop $TargetLoc 'id'
        }

        # -- Move device and set display name ---------------------------------
        $PatchPayload = [ordered]@{ organizationId = $TargetOrgId }
        if ($null -ne $TargetLocationId) {
            $PatchPayload['locationId'] = $TargetLocationId
        }
        if (-not [string]::IsNullOrWhiteSpace($ComputerDesc)) {
            $PatchPayload['displayName'] = $ComputerDesc
        }
        $PatchBody = $PatchPayload | ConvertTo-Json -Compress -Depth 2

        Invoke-NinjaApi -Path "device/$DeviceId" -Method PATCH -Body $PatchBody | Out-Null

        $MoveMsg = "    [OK] Moved to '$OrgName'"
        if ($TargetLocationId) { $MoveMsg += ", location '$LocationName'" }
        Write-Host $MoveMsg -ForegroundColor Green
        if (-not [string]::IsNullOrWhiteSpace($ComputerDesc)) {
            Write-Host "    [OK] Display name set: $ComputerDesc" -ForegroundColor Green
        }

        # -- Write notes to custom field --------------------------------------
        if (-not [string]::IsNullOrWhiteSpace($Notes)) {
            $CfBody = @{ logmeinNotes = $Notes } | ConvertTo-Json -Compress -Depth 2
            try {
                Invoke-NinjaApi -Path "device/$DeviceId/custom-fields" -Method PATCH -Body $CfBody | Out-Null
                Write-Host "    [OK] Notes written to 'logmeinNotes'." -ForegroundColor Green
            } catch {
                Write-Host "    [!] Could not write to 'logmeinNotes': $_" -ForegroundColor Yellow
                Write-Host "        Check the field exists with API Read/Write permission." -ForegroundColor Yellow
                [void]$RowNotes.Add('Custom field write failed')
            }
        } else {
            Write-Host "    [i] No notes for this device." -ForegroundColor Gray
        }

        $Processed++

    } catch {
        Write-Host "    [!] Error on row $RowNum : $_" -ForegroundColor Red
        $RowStatus = 'ERROR'
        [void]$RowNotes.Add("Error: $_")
        $Errors++
    }

    [void]$Results.Add([PSCustomObject]@{
        Row      = $RowNum
        Hostname = $Hostname
        Org      = $OrgName
        Location = $LocationName
        Status   = $RowStatus
        Notes    = ($RowNotes -join '; ')
    })
    Write-Host ""
}

# =============================================================================
#  STEP 4: Summary and results export
# =============================================================================
Write-Host ""
Write-Host "  ================================================================" -ForegroundColor Green
Write-Host "  COMPLETE" -ForegroundColor Green
Write-Host "    Rows processed : $($RowNum - 1)" -ForegroundColor Green
Write-Host "    Migrated OK    : $Processed"     -ForegroundColor Green

if ($Skipped -gt 0) {
    Write-Host "    Skipped        : $Skipped" -ForegroundColor Yellow
} else {
    Write-Host "    Skipped        : $Skipped" -ForegroundColor Green
}
if ($Errors -gt 0) {
    Write-Host "    Errors         : $Errors" -ForegroundColor Red
} else {
    Write-Host "    Errors         : $Errors" -ForegroundColor Green
}
Write-Host "  ================================================================" -ForegroundColor Green
Write-Host ""

if ($Results.Count -gt 0) {
    Write-Host "  Per-row results:" -ForegroundColor Cyan
    $Results | Format-Table -AutoSize Row, Hostname, Org, Location, Status, Notes
}

# Save results CSV next to input file using a safe path resolution
try {
    $CsvDir  = [System.IO.Path]::GetDirectoryName([System.IO.Path]::GetFullPath($CsvPath))
    $OutFile = "NinjaMigration_Results_$(Get-Date -Format 'yyyyMMdd_HHmmss').csv"
    $OutPath = Join-Path $CsvDir $OutFile
    $Results | Export-Csv -LiteralPath $OutPath -NoTypeInformation -Encoding UTF8
    Write-Host "  Results saved to: $OutPath" -ForegroundColor Cyan
} catch {
    Write-Host "  [!] Could not save results CSV: $_" -ForegroundColor Yellow
    Write-Host "  Printing results to console instead:" -ForegroundColor Yellow
    $Results | ConvertTo-Csv -NoTypeInformation | Write-Host
}
Write-Host ""
