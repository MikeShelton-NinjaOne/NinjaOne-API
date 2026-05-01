#Requires -Version 5.1
<#
# This will not be supported by ninja support and is meant to be an example.
.SYNOPSIS
    Generates per-organization patch status reports for Windows devices
    in a NinjaOne account.

.DESCRIPTION
    Authenticates to the NinjaOne API using the OAuth 2.0 Client Credentials
    grant, enumerates organizations and devices, filters to Windows-only
    node classes, pulls OS and software patch lists for each included device,
    categorizes patches by status, and writes one self-contained HTML file
    per organization plus a linking index.html.

    Designed for Windows PowerShell 5.1. No external modules required.

.NOTES
    - Region: Canada (ca.ninjarmm.com)
    - Credentials are embedded below. Rotate them after testing and move
      to a secure store (Credential Manager, env var, or secret vault) for
      production use.
    - Allowed Windows node classes:
        WINDOWS_WORKSTATION, WINDOWS_SERVER, WINDOWS_LAPTOP, WINDOWS_DESKTOP

.EXAMPLE
    PS C:\> .\Generate-NinjaOnePatchReportByOrg.ps1

    Builds the report set in a new timestamped folder under the current
    directory and opens index.html.
#>

# ---------------------------------------------------------------------------
# Configuration
# ---------------------------------------------------------------------------
#Make sure to change the base URL to whatever URL you normally log into
$Script:Config = @{
    BaseUrl          = 'https://<your Login URL>'
    TokenEndpoint    = 'https://<your login URL>/ws/oauth/token'
    ClientId         = '<Your Client ID>'
    ClientSecret     = '<Your Client Secert>'
    Scope            = 'monitoring'
    PageSize         = 1000
    MaxRetries       = 5
    AllowedNodeClass = @(
        'WINDOWS_WORKSTATION',
        'WINDOWS_SERVER',
        'WINDOWS_LAPTOP',
        'WINDOWS_DESKTOP'
    )
}

# Force TLS 1.2 (required by NinjaOne)
[Net.ServicePointManager]::SecurityProtocol = [Net.SecurityProtocolType]::Tls12

# Ensure System.Web is available for URL/HTML encoding
Add-Type -AssemblyName System.Web -ErrorAction SilentlyContinue

# ---------------------------------------------------------------------------
# Reporting window: previous calendar month, local time.
# Example: if run any time in April, window = March 1 00:00:00 .. March 31 23:59:59.
# ---------------------------------------------------------------------------
$Script:ReportWindow = & {
    $now           = Get-Date
    $firstOfThis   = Get-Date -Year $now.Year -Month $now.Month -Day 1 -Hour 0 -Minute 0 -Second 0 -Millisecond 0
    $endExclusive  = $firstOfThis                          # start of current month (exclusive upper bound)
    $startInclusive = $firstOfThis.AddMonths(-1)           # start of previous month
    $endInclusive   = $endExclusive.AddSeconds(-1)         # last second of previous month (for display)

    # Epoch seconds for the NinjaOne API (UTC).
    $epoch = [DateTime]::SpecifyKind([DateTime]'1970-01-01', [DateTimeKind]::Utc)
    $startEpoch = [int64]([DateTime]::SpecifyKind($startInclusive, [DateTimeKind]::Local).ToUniversalTime() - $epoch).TotalSeconds
    $endEpoch   = [int64]([DateTime]::SpecifyKind($endExclusive,   [DateTimeKind]::Local).ToUniversalTime() - $epoch).TotalSeconds

    [pscustomobject]@{
        StartLocal    = $startInclusive
        EndLocalIncl  = $endInclusive
        StartEpochUtc = $startEpoch
        EndEpochUtc   = $endEpoch           # exclusive upper bound (start of current month)
        Label         = "$($startInclusive.ToString('yyyy-MM-dd')) to $($endInclusive.ToString('yyyy-MM-dd'))"
    }
}

# Runtime state
$Script:BearerToken = $null

# Diagnostics: populated by Get-DevicePatches so we can tell at the end how
# many install records the API returned vs how many we kept after filtering.
$Script:InstallDiagnostics = [ordered]@{
    OsRaw           = 0
    OsKept          = 0
    SwRaw           = 0
    SwKept          = 0
    SampleOsRecord  = $null
    SampleSwRecord  = $null
}


# ---------------------------------------------------------------------------
# Helpers
# ---------------------------------------------------------------------------

function ConvertTo-SafeValue {
    <#
        Returns $null for $null, empty strings, or whitespace-only strings.
        Otherwise returns the input unchanged. Used to normalize API
        responses so "missing" and "blank" both render as em-dash.
    #>
    param($Value)
    if ($null -eq $Value) { return $null }
    if ($Value -is [string]) {
        if ([string]::IsNullOrWhiteSpace($Value)) { return $null }
    }
    return $Value
}

function Format-HtmlValue {
    <#
        HTML-escapes a value. Null / empty becomes an em-dash so the report
        never shows the literal string "null".
    #>
    param($Value)
    $safe = ConvertTo-SafeValue $Value
    if ($null -eq $safe) { return '&mdash;' }
    return [System.Web.HttpUtility]::HtmlEncode([string]$safe)
}

function Format-EpochOrIso {
    <#
        NinjaOne timestamps come as Unix epoch seconds in some endpoints and
        ISO-8601 strings in others. Normalize both to a local-time string.
        Returns $null for missing values.
    #>
    param($Value)
    $safe = ConvertTo-SafeValue $Value
    if ($null -eq $safe) { return $null }

    $asDouble = 0.0
    if ([double]::TryParse([string]$safe, [ref]$asDouble)) {
        try {
            $origin = [DateTime]::SpecifyKind([DateTime]'1970-01-01', [DateTimeKind]::Utc)
            return $origin.AddSeconds($asDouble).ToLocalTime().ToString('yyyy-MM-dd HH:mm:ss')
        } catch {
            return [string]$safe
        }
    }

    $parsed = [DateTime]::MinValue
    if ([DateTime]::TryParse([string]$safe, [ref]$parsed)) {
        return $parsed.ToLocalTime().ToString('yyyy-MM-dd HH:mm:ss')
    }

    return [string]$safe
}

function ConvertTo-DateTimeOrNull {
    <#
        Parses a value that may be Unix epoch seconds (number or numeric
        string) or an ISO-8601 / RFC date string into a UTC DateTime.
        Returns $null if the value is missing or unparseable.
    #>
    param($Value)
    $safe = ConvertTo-SafeValue $Value
    if ($null -eq $safe) { return $null }

    $asDouble = 0.0
    if ([double]::TryParse([string]$safe, [ref]$asDouble)) {
        try {
            $origin = [DateTime]::SpecifyKind([DateTime]'1970-01-01', [DateTimeKind]::Utc)
            return $origin.AddSeconds($asDouble)
        } catch { return $null }
    }

    $parsed = [DateTime]::MinValue
    if ([DateTime]::TryParse([string]$safe, [ref]$parsed)) {
        return $parsed.ToUniversalTime()
    }
    return $null
}

function Get-PatchRelevantDate {
    <#
        Pulls the most-meaningful timestamp from a patch record for the
        purpose of date filtering. Priority differs by source:
          Installs: installedAt is the authoritative field per the NinjaOne
                    API ("Get Device OS/Software Patch Installs") -- it's
                    what `installedAfter` / `installedBefore` filter on.
                    Fall back to other date-ish fields for safety.
          Open:     lastChange / timestamp first (when the status last
                    changed), then detectedOn, finally releaseDate.
        Returns a UTC DateTime or $null.
    #>
    param(
        $Patch,
        [string]$SourceHint
    )

    $candidateFields = if ($SourceHint -eq 'Installs') {
        @('installedAt','installedOn','installDate','timestamp','lastChange','detectedOn','releaseDate')
    } else {
        @('timestamp','lastChange','installedAt','installedOn','detectedOn','releaseDate')
    }

    foreach ($field in $candidateFields) {
        if ($Patch.PSObject.Properties[$field]) {
            $dt = ConvertTo-DateTimeOrNull $Patch.$field
            if ($null -ne $dt) { return $dt }
        }
    }
    return $null
}

function Test-InReportWindow {
    <#
        Returns $true if the patch's relevant date falls within the
        configured report window (previous calendar month). Returns
        $false if the record is outside the window. If the record has
        no usable date at all, returns $true for Open (pending/approved)
        records -- we don't want to accidentally hide a current
        outstanding patch just because it lacks a timestamp -- and
        $false for Install records, where a missing date is suspicious.
    #>
    param(
        $Patch,
        [string]$SourceHint
    )
    $dt = Get-PatchRelevantDate -Patch $Patch -SourceHint $SourceHint
    if ($null -eq $dt) {
        return ($SourceHint -ne 'Installs')
    }
    $startUtc = [DateTime]::SpecifyKind($Script:ReportWindow.StartLocal, [DateTimeKind]::Local).ToUniversalTime()
    $endUtc   = [DateTime]::SpecifyKind($Script:ReportWindow.EndLocalIncl.AddSeconds(1), [DateTimeKind]::Local).ToUniversalTime()
    return ($dt -ge $startUtc -and $dt -lt $endUtc)
}

function Get-SafeFileName {
    <#
        Strips/replaces characters invalid in Windows filenames, trims
        trailing dots/spaces, and collapses whitespace. Returns 'Unnamed'
        if the result is empty.
    #>
    param([string]$Name)

    if ([string]::IsNullOrWhiteSpace($Name)) { return 'Unnamed' }

    $invalid = [System.IO.Path]::GetInvalidFileNameChars()
    $sb = New-Object System.Text.StringBuilder
    foreach ($ch in $Name.ToCharArray()) {
        if ($invalid -contains $ch) {
            [void]$sb.Append('_')
        } else {
            [void]$sb.Append($ch)
        }
    }
    $out = $sb.ToString()
    # Windows-specific characters that aren't always in GetInvalidFileNameChars
    foreach ($c in '\','/','*','?','"','<','>','|',':') {
        $out = $out.Replace($c, '_')
    }
    # Collapse whitespace and trim trailing dots/spaces
    $out = ($out -replace '\s+', ' ').Trim()
    $out = $out.TrimEnd('.', ' ')
    if ([string]::IsNullOrWhiteSpace($out)) { return 'Unnamed' }
    # Reserved Windows device names
    $reserved = @('CON','PRN','AUX','NUL','COM1','COM2','COM3','COM4','COM5',
                  'COM6','COM7','COM8','COM9','LPT1','LPT2','LPT3','LPT4',
                  'LPT5','LPT6','LPT7','LPT8','LPT9')
    if ($reserved -contains $out.ToUpperInvariant()) { $out = "_$out" }
    return $out
}


# ---------------------------------------------------------------------------
# Authentication
# ---------------------------------------------------------------------------

function Get-NinjaToken {
    [CmdletBinding()]
    param()

    Write-Host "Requesting OAuth token from $($Script:Config.TokenEndpoint) ..." -ForegroundColor Cyan

    $bodyParts = @(
        "grant_type=client_credentials",
        "client_id=$([System.Web.HttpUtility]::UrlEncode($Script:Config.ClientId))",
        "client_secret=$([System.Web.HttpUtility]::UrlEncode($Script:Config.ClientSecret))",
        "scope=$([System.Web.HttpUtility]::UrlEncode($Script:Config.Scope))"
    )
    $body = $bodyParts -join '&'

    try {
        $response = Invoke-RestMethod `
            -Uri $Script:Config.TokenEndpoint `
            -Method Post `
            -ContentType 'application/x-www-form-urlencoded' `
            -Body $body `
            -UseBasicParsing `
            -ErrorAction Stop

        if (-not $response.access_token) {
            throw "Token endpoint returned no access_token."
        }

        Write-Host "Authentication successful. Token type: $($response.token_type)" -ForegroundColor Green
        return [string]$response.access_token
    }
    catch {
        $detail = $_.Exception.Message
        if ($_.Exception.Response) {
            try {
                $stream = $_.Exception.Response.GetResponseStream()
                $reader = New-Object System.IO.StreamReader($stream)
                $detail = "$detail`n$($reader.ReadToEnd())"
            } catch { }
        }
        Write-Host "FATAL: Authentication failed. $detail" -ForegroundColor Red
        exit 1
    }
}


# ---------------------------------------------------------------------------
# Generic API caller with retry / backoff
# ---------------------------------------------------------------------------

function Invoke-NinjaApi {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $true)][string]$Path,
        [hashtable]$Query
    )

    $url = "$($Script:Config.BaseUrl)$Path"
    if ($Query -and $Query.Count -gt 0) {
        $pairs = foreach ($k in $Query.Keys) {
            $v = $Query[$k]
            if ($null -eq $v) { continue }
            "$([System.Web.HttpUtility]::UrlEncode($k))=$([System.Web.HttpUtility]::UrlEncode([string]$v))"
        }
        if ($pairs) { $url = "$url`?$($pairs -join '&')" }
    }

    $headers = @{
        'Authorization' = "Bearer $Script:BearerToken"
        'Accept'        = 'application/json'
    }

    $attempt = 0
    while ($true) {
        $attempt++
        try {
            $resp = Invoke-RestMethod `
                -Uri $url `
                -Method Get `
                -Headers $headers `
                -UseBasicParsing `
                -ErrorAction Stop
            return $resp
        }
        catch {
            $status = $null
            $retryAfter = $null
            if ($_.Exception.Response) {
                try { $status = [int]$_.Exception.Response.StatusCode } catch { }
                try { $retryAfter = $_.Exception.Response.Headers['Retry-After'] } catch { }
            }

            # 429 -> back off and retry
            if ($status -eq 429 -and $attempt -le $Script:Config.MaxRetries) {
                $delay = 0
                if ($retryAfter) {
                    $parsed = 0
                    if ([int]::TryParse([string]$retryAfter, [ref]$parsed)) { $delay = $parsed }
                }
                if ($delay -le 0) { $delay = [int][Math]::Pow(2, $attempt) }
                Write-Host "Rate limited (429). Retrying in $delay s (attempt $attempt/$($Script:Config.MaxRetries))." -ForegroundColor Yellow
                Start-Sleep -Seconds $delay
                continue
            }

            # 401 -> refresh token and retry (once or twice)
            if ($status -eq 401 -and $attempt -le 2) {
                Write-Host "Got 401. Refreshing token and retrying." -ForegroundColor Yellow
                $Script:BearerToken = Get-NinjaToken
                $headers['Authorization'] = "Bearer $Script:BearerToken"
                continue
            }

            # 5xx -> transient; back off
            if ($status -ge 500 -and $status -lt 600 -and $attempt -le $Script:Config.MaxRetries) {
                $delay = [int][Math]::Pow(2, $attempt)
                Write-Host "Server error $status. Retrying in $delay s (attempt $attempt/$($Script:Config.MaxRetries))." -ForegroundColor Yellow
                Start-Sleep -Seconds $delay
                continue
            }

            throw
        }
    }
}


# ---------------------------------------------------------------------------
# Organizations
# ---------------------------------------------------------------------------

function Get-AllOrganizations {
    [CmdletBinding()]
    param()

    Write-Host "Fetching organizations ..." -ForegroundColor Cyan

    $all = New-Object System.Collections.Generic.List[object]
    $after = $null
    $pageNum = 0

    while ($true) {
        $pageNum++
        $q = @{ pageSize = $Script:Config.PageSize }
        if ($null -ne $after) { $q['after'] = $after }

        try {
            $page = Invoke-NinjaApi -Path '/v2/organizations' -Query $q
        } catch {
            Write-Host "Failed to fetch organizations page $pageNum`: $($_.Exception.Message)" -ForegroundColor Red
            throw
        }

        if ($null -eq $page) { break }
        $items = @($page)
        if ($items.Count -eq 0) { break }

        foreach ($o in $items) { $all.Add($o) | Out-Null }
        Write-Host "  Orgs page $pageNum`: $($items.Count). Running total: $($all.Count)" -ForegroundColor Gray

        if ($items.Count -lt $Script:Config.PageSize) { break }

        $lastId = $items[-1].id
        if ($null -eq $lastId -or $lastId -eq $after) { break }
        $after = $lastId
    }

    Write-Host "Total organizations: $($all.Count)" -ForegroundColor Green

    $map = @{}
    foreach ($o in $all) {
        if ($null -ne $o.id) {
            $name = ConvertTo-SafeValue $o.name
            $map[[string]$o.id] = $name
        }
    }
    return $map
}


# ---------------------------------------------------------------------------
# Device enumeration
# ---------------------------------------------------------------------------

function Get-AllDevices {
    [CmdletBinding()]
    param()

    Write-Host "Fetching devices (page size $($Script:Config.PageSize)) ..." -ForegroundColor Cyan

    $all = New-Object System.Collections.Generic.List[object]
    $after = $null
    $pageNum = 0

    while ($true) {
        $pageNum++
        $q = @{ pageSize = $Script:Config.PageSize }
        if ($null -ne $after) { $q['after'] = $after }

        try {
            $page = Invoke-NinjaApi -Path '/v2/devices' -Query $q
        } catch {
            Write-Host "Failed to fetch devices page $pageNum`: $($_.Exception.Message)" -ForegroundColor Red
            throw
        }

        if ($null -eq $page) { break }
        $items = @($page)
        if ($items.Count -eq 0) { break }

        foreach ($d in $items) { $all.Add($d) | Out-Null }
        Write-Host "  Devices page $pageNum`: $($items.Count). Running total: $($all.Count)" -ForegroundColor Gray

        if ($items.Count -lt $Script:Config.PageSize) { break }

        $lastId = $items[-1].id
        if ($null -eq $lastId -or $lastId -eq $after) { break }
        $after = $lastId
    }

    Write-Host "Total devices retrieved (before filter): $($all.Count)" -ForegroundColor Green
    return $all.ToArray()
}


# ---------------------------------------------------------------------------
# Patch collection per device
# ---------------------------------------------------------------------------

function Get-DevicePatches {
    <#
        NinjaOne splits patch data across FOUR endpoints:
          /v2/device/{id}/os-patches            -> pending / approved / failed / rejected OS patches
          /v2/device/{id}/software-patches      -> pending / approved / failed / rejected third-party patches
          /v2/device/{id}/os-patch-installs     -> installation history for OS patches
          /v2/device/{id}/software-patch-installs -> installation history for third-party patches

        The first two never include INSTALLED patches, which is why the earlier
        version of this script always showed zero installs. We hit all four
        and tag each record with its source category so downstream code
        doesn't have to guess.
    #>
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $true)]$Device
    )

    $deviceId = $Device.id
    $result = [ordered]@{
        OsPatches          = @()   # pending/approved/failed/rejected OS
        SoftwarePatches    = @()   # pending/approved/failed/rejected software
        OsInstalls         = @()   # installed OS history
        SoftwareInstalls   = @()   # installed software history
    }

    # The open-patch endpoints don't support date params, so we fetch them
    # unfiltered and apply the window client-side. The install endpoints
    # DO support `installedAfter` / `installedBefore`, so we narrow the
    # server-side response to the exact window (plus a small buffer) to
    # avoid pulling years of install history.
    $installedAfter  = $Script:ReportWindow.StartEpochUtc
    $installedBefore = $Script:ReportWindow.EndEpochUtc

    try {
        $osResp = Invoke-NinjaApi -Path "/v2/device/$deviceId/os-patches"
        if ($osResp) {
            $result.OsPatches = @($osResp | Where-Object { Test-InReportWindow -Patch $_ -SourceHint 'Open' })
        }
    } catch {
        Write-Host "  [Device $deviceId] OS patches fetch failed: $($_.Exception.Message)" -ForegroundColor Yellow
    }

    try {
        $swResp = Invoke-NinjaApi -Path "/v2/device/$deviceId/software-patches"
        if ($swResp) {
            $result.SoftwarePatches = @($swResp | Where-Object { Test-InReportWindow -Patch $_ -SourceHint 'Open' })
        }
    } catch {
        Write-Host "  [Device $deviceId] Software patches fetch failed: $($_.Exception.Message)" -ForegroundColor Yellow
    }

    try {
        $osInst = Invoke-NinjaApi -Path "/v2/device/$deviceId/os-patch-installs" `
                                  -Query @{
                                      installedAfter  = $installedAfter
                                      installedBefore = $installedBefore
                                  }
        if ($osInst) {
            $raw = @($osInst)
            $kept = @($raw | Where-Object { Test-InReportWindow -Patch $_ -SourceHint 'Installs' })
            $Script:InstallDiagnostics.OsRaw  += $raw.Count
            $Script:InstallDiagnostics.OsKept += $kept.Count
            if ($raw.Count -gt 0 -and -not $Script:InstallDiagnostics.SampleOsRecord) {
                $Script:InstallDiagnostics.SampleOsRecord = $raw[0]
            }
            $result.OsInstalls = $kept
        }
    } catch {
        Write-Host "  [Device $deviceId] OS patch installs fetch failed: $($_.Exception.Message)" -ForegroundColor Yellow
    }

    try {
        $swInst = Invoke-NinjaApi -Path "/v2/device/$deviceId/software-patch-installs" `
                                  -Query @{
                                      installedAfter  = $installedAfter
                                      installedBefore = $installedBefore
                                  }
        if ($swInst) {
            $raw = @($swInst)
            $kept = @($raw | Where-Object { Test-InReportWindow -Patch $_ -SourceHint 'Installs' })
            $Script:InstallDiagnostics.SwRaw  += $raw.Count
            $Script:InstallDiagnostics.SwKept += $kept.Count
            if ($raw.Count -gt 0 -and -not $Script:InstallDiagnostics.SampleSwRecord) {
                $Script:InstallDiagnostics.SampleSwRecord = $raw[0]
            }
            $result.SoftwareInstalls = $kept
        }
    } catch {
        Write-Host "  [Device $deviceId] Software patch installs fetch failed: $($_.Exception.Message)" -ForegroundColor Yellow
    }

    return [pscustomobject]$result
}

function Get-PatchCategory {
    <#
        Categorize a patch record pulled from one of the four patch endpoints.
        Categories returned: Pending, Approved, Installed, or Other.

        The /os-patches and /software-patches endpoints use `status` values
        like PENDING, APPROVED, MANUAL, AUTO_APPROVED, FAILED, REJECTED.
        The /*-patch-installs endpoints use an `installResult` / `status`
        value of INSTALLED (or sometimes FAILED). We honor an explicit
        SourceHint parameter so rows known to come from an installs
        endpoint are always categorized as Installed without ambiguity.
    #>
    param(
        $Patch,
        [string]$SourceHint
    )

    if ($SourceHint -eq 'Installs') {
        # Records from the install endpoints represent install attempts.
        # Per the NinjaOne API, a `status` field of INSTALLED or FAILED
        # indicates the outcome. Count INSTALLED as Installed; treat
        # FAILED (or any other non-INSTALLED value) as Other so failed
        # installs don't inflate the Installed count.
        $s  = ConvertTo-SafeValue $Patch.status
        $ir = ConvertTo-SafeValue $Patch.installResult
        $combined = @()
        if ($s)  { $combined += ([string]$s).ToUpperInvariant() }
        if ($ir) { $combined += ([string]$ir).ToUpperInvariant() }

        foreach ($v in $combined) {
            if ($v -eq 'FAILED' -or $v -eq 'FAILURE') { return 'Other' }
        }
        foreach ($v in $combined) {
            if ($v -eq 'INSTALLED' -or $v -eq 'SUCCESS' -or $v -eq 'SUCCEEDED') { return 'Installed' }
        }
        # No status field present but the record came from an install
        # endpoint -- assume installed (the endpoint itself denotes history).
        return 'Installed'
    }

    $status        = ConvertTo-SafeValue $Patch.status
    $installResult = ConvertTo-SafeValue $Patch.installResult

    if ($installResult) {
        $ir = ([string]$installResult).ToUpperInvariant()
        if ($ir -eq 'INSTALLED' -or $ir -eq 'SUCCESS' -or $ir -eq 'SUCCEEDED') {
            return 'Installed'
        }
    }

    if ($status) {
        $s = ([string]$status).ToUpperInvariant()
        switch -Regex ($s) {
            '^PENDING'   { return 'Pending' }
            '^APPROVED'  { return 'Approved' }
            '^MANUAL'    { return 'Approved' }   # MANUAL_APPROVAL
            '^AUTO'      { return 'Approved' }   # AUTO_APPROVED
            '^INSTALLED' { return 'Installed' }
        }
    }

    return 'Other'
}


# ---------------------------------------------------------------------------
# Shared CSS / JS
# ---------------------------------------------------------------------------

function Get-ReportCss {
    return @'
    * { box-sizing: border-box; }
    body {
        font-family: -apple-system, BlinkMacSystemFont, "Segoe UI", Roboto,
                     "Helvetica Neue", Arial, sans-serif;
        margin: 0; padding: 24px; background: #f5f6f8; color: #222;
    }
    a { color: #1a73e8; text-decoration: none; }
    a:hover { text-decoration: underline; }
    h1 { margin: 0 0 4px 0; font-size: 24px; }
    h2 { margin: 24px 0 12px 0; font-size: 18px; }
    .meta { color: #666; font-size: 13px; margin-bottom: 20px; }
    .back { display: inline-block; margin-bottom: 12px; font-size: 13px; }
    .summary {
        display: grid; grid-template-columns: repeat(auto-fit, minmax(180px, 1fr));
        gap: 12px; margin-bottom: 24px;
    }
    .card {
        background: #fff; border-radius: 8px; padding: 16px;
        box-shadow: 0 1px 3px rgba(0,0,0,0.06);
    }
    .card .label { font-size: 12px; text-transform: uppercase; color: #888; letter-spacing: 0.5px; }
    .card .value { font-size: 28px; font-weight: 600; margin-top: 4px; }
    .card.pending   .value { color: #c0392b; }
    .card.approved  .value { color: #d68910; }
    .card.installed .value { color: #1e8449; }
    .filter-bar {
        position: sticky; top: 0; background: #f5f6f8; padding: 8px 0 16px 0;
        z-index: 20; display: flex; gap: 12px; align-items: center;
    }
    .filter-bar input[type=text] {
        flex: 1; padding: 10px 12px; border: 1px solid #ccc; border-radius: 6px;
        font-size: 14px;
    }
    details.device {
        background: #fff; border-radius: 8px; margin-bottom: 10px;
        box-shadow: 0 1px 3px rgba(0,0,0,0.06); overflow: hidden;
    }
    details.device > summary {
        cursor: pointer; padding: 12px 16px; font-weight: 600;
        display: flex; justify-content: space-between; align-items: center;
        list-style: none;
    }
    details.device > summary::-webkit-details-marker { display: none; }
    details.device > summary::before {
        content: '>'; margin-right: 8px; color: #888; transition: transform 0.15s;
        display: inline-block; font-family: monospace;
    }
    details.device[open] > summary::before { transform: rotate(90deg); }
    .device-meta { font-weight: 400; color: #666; font-size: 13px; }
    .device-body { padding: 0 16px 16px 16px; }
    .badges { display: flex; gap: 6px; }
    .badge {
        padding: 3px 8px; border-radius: 12px; font-size: 11px; font-weight: 600;
        white-space: nowrap;
    }
    .badge.pending   { background: #fdecea; color: #c0392b; }
    .badge.approved  { background: #fef5e7; color: #b9770e; }
    .badge.installed { background: #e8f6ef; color: #1e8449; }
    .badge.other     { background: #eceff1; color: #546e7a; }
    table.patches, table.index {
        width: 100%; border-collapse: collapse; margin-top: 10px;
        font-size: 13px; background: #fff;
    }
    table.patches thead th, table.index thead th {
        position: sticky; top: 0; background: #f0f2f5; text-align: left;
        padding: 8px 10px; border-bottom: 2px solid #d9dce0; font-weight: 600;
    }
    table.patches tbody td, table.index tbody td {
        padding: 8px 10px; border-bottom: 1px solid #eee; vertical-align: top;
    }
    table.patches tbody tr:nth-child(even),
    table.index   tbody tr:nth-child(even) { background: #fafbfc; }
    table.index tfoot td {
        padding: 10px; font-weight: 700; border-top: 2px solid #d9dce0;
        background: #f0f2f5;
    }
    .num { text-align: right; font-variant-numeric: tabular-nums; }
    .section-title { margin: 14px 0 4px 0; font-size: 14px; font-weight: 600; color: #333; }
    .empty { color: #888; font-style: italic; font-size: 13px; }
    .hidden { display: none !important; }
'@
}

function Get-FilterJs {
    return @'
    (function() {
        var input = document.getElementById('hostFilter');
        if (!input) return;
        input.addEventListener('input', function() {
            var q = input.value.trim().toLowerCase();
            var devices = document.querySelectorAll('details.device');
            devices.forEach(function(d) {
                var h = (d.getAttribute('data-hostname') || '').toLowerCase();
                if (q === '' || h.indexOf(q) !== -1) {
                    d.classList.remove('hidden');
                } else {
                    d.classList.add('hidden');
                }
            });
        });
    })();
'@
}


# ---------------------------------------------------------------------------
# Per-organization HTML report
# ---------------------------------------------------------------------------

function New-OrganizationReport {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $true)][string]$OrgDisplayName,
        [Parameter(Mandatory = $true)]$Devices,
        [Parameter(Mandatory = $true)]$DevicePatchMap,
        [Parameter(Mandatory = $true)][string]$OutputPath,
        [Parameter(Mandatory = $true)][string]$GeneratedAt
    )

    # Aggregate totals for this org
    $totalPending   = 0
    $totalApproved  = 0
    $totalInstalled = 0

    foreach ($d in $Devices) {
        $entry = $DevicePatchMap[[string]$d.id]
        if ($null -eq $entry) { continue }
        foreach ($p in @($entry.OsPatches) + @($entry.SoftwarePatches)) {
            switch (Get-PatchCategory -Patch $p -SourceHint 'Open') {
                'Pending'   { $totalPending++ }
                'Approved'  { $totalApproved++ }
                'Installed' { $totalInstalled++ }
            }
        }
        foreach ($p in @($entry.OsInstalls) + @($entry.SoftwareInstalls)) {
            switch (Get-PatchCategory -Patch $p -SourceHint 'Installs') {
                'Pending'   { $totalPending++ }
                'Approved'  { $totalApproved++ }
                'Installed' { $totalInstalled++ }
            }
        }
    }

    $sb = New-Object System.Text.StringBuilder
    [void]$sb.AppendLine('<!DOCTYPE html>')
    [void]$sb.AppendLine('<html lang="en"><head><meta charset="utf-8">')
    [void]$sb.AppendLine('<meta name="viewport" content="width=device-width,initial-scale=1">')
    [void]$sb.AppendLine("<title>$(Format-HtmlValue $OrgDisplayName) - NinjaOne Patch Report</title>")
    [void]$sb.AppendLine('<style>')
    [void]$sb.AppendLine((Get-ReportCss))
    [void]$sb.AppendLine('</style></head><body>')

    [void]$sb.AppendLine('<a class="back" href="index.html">&laquo; Back to index</a>')
    [void]$sb.AppendLine("<h1>$(Format-HtmlValue $OrgDisplayName)</h1>")
    [void]$sb.AppendLine("<div class=""meta"">Generated: $(Format-HtmlValue $GeneratedAt) &middot; Region: ca.ninjarmm.com &middot; Devices: $($Devices.Count) &middot; Reporting window: $(Format-HtmlValue $Script:ReportWindow.Label) (previous calendar month)</div>")

    # Summary cards
    [void]$sb.AppendLine('<div class="summary">')
    [void]$sb.AppendLine("<div class=""card""><div class=""label"">Devices</div><div class=""value"">$($Devices.Count)</div></div>")
    [void]$sb.AppendLine("<div class=""card pending""><div class=""label"">Pending Patches</div><div class=""value"">$totalPending</div></div>")
    [void]$sb.AppendLine("<div class=""card approved""><div class=""label"">Approved Patches</div><div class=""value"">$totalApproved</div></div>")
    [void]$sb.AppendLine("<div class=""card installed""><div class=""label"">Installed Patches</div><div class=""value"">$totalInstalled</div></div>")
    [void]$sb.AppendLine('</div>')

    # Filter bar
    [void]$sb.AppendLine('<div class="filter-bar">')
    [void]$sb.AppendLine('<input type="text" id="hostFilter" placeholder="Filter devices by hostname...">')
    [void]$sb.AppendLine('</div>')

    [void]$sb.AppendLine('<h2>Devices</h2>')

    # Sort by hostname for readability
    $sortedDevices = $Devices | Sort-Object -Property @{Expression = {
        $h = ConvertTo-SafeValue $_.systemName
        if (-not $h) { $h = ConvertTo-SafeValue $_.dnsName }
        if (-not $h) { $h = ConvertTo-SafeValue $_.displayName }
        if (-not $h) { $h = [string]$_.id }
        [string]$h
    }}

    foreach ($d in $sortedDevices) {
        $deviceId = [string]$d.id
        $hostname = ConvertTo-SafeValue $d.systemName
        if (-not $hostname) { $hostname = ConvertTo-SafeValue $d.dnsName }
        if (-not $hostname) { $hostname = ConvertTo-SafeValue $d.displayName }

        $os = $null
        if ($d.os) {
            $os = ConvertTo-SafeValue $d.os.name
            if (-not $os) { $os = ConvertTo-SafeValue $d.os.manufacturer }
        }
        if (-not $os) { $os = ConvertTo-SafeValue $d.nodeClass }

        $entry = $DevicePatchMap[$deviceId]

        $devPending   = New-Object System.Collections.Generic.List[object]
        $devApproved  = New-Object System.Collections.Generic.List[object]
        $devInstalled = New-Object System.Collections.Generic.List[object]

        if ($entry) {
            $all = @()
            foreach ($p in @($entry.OsPatches))        { $all += [pscustomobject]@{ Patch = $p; Type = 'OS';       Source = 'Open' } }
            foreach ($p in @($entry.SoftwarePatches))  { $all += [pscustomobject]@{ Patch = $p; Type = 'Software'; Source = 'Open' } }
            foreach ($p in @($entry.OsInstalls))       { $all += [pscustomobject]@{ Patch = $p; Type = 'OS';       Source = 'Installs' } }
            foreach ($p in @($entry.SoftwareInstalls)) { $all += [pscustomobject]@{ Patch = $p; Type = 'Software'; Source = 'Installs' } }

            foreach ($row in $all) {
                switch (Get-PatchCategory -Patch $row.Patch -SourceHint $row.Source) {
                    'Pending'   { $devPending.Add($row)   | Out-Null }
                    'Approved'  { $devApproved.Add($row)  | Out-Null }
                    'Installed' { $devInstalled.Add($row) | Out-Null }
                    default     { }
                }
            }
        }

        $hostnameAttr = if ($hostname) { [System.Web.HttpUtility]::HtmlAttributeEncode([string]$hostname) } else { '' }

        [void]$sb.AppendLine("<details class=""device"" data-hostname=""$hostnameAttr"">")
        [void]$sb.AppendLine('<summary>')
        [void]$sb.AppendLine("<span>$(Format-HtmlValue $hostname) <span class=""device-meta"">&middot; $(Format-HtmlValue $os)</span></span>")
        [void]$sb.AppendLine('<span class="badges">')
        [void]$sb.AppendLine("<span class=""badge pending"">Pending: $($devPending.Count)</span>")
        [void]$sb.AppendLine("<span class=""badge approved"">Approved: $($devApproved.Count)</span>")
        [void]$sb.AppendLine("<span class=""badge installed"">Installed: $($devInstalled.Count)</span>")
        [void]$sb.AppendLine('</span>')
        [void]$sb.AppendLine('</summary>')
        [void]$sb.AppendLine('<div class="device-body">')

        foreach ($grp in @(
            @{ Title = 'Pending';   Badge = 'pending';   List = $devPending },
            @{ Title = 'Approved';  Badge = 'approved';  List = $devApproved },
            @{ Title = 'Installed'; Badge = 'installed'; List = $devInstalled }
        )) {
            [void]$sb.AppendLine("<div class=""section-title"">$($grp.Title) ($($grp.List.Count))</div>")
            if ($grp.List.Count -eq 0) {
                [void]$sb.AppendLine('<div class="empty">No patches in this category.</div>')
                continue
            }
            [void]$sb.AppendLine('<table class="patches"><thead><tr>')
            [void]$sb.AppendLine('<th>Name / KB</th><th>Severity</th><th>Status</th><th>Type</th><th>Last Status Change</th><th>Result / Error</th>')
            [void]$sb.AppendLine('</tr></thead><tbody>')
            foreach ($row in $grp.List) {
                $p = $row.Patch

                $name = ConvertTo-SafeValue $p.name
                if (-not $name) { $name = ConvertTo-SafeValue $p.title }
                $kb = ConvertTo-SafeValue $p.kbArticleId
                if (-not $kb) { $kb = ConvertTo-SafeValue $p.kb }
                $nameDisplay = if ($name -and $kb) { "$name (KB$kb)" }
                               elseif ($kb)       { "KB$kb" }
                               else               { $name }

                $severity = ConvertTo-SafeValue $p.severity
                $status   = ConvertTo-SafeValue $p.status
                # Belt-and-suspenders: if `status` happens to be blank on an
                # install record, fall back to installResult or a literal so
                # the badge isn't empty.
                if (-not $status -and $row.Source -eq 'Installs') {
                    $status = ConvertTo-SafeValue $p.installResult
                    if (-not $status) { $status = 'INSTALLED' }
                }

                $lastChange = $null
                foreach ($field in 'installedAt','timestamp','lastChange','installedOn','detectedOn','releaseDate') {
                    if ($p.PSObject.Properties[$field]) {
                        $lastChange = Format-EpochOrIso $p.$field
                        if ($lastChange) { break }
                    }
                }

                $result = ConvertTo-SafeValue $p.installResult
                if (-not $result) { $result = ConvertTo-SafeValue $p.errorMessage }
                if (-not $result) { $result = ConvertTo-SafeValue $p.result }

                [void]$sb.AppendLine('<tr>')
                [void]$sb.AppendLine("<td>$(Format-HtmlValue $nameDisplay)</td>")
                [void]$sb.AppendLine("<td>$(Format-HtmlValue $severity)</td>")
                [void]$sb.AppendLine("<td><span class=""badge $($grp.Badge)"">$(Format-HtmlValue $status)</span></td>")
                [void]$sb.AppendLine("<td>$(Format-HtmlValue $row.Type)</td>")
                [void]$sb.AppendLine("<td>$(Format-HtmlValue $lastChange)</td>")
                [void]$sb.AppendLine("<td>$(Format-HtmlValue $result)</td>")
                [void]$sb.AppendLine('</tr>')
            }
            [void]$sb.AppendLine('</tbody></table>')
        }

        [void]$sb.AppendLine('</div></details>')
    }

    if ($Devices.Count -eq 0) {
        [void]$sb.AppendLine('<div class="empty">No Windows devices in this organization.</div>')
    }

    [void]$sb.AppendLine('<script>')
    [void]$sb.AppendLine((Get-FilterJs))
    [void]$sb.AppendLine('</script>')
    [void]$sb.AppendLine('</body></html>')

    [System.IO.File]::WriteAllText($OutputPath, $sb.ToString(), [System.Text.Encoding]::UTF8)

    return [pscustomobject]@{
        Devices   = $Devices.Count
        Pending   = $totalPending
        Approved  = $totalApproved
        Installed = $totalInstalled
    }
}


# ---------------------------------------------------------------------------
# Index report
# ---------------------------------------------------------------------------

function New-IndexReport {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $true)]$OrgSummaries,   # array of [pscustomobject] rows
        [Parameter(Mandatory = $true)][string]$OutputPath,
        [Parameter(Mandatory = $true)][string]$GeneratedAt,
        [Parameter(Mandatory = $true)][int]$FilteredOutCount,
        [Parameter(Mandatory = $true)][int]$TotalDevicesSeen
    )

    $grandDevices   = 0
    $grandPending   = 0
    $grandApproved  = 0
    $grandInstalled = 0
    foreach ($row in $OrgSummaries) {
        $grandDevices   += [int]$row.Devices
        $grandPending   += [int]$row.Pending
        $grandApproved  += [int]$row.Approved
        $grandInstalled += [int]$row.Installed
    }

    $sb = New-Object System.Text.StringBuilder
    [void]$sb.AppendLine('<!DOCTYPE html>')
    [void]$sb.AppendLine('<html lang="en"><head><meta charset="utf-8">')
    [void]$sb.AppendLine('<meta name="viewport" content="width=device-width,initial-scale=1">')
    [void]$sb.AppendLine("<title>NinjaOne Patch Reports - $(Format-HtmlValue $GeneratedAt)</title>")
    [void]$sb.AppendLine('<style>')
    [void]$sb.AppendLine((Get-ReportCss))
    [void]$sb.AppendLine('</style></head><body>')

    [void]$sb.AppendLine("<h1>NinjaOne Patch Reports &mdash; $(Format-HtmlValue $GeneratedAt)</h1>")
    [void]$sb.AppendLine("<div class=""meta"">Region: ca.ninjarmm.com &middot; Reporting window: $(Format-HtmlValue $Script:ReportWindow.Label) (previous calendar month) &middot; Windows devices included: $grandDevices &middot; Non-Windows devices filtered out: $FilteredOutCount &middot; Total devices seen: $TotalDevicesSeen</div>")

    if ($OrgSummaries.Count -eq 0) {
        [void]$sb.AppendLine('<div class="empty">No Windows devices were found in any organization.</div>')
        [void]$sb.AppendLine('</body></html>')
        [System.IO.File]::WriteAllText($OutputPath, $sb.ToString(), [System.Text.Encoding]::UTF8)
        return
    }

    # Sort orgs alphabetically, case-insensitive
    $sorted = $OrgSummaries | Sort-Object -Property @{Expression = { ([string]$_.OrgName).ToLowerInvariant() }}

    [void]$sb.AppendLine('<table class="index"><thead><tr>')
    [void]$sb.AppendLine('<th>Organization</th><th class="num">Devices</th><th class="num">Pending</th><th class="num">Approved</th><th class="num">Installed</th>')
    [void]$sb.AppendLine('</tr></thead><tbody>')

    foreach ($row in $sorted) {
        $link = [System.Web.HttpUtility]::HtmlAttributeEncode([string]$row.FileName)
        [void]$sb.AppendLine('<tr>')
        [void]$sb.AppendLine("<td><a href=""$link"">$(Format-HtmlValue $row.OrgName)</a></td>")
        [void]$sb.AppendLine("<td class=""num"">$($row.Devices)</td>")
        [void]$sb.AppendLine("<td class=""num"">$($row.Pending)</td>")
        [void]$sb.AppendLine("<td class=""num"">$($row.Approved)</td>")
        [void]$sb.AppendLine("<td class=""num"">$($row.Installed)</td>")
        [void]$sb.AppendLine('</tr>')
    }

    [void]$sb.AppendLine('</tbody><tfoot><tr>')
    [void]$sb.AppendLine('<td>Grand Totals</td>')
    [void]$sb.AppendLine("<td class=""num"">$grandDevices</td>")
    [void]$sb.AppendLine("<td class=""num"">$grandPending</td>")
    [void]$sb.AppendLine("<td class=""num"">$grandApproved</td>")
    [void]$sb.AppendLine("<td class=""num"">$grandInstalled</td>")
    [void]$sb.AppendLine('</tr></tfoot></table>')

    [void]$sb.AppendLine('</body></html>')
    [System.IO.File]::WriteAllText($OutputPath, $sb.ToString(), [System.Text.Encoding]::UTF8)
}


# ---------------------------------------------------------------------------
# Main
# ---------------------------------------------------------------------------

function Main {
    [CmdletBinding()]
    param()

    Write-Host "Reporting window (previous calendar month): $($Script:ReportWindow.Label)" -ForegroundColor Cyan

    $Script:BearerToken = Get-NinjaToken

    $orgMap = Get-AllOrganizations
    $allDevices = Get-AllDevices

    # Apply Windows node-class filter
    $allowed = $Script:Config.AllowedNodeClass
    $winDevices = @($allDevices | Where-Object {
        $nc = ConvertTo-SafeValue $_.nodeClass
        if ($null -eq $nc) { return $false }
        $allowed -contains ([string]$nc).ToUpperInvariant()
    })
    $filteredOut = $allDevices.Count - $winDevices.Count
    Write-Host "Filtered out $filteredOut non-Windows device(s). Windows devices to process: $($winDevices.Count)" -ForegroundColor Green

    # Group by organization name (with unassigned bucket)
    $byOrg = @{}  # key: display name (or 'Unassigned'), value: { Name = ..., OrgId = ..., Devices = [list] }

    foreach ($d in $winDevices) {
        $orgId = $null
        if ($d.PSObject.Properties['organizationId']) { $orgId = $d.organizationId }

        $resolvedName = $null
        if ($null -ne $orgId -and $orgMap.ContainsKey([string]$orgId)) {
            $resolvedName = $orgMap[[string]$orgId]
        }
        if (-not (ConvertTo-SafeValue $resolvedName)) {
            $bucketKey = '___UNASSIGNED___'
            $displayName = 'Unassigned'
            $bucketOrgId = $null
        } else {
            $bucketKey = "$resolvedName|$orgId"
            $displayName = [string]$resolvedName
            $bucketOrgId = [string]$orgId
        }

        if (-not $byOrg.ContainsKey($bucketKey)) {
            $byOrg[$bucketKey] = [pscustomobject]@{
                DisplayName = $displayName
                OrgId       = $bucketOrgId
                Devices     = New-Object System.Collections.Generic.List[object]
            }
        }
        $byOrg[$bucketKey].Devices.Add($d) | Out-Null
    }

    # Collect patches for each device, one device at a time with progress
    $patchMap = @{}
    $total = $winDevices.Count
    $i = 0

    # Walk orgs so we can show the current org in progress
    foreach ($key in $byOrg.Keys) {
        $bucket = $byOrg[$key]
        foreach ($d in $bucket.Devices) {
            $i++
            $hostLabel = ConvertTo-SafeValue $d.systemName
            if (-not $hostLabel) { $hostLabel = ConvertTo-SafeValue $d.dnsName }
            if (-not $hostLabel) { $hostLabel = "id $($d.id)" }

            $pct = if ($total -gt 0) { [int](($i / $total) * 100) } else { 100 }
            Write-Progress -Activity "Collecting patches" `
                           -Status "[$i/$total] $($bucket.DisplayName) :: $hostLabel" `
                           -PercentComplete $pct

            try {
                $patchMap[[string]$d.id] = Get-DevicePatches -Device $d
            } catch {
                Write-Host "[Device $($d.id)] Unhandled error: $($_.Exception.Message)" -ForegroundColor Red
                $patchMap[[string]$d.id] = [pscustomobject]@{
                    OsPatches        = @()
                    SoftwarePatches  = @()
                    OsInstalls       = @()
                    SoftwareInstalls = @()
                }
            }
        }
    }
    Write-Progress -Activity "Collecting patches" -Completed

    # Prepare output folder
    $timestamp   = Get-Date -Format 'yyyyMMdd-HHmmss'
    $generatedAt = (Get-Date).ToString('yyyy-MM-dd HH:mm:ss zzz')
    $outputDir   = Join-Path -Path (Get-Location).Path -ChildPath "NinjaOne-PatchReports-$timestamp"
    New-Item -ItemType Directory -Path $outputDir -Force | Out-Null

    Write-Host "Writing reports to: $outputDir" -ForegroundColor Cyan

    # Build per-org files with filename collision handling
    $usedFileNames = @{}
    $orgSummaries = New-Object System.Collections.Generic.List[object]

    foreach ($key in $byOrg.Keys) {
        $bucket = $byOrg[$key]

        if ($key -eq '___UNASSIGNED___') {
            $fileName = 'Unassigned-PatchReport.html'
        } else {
            $safeBase = Get-SafeFileName -Name $bucket.DisplayName
            $fileName = "$safeBase-PatchReport.html"
            if ($usedFileNames.ContainsKey($fileName.ToLowerInvariant())) {
                $suffix = if ($bucket.OrgId) { $bucket.OrgId } else { [Guid]::NewGuid().ToString('N').Substring(0,8) }
                $fileName = "$safeBase-$suffix-PatchReport.html"
            }
        }
        $usedFileNames[$fileName.ToLowerInvariant()] = $true

        $outPath = Join-Path -Path $outputDir -ChildPath $fileName

        $summary = New-OrganizationReport `
            -OrgDisplayName $bucket.DisplayName `
            -Devices $bucket.Devices.ToArray() `
            -DevicePatchMap $patchMap `
            -OutputPath $outPath `
            -GeneratedAt $generatedAt

        $orgSummaries.Add([pscustomobject]@{
            OrgName   = $bucket.DisplayName
            FileName  = $fileName
            Devices   = $summary.Devices
            Pending   = $summary.Pending
            Approved  = $summary.Approved
            Installed = $summary.Installed
        }) | Out-Null

        Write-Host ("  {0,-40}  devices={1,-4} pending={2,-4} approved={3,-4} installed={4}" -f `
            $bucket.DisplayName, $summary.Devices, $summary.Pending, $summary.Approved, $summary.Installed) -ForegroundColor Gray
    }

    # Build index
    $indexPath = Join-Path -Path $outputDir -ChildPath 'index.html'
    New-IndexReport `
        -OrgSummaries $orgSummaries.ToArray() `
        -OutputPath $indexPath `
        -GeneratedAt $generatedAt `
        -FilteredOutCount $filteredOut `
        -TotalDevicesSeen $allDevices.Count

    Write-Host ""
    Write-Host "Index written to: $indexPath" -ForegroundColor Green

    # ---- Install diagnostics ----------------------------------------------
    Write-Host ""
    Write-Host "Install endpoint diagnostics:" -ForegroundColor Cyan
    Write-Host ("  OS patch installs:       API returned {0}, kept after date filter: {1}" -f `
        $Script:InstallDiagnostics.OsRaw, $Script:InstallDiagnostics.OsKept)
    Write-Host ("  Software patch installs: API returned {0}, kept after date filter: {1}" -f `
        $Script:InstallDiagnostics.SwRaw, $Script:InstallDiagnostics.SwKept)

    if ($Script:InstallDiagnostics.OsRaw -eq 0 -and $Script:InstallDiagnostics.SwRaw -eq 0) {
        Write-Host "  (The API returned zero install records across all devices for this window." -ForegroundColor Yellow
        Write-Host "   Either no patches were installed in $($Script:ReportWindow.Label), or the API credentials" -ForegroundColor Yellow
        Write-Host "   lack the 'monitoring' scope required to read install history.)" -ForegroundColor Yellow
    } elseif ($Script:InstallDiagnostics.OsRaw -gt 0 -and $Script:InstallDiagnostics.OsKept -eq 0) {
        Write-Host "  Sample OS install record (showing field names so you can see what date field to filter on):" -ForegroundColor Yellow
        $Script:InstallDiagnostics.SampleOsRecord | Format-List | Out-String | Write-Host
    }

    try {
        Start-Process -FilePath $indexPath | Out-Null
    } catch {
        Write-Host "(Could not auto-open the index: $($_.Exception.Message))" -ForegroundColor Yellow
    }
}

Main
