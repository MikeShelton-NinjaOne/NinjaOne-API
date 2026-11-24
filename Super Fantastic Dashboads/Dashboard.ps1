<#
.SYNOPSIS
    Generates a self-refreshing HTML dashboard of all NinjaOne servers.

.DESCRIPTION
    Every -RefreshSeconds (default 30) the script queries the NinjaOne API and rewrites an HTML file.
    The page itself also reloads every 30 seconds (meta refresh), so just leave it open in a browser.

    Page contents:
      - Table of all servers: health, active alerts summary, reboot pending, pending / failed patches
      - Donut 1: All servers vs Healthy servers
      - Donut 2: Online vs Offline servers (total in the centre)

    Health logic (edit Get-HealthStatus to taste):
      Offline  = device offline
      Critical = any CRITICAL or MAJOR active alert
      Warning  = any other active alert, any failed patch, or a reboot pending
      Healthy  = none of the above

.NOTES
    NinjaOne setup:
      Administration > Apps > API > Client app IDs > Add client app
        Application platform : API Services (machine-to-machine)
        Scopes               : Monitoring
        Allowed grant types  : Client Credentials
      Use the instance host that matches your tenant, e.g. app.ninjarmm.com, us2.ninjarmm.com,
      ca.ninjarmm.com, eu.ninjarmm.com, oc.ninjarmm.com.

.EXAMPLE
    .\NinjaOne-ServerDashboard.ps1 -Instance ca.ninjarmm.com -ClientId 'xxxx' -ClientSecret 'yyyy' -OpenBrowser

.EXAMPLE
    $env:NINJA_CLIENT_ID = 'xxxx'; $env:NINJA_CLIENT_SECRET = 'yyyy'
    .\NinjaOne-ServerDashboard.ps1 -Instance app.ninjarmm.com -OpenBrowser
#>
[CmdletBinding()]
param(
    [string]$Instance      = 'app.ninjarmm.com',
    [string]$ClientId      = $env:NINJA_CLIENT_ID,
    [string]$ClientSecret  = $env:NINJA_CLIENT_SECRET,
    [string]$OutputPath,
    [int]   $RefreshSeconds = 30,
    [string]$NodeClasses   = 'WINDOWS_SERVER,LINUX_SERVER,MAC_SERVER',
    [switch]$OpenBrowser,
    [switch]$Once
)

$ErrorActionPreference = 'Stop'
[Net.ServicePointManager]::SecurityProtocol = [Net.SecurityProtocolType]::Tls12

if (-not $ClientId -or -not $ClientSecret) {
    throw 'ClientId and ClientSecret are required (parameters or NINJA_CLIENT_ID / NINJA_CLIENT_SECRET env vars).'
}
if (-not $OutputPath) {
    $base = if ($PSScriptRoot) { $PSScriptRoot } else { (Get-Location).Path }
    $OutputPath = Join-Path $base 'NinjaOne-Server-Dashboard.html'
}

$script:Token       = $null
$script:TokenExpiry = [datetime]::MinValue
$DeviceFilter       = [uri]::EscapeDataString("class in ($NodeClasses)")
$Inv                = [Globalization.CultureInfo]::InvariantCulture

#region API helpers
function Get-NinjaToken {
    if ($script:Token -and (Get-Date) -lt $script:TokenExpiry) { return $script:Token }
    $body = @{
        grant_type    = 'client_credentials'
        client_id     = $ClientId
        client_secret = $ClientSecret
        scope         = 'monitoring'
    }
    try {
        $r = Invoke-RestMethod -Method Post -Uri "https://$Instance/ws/oauth/token" `
                -Body $body -ContentType 'application/x-www-form-urlencoded'
    } catch {
        $detail = $_.ErrorDetails.Message
        throw "TOKEN request to https://$Instance/ws/oauth/token failed: $($_.Exception.Message) $detail"
    }
    $script:Token       = $r.access_token
    $script:TokenExpiry = (Get-Date).AddSeconds([int]$r.expires_in - 60)
    return $script:Token
}

function Invoke-Ninja {
    param([Parameter(Mandatory)][string]$PathAndQuery, [string]$Method = 'Get', [string]$Body = $null)
    $headers = @{ Authorization = "Bearer $(Get-NinjaToken)"; Accept = 'application/json' }
    try {
        # Capture first, then return: makes Windows PowerShell 5.1 hand back a flat array
        # (it otherwise returns a JSON array as ONE nested object, which breaks the [int] casts later)
        if ($Method -eq 'Post') {
            $resp = Invoke-RestMethod -Method Post -Uri "https://$Instance$PathAndQuery" -Headers $headers -Body $Body -ContentType 'application/json'
        } else {
            $resp = Invoke-RestMethod -Method Get -Uri "https://$Instance$PathAndQuery" -Headers $headers
        }
        return $resp
    } catch {
        $detail = $_.ErrorDetails.Message
        throw "API call $PathAndQuery failed: $($_.Exception.Message) $detail"
    }
}

# /v2/devices pages with ?after=<last id>
function Get-AllServers {
    $all = @(); $after = 0; $size = 1000
    do {
        $page = @(Invoke-Ninja "/v2/devices?df=$DeviceFilter&pageSize=$size&after=$after")
        $page = @($page | Where-Object { $_ })
        $all += $page
        if ($page.Count) { $after = $page[-1].id }
    } while ($page.Count -ge $size)
    return $all
}

# /v2/queries/* endpoints page with a cursor and return { results: [], cursor: { name } }
function Get-CursorQuery {
    param([string]$Path, [string]$ExtraQuery = '', [switch]$AllDevices, [string]$Df = '')
    $out = @(); $cursor = $null
    do {
        $uri = "$Path`?df=$DeviceFilter&pageSize=1000$ExtraQuery"
        if ($AllDevices) { $uri = "$Path`?pageSize=1000$ExtraQuery" }
        if ($Df)         { $uri = "$Path`?df=$Df&pageSize=1000$ExtraQuery" }
        if ($cursor) { $uri += "&cursor=$cursor" }
        $r = Invoke-Ninja $uri
        $res = @()
        if ($r -and $r.results) { $res = @($r.results) }
        $out += $res
        $cursor = $null
        if ($res.Count -gt 0 -and $r.cursor -and $r.cursor.name) { $cursor = $r.cursor.name }
    } while ($cursor)
    return $out
}
$script:Warnings = New-Object System.Collections.Generic.List[string]
$script:Maintenance = @()
$script:Tickets = $null
$script:SwPatches = $null
$script:Nms = $null
$script:Cloud = $null
$script:BackupOv = $null
$script:BitLocker = $null

# Runs a section that may be unavailable; records a warning instead of failing the whole refresh.
function Get-SafeData {
    param([string]$Label, [scriptblock]$Script)
    try { return @(& $Script) }
    catch { $script:Warnings.Add("$Label unavailable: $($_.Exception.Message) $($_.ErrorDetails.Message)"); return @() }
}

# NinjaOne timestamps are epoch seconds (sometimes ms) - convert to local DateTime
function ConvertFrom-NinjaTime {
    param($Value)
    if ($null -eq $Value -or "$Value" -eq '') { return $null }
    try {
        $n = [double]::Parse("$Value", $Inv)
        if ($n -gt 100000000000) { return [DateTimeOffset]::FromUnixTimeMilliseconds([long]$n).LocalDateTime }
        return [DateTimeOffset]::FromUnixTimeSeconds([long][math]::Floor($n)).LocalDateTime
    } catch {
        try { return [datetime]$Value } catch { return $null }
    }
}

# Patch management activity feed (newest first). Capped at 10 pages of 1000 to keep refreshes quick.
function Get-PatchActivities {
    $out = @(); $older = $null
    for ($i = 0; $i -lt 10; $i++) {
        $uri = "/v2/activities?df=$DeviceFilter&type=PATCH_MANAGEMENT&pageSize=1000"
        if ($older) { $uri += "&olderThan=$older" }
        $r = Invoke-Ninja $uri
        $acts = @(); if ($r -and $r.activities) { $acts = @($r.activities) }
        if ($acts.Count -eq 0) { break }
        $out += $acts
        if (-not $r.lastActivityId) { break }
        $older = $r.lastActivityId
    }
    return $out
}
# Ninja Backup activity feed (newest first). Failed backup jobs are recorded here;
# /v2/backup/jobs only returns completed jobs. Capped at 10 pages of 1000.
function Get-BackupActivities {
    param([switch]$AllDevices)
    $out = @(); $older = $null
    for ($i = 0; $i -lt 10; $i++) {
        $uri = "/v2/activities?df=$DeviceFilter&type=NINJA_BACKUP&pageSize=1000"
        if ($AllDevices) { $uri = "/v2/activities?type=NINJA_BACKUP&pageSize=1000" }
        if ($older) { $uri += "&olderThan=$older" }
        $r = Invoke-Ninja $uri
        $acts = @(); if ($r -and $r.activities) { $acts = @($r.activities) }
        if ($acts.Count -eq 0) { break }
        $out += $acts
        if (-not $r.lastActivityId) { break }
        $older = $r.lastActivityId
    }
    return $out
}
# Ticket counts per status (parents + child statuses). A ticket list can be thousands of rows,
# so the result is cached and re-read every 5 minutes instead of on every 30-second refresh.
# Ticket status colors, matched to the dots NinjaOne shows beside each status.
# If the API ever returns a color on a status it is used as-is; otherwise match on the status name.
function Get-TicketColor {
    param([string]$Name, [string]$ApiColor = '', [string]$Fallback = '#7a8497')
    if ($ApiColor -match '^#?[0-9a-fA-F]{6}$') { return ('#' + $ApiColor.TrimStart('#')) }
    switch -Regex ($Name) {
        '(?i)^new'                      { return '#b5313a' }  # red
        '(?i)on[\s-]*hold'              { return '#b5313a' }  # red
        '(?i)^open'                     { return '#a8860f' }  # gold
        '(?i)sched'                     { return '#a8860f' }  # gold
        '(?i)^wait'                     { return '#16909c' }  # teal
        '(?i)paus'                      { return '#5b3fb0' }  # purple
        '(?i)approv'                    { return '#d98a14' }  # orange
        '(?i)reject'                    { return '#8a6414' }  # brown
        '(?i)resolv|complet|done'       { return '#1f7a4d' }  # green
        '(?i)clos|cancel'               { return '#5b6470' }  # grey
    }
    return $Fallback
}

$script:TicketCache     = $null
$script:TicketCacheTime = [datetime]::MinValue

function Get-TicketSummary {
    if ($script:TicketCache -and ((Get-Date) - $script:TicketCacheTime).TotalSeconds -lt 300) { return $script:TicketCache }

    # Status tree is read live from the tenant, so custom parent/child statuses are picked up automatically.
    $statuses = @(Invoke-Ninja '/v2/ticketing/statuses' | Where-Object { $_ -and $null -ne $_.statusId })
    $byId = @{}; foreach ($s in $statuses) { $byId["$($s.statusId)"] = $s }

    # Walk up the parent chain (any depth) to find the top-level status a status belongs to.
    function Get-RootId([string]$Id) {
        $cur = $Id
        for ($d = 0; $d -lt 10; $d++) {
            if (-not $byId.ContainsKey($cur)) { return $cur }
            $p = $byId[$cur].parentId
            if ($null -eq $p -or "$p" -eq '' -or "$p" -eq $cur -or -not $byId.ContainsKey("$p")) { return $cur }
            $cur = "$p"
        }
        return $cur
    }
    $closedRx = '(?i)^(resolved|closed|completed|done|cancel+ed)'

    $counts = @{}; $unassigned = 0; $total = 0
    $cursor = 0; $size = 500
    for ($i = 0; $i -lt 200; $i++) {
        $payload = @{ pageSize = $size; includeColumns = @('status', 'assignedAppUser') }
        if ($cursor -gt 0) { $payload['lastCursorId'] = $cursor }
        $r = Invoke-Ninja '/v2/ticketing/trigger/board/2/run' 'Post' ($payload | ConvertTo-Json -Compress)
        $rows = @(); if ($r -and $r.data) { $rows = @($r.data) }
        if ($rows.Count -eq 0) { break }
        foreach ($t in $rows) {
            $total++
            $sid = "$($t.status.statusId)"
            if ($counts.ContainsKey($sid)) { $counts[$sid]++ } else { $counts[$sid] = 1 }
            if (-not $t.assignedAppUser) {
                $rootName = ''
                $rid = Get-RootId $sid
                if ($byId.ContainsKey($rid)) { $rootName = "$($byId[$rid].displayName)" } elseif ($t.status.displayName) { $rootName = "$($t.status.displayName)" }
                if ($rootName -notmatch $closedRx) { $unassigned++ }
            }
            # remember statuses seen on tickets that the status list did not return
            if (-not $byId.ContainsKey($sid) -and $t.status.displayName) {
                $byId[$sid] = [pscustomobject]@{ statusId = $t.status.statusId; displayName = $t.status.displayName; parentId = $t.status.parentId; sortId = $t.status.sortId }
                $statuses += $byId[$sid]
            }
        }
        $next = 0; if ($r.metadata -and $r.metadata.lastCursorId) { $next = [int]$r.metadata.lastCursorId }
        if ($next -le $cursor) { $next = $cursor + $rows.Count }
        $cursor = $next
    }

    # Build groups: every top-level status, its direct/indirect children, counts rolled up.
    $roots = @($statuses | Where-Object { (Get-RootId "$($_.statusId)") -eq "$($_.statusId)" } | Sort-Object { [int]$_.sortId }, { [int]$_.statusId })
    $groups = @(foreach ($p in $roots) {
        $pid2 = "$($p.statusId)"
        $own = 0; if ($counts.ContainsKey($pid2)) { $own = $counts[$pid2] }
        $kidObjs = @(foreach ($c in ($statuses | Where-Object { "$($_.statusId)" -ne $pid2 -and (Get-RootId "$($_.statusId)") -eq $pid2 } | Sort-Object { [int]$_.sortId }, { [int]$_.statusId })) {
            $n = 0; if ($counts.ContainsKey("$($c.statusId)")) { $n = $counts["$($c.statusId)"] }
            [pscustomobject]@{ Name = "$($c.displayName)"; Count = $n; ApiColor = "$($c.color)$($c.colour)" }
        })
        $sum = $own; foreach ($ko in $kidObjs) { $sum += $ko.Count }
        [pscustomobject]@{ Id = $pid2; Name = "$($p.displayName)"; Total = $sum; Own = $own; Children = $kidObjs; ApiColor = "$($p.color)$($p.colour)"; Closed = ("$($p.displayName)" -match $closedRx) }
    })

    $script:TicketCache = [pscustomobject]@{ Groups = $groups; Unassigned = $unassigned; Total = $total; Time = (Get-Date) }
    $script:TicketCacheTime = Get-Date
    return $script:TicketCache
}

# Cloud monitors are devices of class CLOUD_MONITOR_TARGET; their alerts come from the normal alerts feed.
function Get-CloudMonitors {
    $df = [uri]::EscapeDataString('class in (CLOUD_MONITOR_TARGET)')
    $all = @(); $after = 0; $size = 1000
    do {
        $page = @(Invoke-Ninja "/v2/devices?df=$df&pageSize=$size&after=$after")
        $page = @($page | Where-Object { $_ })
        $all += $page
        if ($page.Count) { $after = $page[-1].id }
    } while ($page.Count -ge $size)
    $orgs = @{}
    try { foreach ($o in @(Invoke-Ninja '/v2/organizations')) { if ($o) { $orgs["$($o.id)"] = "$($o.name)" } } } catch { }
    $byId = @{}
    $devs = @(foreach ($d in $all) {
        $nm = $d.displayName; if (-not $nm) { $nm = $d.systemName }; if (-not $nm) { $nm = $d.dnsName }; if (-not $nm) { $nm = "Monitor $($d.id)" }
        $o = [pscustomobject]@{ Id = $d.id; Name = $nm; Org = $orgs["$($d.organizationId)"]; Online = -not [bool]$d.offline }
        $byId["$($d.id)"] = $o
        $o
    })
    $al = @(@(Invoke-Ninja '/v2/alerts') | Where-Object { $_ -and $byId.ContainsKey("$($_.deviceId)") } | ForEach-Object {
        $msg = "$($_.message)"; if (-not $msg) { $msg = "$($_.subject)" }
        $t = ConvertFrom-NinjaTime $(if ($_.createTime) { $_.createTime } else { $_.createdTime })
        [pscustomobject]@{ Name = $byId["$($_.deviceId)"].Name; Org = $byId["$($_.deviceId)"].Org; Group = (Get-AlertGroup $_); Health = (Get-AlertLabel $_); When = $t; Message = $msg }
    })
    [pscustomobject]@{ Devices = $devs; Alerts = $al }
}

# Network-monitoring (NMS) devices: every device whose class starts with NMS_ (switches, routers, firewalls, printers, ...).
# Classes are not hard-coded, so any NMS type a tenant has is picked up automatically.
function Get-NmsDevices {
    $all = @(); $after = 0; $size = 1000
    do {
        $page = @(Invoke-Ninja "/v2/devices?pageSize=$size&after=$after")
        $page = @($page | Where-Object { $_ })
        $all += $page
        if ($page.Count) { $after = $page[-1].id }
    } while ($page.Count -ge $size)
    $orgs = @{}
    try { foreach ($o in @(Invoke-Ninja '/v2/organizations')) { if ($o) { $orgs["$($o.id)"] = "$($o.name)" } } } catch { }
    foreach ($d in ($all | Where-Object { "$($_.nodeClass)" -like 'NMS_*' })) {
        $nm = $d.displayName; if (-not $nm) { $nm = $d.systemName }; if (-not $nm) { $nm = $d.dnsName }; if (-not $nm) { $nm = "Device $($d.id)" }
        $seen = ConvertFrom-NinjaTime $d.lastContact
        if (-not $seen) { $seen = ConvertFrom-NinjaTime $d.lastUpdate }
        [pscustomobject]@{
            Id       = $d.id
            Name     = $nm
            Class    = ("$($d.nodeClass)" -replace '^NMS_', '' -replace '_', ' ')
            Org      = $orgs["$($d.organizationId)"]
            Online   = -not [bool]$d.offline
            LastSeen = $seen
        }
    }
}

# All devices (any class, not just servers) currently in or scheduled for maintenance mode
function Get-MaintenanceDevices {
    $all = @(); $after = 0; $size = 1000
    do {
        $page = @(Invoke-Ninja "/v2/devices?pageSize=$size&after=$after")
        $page = @($page | Where-Object { $_ })
        $all += $page
        if ($page.Count) { $after = $page[-1].id }
    } while ($page.Count -ge $size)

    $orgs = @{}
    try { foreach ($o in @(Invoke-Ninja '/v2/organizations')) { if ($o) { $orgs["$($o.id)"] = "$($o.name)" } } } catch { }

    # The device record has no list of suppressed features; NinjaOne logs them in the
    # "Maintenance Mode is Active, suppressing ..." activity. Newest first, so the first hit per device wins.
    $featBy = @{}
    $need   = @($all | Where-Object { $_.maintenance } | ForEach-Object { "$($_.id)" })
    # Per-device history (the global activity filter does not accept maintenance entries). Only devices
    # currently in maintenance are checked; up to 5 pages of 100 each, stopping at the first hit.
    foreach ($id in $need) {
        try {
            $older = $null
            for ($i = 0; $i -lt 5; $i++) {
                $uri = "/v2/device/$id/activities?pageSize=100"
                if ($older) { $uri += "&olderThan=$older" }
                $r = Invoke-Ninja $uri
                $acts = @(); if ($r -and $r.activities) { $acts = @($r.activities) }
                if ($acts.Count -eq 0) { break }
                foreach ($a in $acts) {
                    if ("$($a.statusCode)" -ne 'MAINTENANCE_MODE_ACTIVE') { continue }
                    $opt = "$($a.data.message.params.options)"
                    if (-not $opt -and "$($a.message)" -match 'suppressing ([A-Z_,]+)') { $opt = $Matches[1] }
                    if ($opt) { $featBy[$id] = @($opt.Split(',') | Where-Object { $_ }); break }
                }
                if ($featBy.ContainsKey($id)) { break }
                if (-not $r.lastActivityId) { break }
                $older = $r.lastActivityId
            }
        } catch { }
    }

    foreach ($d in $all) {
        $m = $d.maintenance
        if (-not $m) { continue }
        $feat = @(); if ($m.disabledFeatures) { $feat = @($m.disabledFeatures) } elseif ($featBy.ContainsKey("$($d.id)")) { $feat = @($featBy["$($d.id)"]) }
        $nm = $d.displayName; if (-not $nm) { $nm = $d.systemName }; if (-not $nm) { $nm = $d.dnsName }; if (-not $nm) { $nm = "Device $($d.id)" }
        [pscustomobject]@{
            Name     = $nm
            Class    = "$($d.nodeClass)"
            Org      = $orgs["$($d.organizationId)"]
            Status   = "$($m.status)"
            Start    = (ConvertFrom-NinjaTime $m.start)
            End      = (ConvertFrom-NinjaTime $m.end)
            Features = $feat
        }
    }
}

# Antivirus / EDR threats reported to NinjaOne (CrowdStrike, SentinelOne, Defender, etc.)
function Get-ThreatQuery {
    $out = @(); $cursor = $null
    do {
        $uri = "/v2/queries/antivirus-threats?df=$DeviceFilter&pageSize=1000"
        if ($cursor) { $uri += "&cursor=$cursor" }
        $r = Invoke-Ninja $uri
        $res = @()
        if ($r -and $r.results) { $res = @($r.results) }
        elseif ($r -is [array]) { $res = @($r | Where-Object { $_ }) }
        $out += $res
        $cursor = $null
        if ($r -and -not ($r -is [array]) -and $res.Count -gt 0 -and $r.cursor -and $r.cursor.name) { $cursor = $r.cursor.name }
    } while ($cursor)
    return $out
}

# First non-empty property from a list of possible names (field names differ between AV vendors)
function Get-Prop {
    param($Obj, [string[]]$Names)
    foreach ($n in $Names) { $v = $Obj.$n; if ($null -ne $v -and "$v" -ne '') { return "$v" } }
    return ''
}

function ConvertTo-Threat {
    param($T)
    $raw = Get-Prop $T @('status','threatStatus','state','mitigationStatus')
    $u = $raw.ToUpper()
    $st = 'Other'
    if     ($u -match 'QUARANTIN')                          { $st = 'Quarantined' }
    elseif ($u -match 'BLOCK')                              { $st = 'Blocked' }
    elseif ($u -match 'ACTIVE' -and $u -notmatch 'INACTIVE') { $st = 'Active' }
    $when = $null
    foreach ($n in 'detectedTime','detectedAt','timestamp','createTime','createdTime','firstSeen','time') {
        if ($T.$n) { $when = ConvertFrom-NinjaTime $T.$n; if ($when) { break } }
    }
    [pscustomobject]@{
        DeviceId = $T.deviceId
        Status   = $st
        Raw      = $raw
        Name     = (Get-Prop $T @('threatName','name','threat','fileName','path'))
        Vendor   = (Get-Prop $T @('vendor','productName','product','source','antivirusProduct'))
        Category = (Get-Prop $T @('category','classification','type'))
        When     = $when
    }
}
#endregion

#region Model
$SevRank = @{ NONE = 0; MINOR = 1; MODERATE = 2; MAJOR = 3; CRITICAL = 4 }

function Get-HealthStatus {
    param($Online, $MaxSev, $Failed, $Reboot)
    if (-not $Online)                      { return 'Offline'  }
    if ($MaxSev -ge 3)                     { return 'Critical' }
    if ($MaxSev -ge 1 -or $Failed -gt 0 -or $Reboot) { return 'Warning' }
    return 'Healthy'
}

# Maps NinjaOne's own health status (HEALTHY / NEEDS_ATTENTION / UNHEALTHY / UNKNOWN) to a label.
# If the health query returned nothing at all, falls back to the dashboard's derived health.
function ConvertTo-NinjaHealth {
    param([string]$Raw, [string]$Derived, [bool]$HaveData)
    if ($HaveData) {
        $r = $Raw.ToUpper()
        if ($r -match 'UNHEALTHY')       { return 'Unhealthy' }
        if ($r -match 'NEEDS|ATTENTION') { return 'Needs attention' }
        if ($r -match 'HEALTHY')         { return 'Healthy' }
        return 'Unknown'
    }
    switch ($Derived) {
        'Healthy'  { return 'Healthy' }
        'Warning'  { return 'Needs attention' }
        'Critical' { return 'Unhealthy' }
        default    { return 'Unknown' }
    }
}

# Backup storage usage per organization (same figures as the Backup > Device overview table in NinjaOne).
# Cached for 5 minutes: it needs one API call per organization.
$script:UsageCache = $null
$script:UsageCacheTime = [datetime]::MinValue
function Get-BackupUsage {
    if ($script:UsageCache -and ((Get-Date) - $script:UsageCacheTime).TotalSeconds -lt 300) { return $script:UsageCache }
    $local = 0.0; $cloud = 0.0
    foreach ($o in @(Invoke-Ninja '/v2/organizations')) {
        if (-not $o) { continue }
        $rows = @(Invoke-Ninja "/v2/organization/$($o.id)/locations/backup/usage")
        foreach ($r in $rows) {
            if (-not $r) { continue }
            $local += [double]$r.localTotalSize
            $cloud += [double]$r.cloudTotalSize
        }
    }
    $script:UsageCache = [pscustomobject]@{ Local = $local; Cloud = $cloud }
    $script:UsageCacheTime = Get-Date
    return $script:UsageCache
}

# BitLocker coverage of Windows desktops and laptops (WINDOWS_WORKSTATION). A device counts as protected when its
# system volume (C: if present, otherwise its first volume) reports BitLocker protection ON / fully encrypted.
# Cached for 5 minutes because it queries every workstation volume.
$script:BlCache = $null
$script:BlCacheTime = [datetime]::MinValue
function Get-BitLockerSummary {
    if ($script:BlCache -and ((Get-Date) - $script:BlCacheTime).TotalSeconds -lt 300) { return $script:BlCache }
    $df = [uri]::EscapeDataString('class in (WINDOWS_WORKSTATION)')
    $devs = @(); $after = 0; $size = 1000
    do {
        $page = @(Invoke-Ninja "/v2/devices?df=$df&pageSize=$size&after=$after")
        $page = @($page | Where-Object { $_ })
        $devs += $page
        if ($page.Count) { $after = $page[-1].id }
    } while ($page.Count -ge $size)
    $vols = @(Get-CursorQuery '/v2/queries/volumes' '&include=bl' -Df $df)
    $byDev = @{}
    foreach ($v in $vols) { if ($v) { $k = "$($v.deviceId)"; if (-not $byDev.ContainsKey($k)) { $byDev[$k] = New-Object System.Collections.ArrayList }; [void]$byDev[$k].Add($v) } }
    $on = 0; $off = 0; $none = 0
    foreach ($d in $devs) {
        $k = "$($d.id)"
        $bl = $null
        if ($byDev.ContainsKey($k)) {
            $sys = @($byDev[$k] | Where-Object { "$($_.driveLetter)$($_.name)" -match '^C' } | Select-Object -First 1)
            if ($sys.Count -eq 0) { $sys = @($byDev[$k] | Select-Object -First 1) }
            if ($sys.Count -gt 0) { $bl = $sys[0].bitLockerStatus }
        }
        if (-not $bl) { $none++; continue }
        if ("$($bl.protectionStatus)".ToUpper() -eq 'PROTECTED' -or "$($bl.conversionStatus)".ToUpper() -match 'FULLY_ENCRYPTED') { $on++ } else { $off++ }
    }
    $script:BlCache = [pscustomobject]@{ Enabled = $on; Disabled = $off; NoData = $none; Total = $devs.Count }
    $script:BlCacheTime = Get-Date
    return $script:BlCache
}

function Get-BackupBucket([string]$Status) {
    $u = $Status.ToUpper()
    if ($u -match 'FAIL')                                  { return 'Failed' }
    if ($u -match 'WARN|PARTIAL')                          { return 'Warnings' }
    if ($u -match 'PROCESS|RUN|PROGRESS|START|QUEUE')      { return 'Running' }
    if ($u -match 'COMPLET|SUCCESS')                       { return 'Success' }
    return 'Other'
}

function Format-Bytes([double]$b) {
    if ($b -ge 1TB) { return ('{0:N2} TB' -f ($b / 1TB)) }
    if ($b -ge 1GB) { return ('{0:N2} GB' -f ($b / 1GB)) }
    if ($b -ge 1MB) { return ('{0:N1} MB' -f ($b / 1MB)) }
    return ('{0:N0} KB' -f ($b / 1KB))
}

function Get-DashboardData {
    $servers = Get-AllServers
    $ids     = @{}; foreach ($s in $servers) { $ids[[int]$s.id] = $true }

    $alerts  = @(Invoke-Ninja '/v2/alerts') | Where-Object { $_ -and $ids.ContainsKey([int]$_.deviceId) }
    $os      = Get-CursorQuery '/v2/queries/operating-systems'
    # One unfiltered query, classified client-side: NinjaOne uses several "not installed" statuses
    # (PENDING, APPROVED, MANUAL, ...), so filtering on status=PENDING alone misses most patches.
    $allPatches = Get-CursorQuery '/v2/queries/os-patches'
    $pending = @($allPatches | Where-Object { "$($_.status)".ToUpper() -notmatch 'FAIL|REJECT|INSTALLED' })
    $failed  = @($allPatches | Where-Object { "$($_.status)".ToUpper() -match 'FAIL' })
    $installs = Get-SafeData 'Patch install history' { Get-CursorQuery '/v2/queries/os-patch-installs' "&installedAfter=$((Get-Date).AddDays(-90).ToString('yyyyMMdd'))" }
    $acts     = Get-SafeData 'Patch scan activity'   { Get-PatchActivities }
    $volumes  = Get-SafeData 'Disk volumes'          { Get-CursorQuery '/v2/queries/volumes' }
    $bjobs    = Get-SafeData 'Backup jobs'           { Get-CursorQuery '/v2/backup/jobs' }
    $hlth     = Get-SafeData 'Device health status'  { Get-CursorQuery '/v2/queries/device-health' }
    $bacts    = Get-SafeData 'Backup activity'       { Get-BackupActivities }
    $threatsRaw = Get-SafeData 'Antivirus threats'   { Get-ThreatQuery }
    $script:Maintenance = Get-SafeData 'Maintenance mode' { Get-MaintenanceDevices }
    $tkRes = @(Get-SafeData 'Ticketing' { Get-TicketSummary })
    if ($tkRes.Count -gt 0 -and $tkRes[0]) { $script:Tickets = $tkRes[0] }

    $nmsRes = @(Get-SafeData 'NMS devices' { @(Get-NmsDevices) })
    $script:Nms = $nmsRes
    $blRes = @(Get-SafeData 'BitLocker' { Get-BitLockerSummary })
    $script:BitLocker = $null; if ($blRes.Count -gt 0 -and $blRes[0]) { $script:BitLocker = $blRes[0] }
    $cmRes = @(Get-SafeData 'Cloud monitors' { Get-CloudMonitors })
    $script:Cloud = $null; if ($cmRes.Count -gt 0 -and $cmRes[0]) { $script:Cloud = $cmRes[0] }

    # Software (third-party) patches: same classification as the OS patch chart
    $swRes = @(Get-SafeData 'Software patches' {
        $sp = @(Get-CursorQuery '/v2/queries/software-patches' | Where-Object { $_ -and $ids.ContainsKey([int]$_.deviceId) })
        $si = @(Get-CursorQuery '/v2/queries/software-patch-installs' "&installedAfter=$((Get-Date).AddDays(-90).ToString('yyyyMMdd'))" | Where-Object { $_ -and $ids.ContainsKey([int]$_.deviceId) })
        $appr = @($sp | Where-Object { "$($_.status)".ToUpper() -eq 'APPROVED' }).Count
        $fail = @($sp | Where-Object { "$($_.status)".ToUpper() -match 'FAIL' }).Count
        $pendAll = @($sp | Where-Object { "$($_.status)".ToUpper() -notmatch 'FAIL|REJECT|INSTALLED' }).Count
        [pscustomobject]@{ Installed = $si.Count; Approved = $appr; Failed = $fail; Pending = [Math]::Max(0, $pendAll - $appr) }
    })
    $script:SwPatches = $null
    if ($swRes.Count -gt 0 -and $swRes[0]) { $script:SwPatches = $swRes[0] }

    $alertsBy  = $alerts  | Group-Object deviceId -AsHashTable -AsString
    $pendingBy = $pending | Group-Object deviceId -AsHashTable -AsString
    $failedBy  = $failed  | Group-Object deviceId -AsHashTable -AsString
    $osBy      = @{}; foreach ($o in $os) { $osBy["$($o.deviceId)"] = $o }
    if (-not $alertsBy)  { $alertsBy  = @{} }
    if (-not $pendingBy) { $pendingBy = @{} }
    if (-not $failedBy)  { $failedBy  = @{} }
    # Devices that have any patch data at all (used to infer "patch-enabled")
    $patchDev = @{}
    foreach ($p in $allPatches) { $patchDev["$($p.deviceId)"] = $true }
    # Per-device approved-patch and installed-patch counts (for the OS Patches Installed chart)
    $approvedBy  = @($allPatches | Where-Object { "$($_.status)".ToUpper() -eq 'APPROVED' }) | Group-Object deviceId -AsHashTable -AsString
    $installedBy = @($installs) | Group-Object deviceId -AsHashTable -AsString
    if (-not $approvedBy)  { $approvedBy  = @{} }
    if (-not $installedBy) { $installedBy = @{} }

    # Threats per device: only Active / Blocked / Quarantined are kept
    $threatsBy = @{}
    foreach ($th in $threatsRaw) {
        $n = ConvertTo-Threat $th
        if ($n.Status -eq 'Other') { continue }
        $k = "$($n.DeviceId)"
        if (-not $threatsBy.ContainsKey($k)) { $threatsBy[$k] = New-Object System.Collections.ArrayList }
        [void]$threatsBy[$k].Add($n)
    }

    # NinjaOne's own health status per device
    $healthBy = @{}
    foreach ($h in $hlth) {
        $st = "$($h.healthStatus)"; if (-not $st) { $st = "$($h.health)" }
        $healthBy["$($h.deviceId)"] = $st
    }

    # Backup: newest job per device (a device with no jobs is treated as "not enabled")
    $backupBy = @{}
    foreach ($b in $bjobs) {
        $t = ConvertFrom-NinjaTime $b.jobStartTime
        if (-not $t) { continue }
        $k = "$($b.deviceId)"
        if (-not $backupBy.ContainsKey($k) -or $t -gt $backupBy[$k].Time) {
            $backupBy[$k] = [pscustomobject]@{ Time = $t; Status = "$($b.jobStatus)"; Plan = "$($b.planName)" }
        }
    }

    # Backup activity: picks up FAILED / warning / cancelled jobs that the jobs list does not return.
    # Whichever record (job or activity) is newest for a device wins.
    foreach ($b in $bacts) {
        $code = "$($b.statusCode)"
        if ($code -notmatch 'BACKUPJOB_(COMPLETED|FAILED|CANCELLED|CANCELED)') { continue }
        $t = ConvertFrom-NinjaTime $b.activityTime
        if (-not $t) { continue }
        $st = 'COMPLETED'
        if     ($code -match 'FAIL')     { $st = 'FAILED' }
        elseif ($code -match 'WARNING')  { $st = 'WARNING' }
        elseif ($code -match 'CANCEL')   { $st = 'CANCELLED' }
        $plan = "$($b.data.message.params.plan_name)"
        $k = "$($b.deviceId)"
        if (-not $backupBy.ContainsKey($k) -or $t -gt $backupBy[$k].Time) {
            $backupBy[$k] = [pscustomobject]@{ Time = $t; Status = $st; Plan = $plan }
        }
    }

    # Backup overview: ALL devices (not just servers). Last result per device + backup plan,
    # combining the jobs list with backup activity (which carries failed jobs the list omits).
    $ovJobs = @(Get-SafeData 'Backup overview jobs' { Get-CursorQuery '/v2/backup/jobs' -AllDevices })
    $ovActs = @(Get-SafeData 'Backup overview activity' { Get-BackupActivities -AllDevices })
    $lastRes = @{}; $lastPlan = @{}; $devSeen = @{}
    foreach ($j in $ovJobs) {
        $t = ConvertFrom-NinjaTime $j.jobStartTime; if (-not $t) { continue }
        $devSeen["$($j.deviceId)"] = $true
        $k = "$($j.deviceId)|$($j.planName)"
        if (-not $lastRes.ContainsKey($k) -or $t -gt $lastRes[$k].Time) { $lastRes[$k] = [pscustomobject]@{ Time = $t; Status = "$($j.jobStatus)" } }
        $k2 = "$($j.deviceId)|$($j.planGuid)"
        if (-not $lastPlan.ContainsKey($k2) -or $t -gt $lastPlan[$k2].Time) { $lastPlan[$k2] = [pscustomobject]@{ Time = $t; Dest = "$($j.destination)".ToUpper(); Bytes = [double]$j.totalStoredBytes } }
    }
    foreach ($b in $ovActs) {
        $code = "$($b.statusCode)"
        if ($code -notmatch 'BACKUPJOB_(COMPLETED|FAILED|CANCELLED|CANCELED|WARNING)') { continue }
        $t = ConvertFrom-NinjaTime $b.activityTime; if (-not $t) { continue }
        $st = 'COMPLETED'
        if     ($code -match 'FAIL')    { $st = 'FAILED' }
        elseif ($code -match 'WARNING') { $st = 'WARNING' }
        elseif ($code -match 'CANCEL')  { $st = 'CANCELLED' }
        $devSeen["$($b.deviceId)"] = $true
        $k = "$($b.deviceId)|$($b.data.message.params.plan_name)"
        if (-not $lastRes.ContainsKey($k) -or $t -gt $lastRes[$k].Time) { $lastRes[$k] = [pscustomobject]@{ Time = $t; Status = $st } }
    }
    $ovCount = @{ Success = 0; Warnings = 0; Failed = 0; Running = 0 }
    foreach ($v in $lastRes.Values) { $bk = Get-BackupBucket $v.Status; if ($ovCount.ContainsKey($bk)) { $ovCount[$bk]++ } }

    # Integrity checks: the endpoint path differs between NinjaOne versions, so try the known spellings.
    $intRes = @(); $intErr = $null
    foreach ($path in '/v2/backup/integrity-check-jobs', '/v2/backup/integrity-check/jobs', '/v2/backup/integrity-checks') {
        try { $intRes = @(Get-CursorQuery $path -AllDevices); $intErr = $null; break } catch { $intErr = $_ }
    }
    if ($intErr) { $script:Warnings.Add("Integrity checks unavailable: $($intErr.Exception.Message)") }
    $intLast = @{}
    foreach ($j in $intRes) {
        $t = ConvertFrom-NinjaTime $j.jobStartTime; if (-not $t) { continue }
        $k = "$($j.deviceId)|$($j.planGuid)"
        if (-not $intLast.ContainsKey($k) -or $t -gt $intLast[$k].Time) { $intLast[$k] = [pscustomobject]@{ Time = $t; Status = "$($j.jobStatus)" } }
    }
    $intCount = @{ Success = 0; Warnings = 0; Failed = 0; Running = 0 }
    foreach ($v in $intLast.Values) { $bk = Get-BackupBucket $v.Status; if ($intCount.ContainsKey($bk)) { $intCount[$bk]++ } }
    $cloudB = 0.0; $localB = 0.0; $usageOk = $false
    $usRes = @(Get-SafeData 'Backup usage' { Get-BackupUsage })
    if ($usRes.Count -gt 0 -and $usRes[0]) { $cloudB = $usRes[0].Cloud; $localB = $usRes[0].Local; $usageOk = $true }
    $script:BackupOv = [pscustomobject]@{
        Jobs = $ovCount; Integrity = $intCount; IntegrityN = $intLast.Count; IntegrityOk = (-not $intErr)
        Devices = $devSeen.Count; Local = $localB; Cloud = $cloudB; UsageOk = $usageOk
    }

    # Disk: all volumes per device
    $diskBy = @{}
    foreach ($v in $volumes) {
        $cap = [double]$v.capacity
        if ($cap -le 0) { continue }
        $pct = [math]::Round(100 * ($cap - [double]$v.freeSpace) / $cap, 0)
        $k = "$($v.deviceId)"
        if (-not $diskBy.ContainsKey($k)) { $diskBy[$k] = New-Object System.Collections.ArrayList }
        $nm = "$($v.name)"; if (-not $nm) { $nm = "$($v.label)" }
        [void]$diskBy[$k].Add([pscustomobject]@{ Pct = $pct; Name = $nm })
    }

    # Last patch apply per device = newest record in install history (last 90 days)
    $lastApply = @{}
    foreach ($i in $installs) {
        $t = ConvertFrom-NinjaTime $i.installedAt
        if (-not $t) { continue }
        $k = "$($i.deviceId)"
        if (-not $lastApply.ContainsKey($k) -or $t -gt $lastApply[$k].Time) {
            $lastApply[$k] = [pscustomobject]@{ Time = $t; Failed = ("$($i.status)" -match 'FAIL') }
        }
    }
    # Last patch scan per device (and apply fallback) from the activity feed, newest first
    $lastScan = @{}
    foreach ($a in ($acts | Sort-Object { [double]$_.activityTime } -Descending)) {
        $code = "$($a.statusCode)"
        if ($code -notmatch 'COMPLETED|FAIL|SUCCESS') { continue }
        $t = ConvertFrom-NinjaTime $a.activityTime
        if (-not $t) { continue }
        $k      = "$($a.deviceId)"
        $isFail = ($code -match 'FAIL')
        if ($code -match 'SCAN') {
            if (-not $lastScan.ContainsKey($k)) { $lastScan[$k] = [pscustomobject]@{ Time = $t; Failed = $isFail } }
        } elseif ($code -match 'APPLY|INSTALL') {
            if (-not $lastApply.ContainsKey($k)) { $lastApply[$k] = [pscustomobject]@{ Time = $t; Failed = $isFail } }
        }
    }

    $rows = foreach ($s in $servers) {
        $key    = "$($s.id)"
        $online = -not [bool]$s.offline
        $devAl  = @(); if ($alertsBy.ContainsKey($key)) { $devAl = @($alertsBy[$key]) }
        $maxSev = 0
        foreach ($a in $devAl) {
            $r = $SevRank["$($a.severity)".ToUpper()]
            if ($r -gt $maxSev) { $maxSev = $r }
        }
        $pCount = 0; if ($pendingBy.ContainsKey($key)) { $pCount = @($pendingBy[$key]).Count }
        $fCount = 0; if ($failedBy.ContainsKey($key))  { $fCount = @($failedBy[$key]).Count }
        $reboot = $false; if ($osBy.ContainsKey($key) -and $osBy[$key].needsReboot) { $reboot = $true }

        $name = $s.displayName; if (-not $name) { $name = $s.systemName }; if (-not $name) { $name = $s.dnsName }; if (-not $name) { $name = "Device $($s.id)" }
        [pscustomobject]@{
            Id       = $s.id
            Name     = $name
            Class    = $s.nodeClass
            Online   = $online
            Alerts   = $devAl
            MaxSev   = $maxSev
            Reboot   = $reboot
            Pending  = $pCount
            Failed   = $fCount
            Health   = Get-HealthStatus -Online $online -MaxSev $maxSev -Failed $fCount -Reboot $reboot
            NinjaHealth = ConvertTo-NinjaHealth -Raw "$($healthBy[$key])" -Derived (Get-HealthStatus -Online $online -MaxSev $maxSev -Failed $fCount -Reboot $reboot) -HaveData ($healthBy.Count -gt 0)
            PatchEnabled = [bool]($patchDev.ContainsKey($key) -or $lastScan[$key] -or $lastApply[$key])
            ApprovedN  = $(if ($approvedBy.ContainsKey($key))  { @($approvedBy[$key]).Count }  else { 0 })
            InstalledN = $(if ($installedBy.ContainsKey($key)) { @($installedBy[$key]).Count } else { 0 })
            Threats   = $threatsBy[$key]
            Backup    = $backupBy[$key]
            Disk      = $diskBy[$key]
            LastBoot  = $(if ($osBy.ContainsKey($key)) { ConvertFrom-NinjaTime $osBy[$key].lastBootTime })
            LastScan  = $lastScan[$key]
            LastApply = $lastApply[$key]
        }
    }
    return @($rows)
}
#endregion

#region HTML helpers
function Enc([string]$t) { [Net.WebUtility]::HtmlEncode($t) }

function New-Donut {
    param($Segments, [string]$CenterTop, [string]$CenterBottom)
    $total = 0; foreach ($s in $Segments) { $total += $s.Value }
    $svg = '<svg viewBox="0 0 42 42" class="donut" role="img">' +
           '<circle cx="21" cy="21" r="15.9155" fill="none" stroke="#2a3140" stroke-width="5"/>'
    $offset = 25.0
    if ($total -gt 0) {
        foreach ($s in $Segments) {
            if ($s.Value -le 0) { continue }
            $pct = 100.0 * $s.Value / $total
            $p   = $pct.ToString('0.####', $Inv)
            $r   = (100 - $pct).ToString('0.####', $Inv)
            $o   = $offset.ToString('0.####', $Inv)
            $svg += "<circle cx=`"21`" cy=`"21`" r=`"15.9155`" fill=`"none`" stroke=`"$($s.Color)`" stroke-width=`"5`" stroke-dasharray=`"$p $r`" stroke-dashoffset=`"$o`"/>"
            $offset -= $pct
        }
    }
    $svg += "<text x=`"21`" y=`"21.5`" text-anchor=`"middle`" class=`"c1`">$(Enc $CenterTop)</text>"
    $svg += "<text x=`"21`" y=`"26.5`" text-anchor=`"middle`" class=`"c2`">$(Enc $CenterBottom)</text></svg>"

    $legend = '<ul class="legend">'
    foreach ($s in $Segments) {
        $legend += "<li><span class=`"dot`" style=`"background:$($s.Color)`"></span>$(Enc $s.Label)<b>$($s.Value)</b></li>"
    }
    $legend += '</ul>'
    return "<div class=`"donutwrap`">$svg$legend</div>"
}

# NinjaOne alerts often carry severity NONE; the real signal is conditionHealthStatus
# (UNHEALTHY / NEEDS_ATTENTION). Group 0 = red, 1 = yellow, 2 = everything else.
function Get-AlertGroup($a) {
    $sev = "$($a.severity)".ToUpper(); $hs = "$($a.conditionHealthStatus)".ToUpper()
    if (($sev -eq 'CRITICAL') -or ($sev -eq 'MAJOR') -or ($hs -eq 'UNHEALTHY')) { return 0 }
    if (($sev -eq 'MODERATE') -or ($sev -eq 'MINOR') -or ($hs -eq 'NEEDS_ATTENTION')) { return 1 }
    return 2
}

function Get-AlertLabel($a) {
    $sev = "$($a.severity)".ToUpper(); $hs = "$($a.conditionHealthStatus)".ToUpper()
    if ($sev -and $sev -ne 'NONE') { return (Get-Culture).TextInfo.ToTitleCase($sev.ToLower()) }
    if ($hs -eq 'UNHEALTHY')       { return 'Unhealthy' }
    if ($hs -eq 'NEEDS_ATTENTION') { return 'Needs attention' }
    return 'None'
}

function Get-HealthClass([string]$h) {
    switch ($h) {
        'Healthy'         { return 'Healthy'  }
        'Needs attention' { return 'Warning'  }
        'Unhealthy'       { return 'Critical' }
        default           { return 'Offline'  }
    }
}

function Format-Backup {
    param($Job)
    if (-not $Job) { return '<div class="pe"><span class="mute" style="width:auto">Backup Status:</span> <span class="pill Offline" title="No recent backup jobs or backup activity were found for this server">No recent backup data</span></div><div class="pe"><span class="mute" style="width:auto">Last Job:</span> <span class="zero">No data</span></div>' }
    $s = $Job.Status.ToUpper()
    if     ($s -match 'COMPLETE|SUCCESS') { $cls = 'Healthy';  $txt = 'Success' }
    elseif ($s -match 'FAIL')             { $cls = 'Critical'; $txt = 'Failed' }
    else   { $cls = 'Warning'; $txt = (Get-Culture).TextInfo.ToTitleCase($s.ToLower()) }
    return "<div class=`"pe`"><span class=`"mute`" style=`"width:auto`">Backup Status:</span> <span class=`"pill Healthy`">Enabled</span></div>" +
           "<div class=`"pe`"><span class=`"mute`" style=`"width:auto`">Last Job:</span> <span class=`"pill $cls`" title=`"Plan: $(Enc $Job.Plan)`">$txt</span> <span class=`"mute`" style=`"width:auto`">on $($Job.Time.ToString('yyyy-MM-dd HH:mm'))</span></div>"
}

function Format-Disk {
    param($Disks)
    if (-not $Disks -or $Disks.Count -eq 0) { return '<span class="zero">N/A</span>' }
    $out = ''
    foreach ($d in ($Disks | Sort-Object Name)) {
        $cls = ''
        if ($d.Pct -ge 90) { $cls = 'bad' } elseif ($d.Pct -ge 80) { $cls = 'pend' }
        $out += "<div class=`"pe`"><span class=`"mute`" style=`"width:auto`">$(Enc $d.Name)</span> <span class=`"$cls`">$($d.Pct)%</span></div>"
    }
    return $out
}

# Friendly names for maintenance-mode feature codes (matches the labels in the NinjaOne console)
function Format-Feature([string]$f) {
    switch ($f.ToUpper()) {
        'ALERTS'                { return 'Condition alerts' }
        'TASKS'                 { return 'Automations/tasks' }
        'PATCHING'              { return 'Patching' }
        'NOTIFICATIONS'         { return 'Notifications' }
        'NOTIFICATION_CHANNELS' { return 'Notification channels' }
        'AVSCANS'               { return 'Antivirus' }
        default                 { return (Get-Culture).TextInfo.ToTitleCase($f.ToLower().Replace('_',' ')) }
    }
}

function Format-Uptime {
    param($Boot, [bool]$Online)
    if (-not $Online) { return '<span class="zero">Offline</span>' }
    if (-not $Boot)   { return '<span class="zero">N/A</span>' }
    $ts = (Get-Date) - $Boot
    if ($ts.TotalSeconds -lt 0) { return '<span class="zero">N/A</span>' }
    if ($ts.TotalDays -ge 1)      { $txt = '{0}d {1}h' -f [int][math]::Floor($ts.TotalDays), $ts.Hours }
    elseif ($ts.TotalHours -ge 1) { $txt = '{0}h {1}m' -f $ts.Hours, $ts.Minutes }
    else                          { $txt = '{0}m' -f [int]$ts.Minutes }
    return "<span title=`"Last boot $($Boot.ToString('yyyy-MM-dd HH:mm'))`">$txt</span>"
}

function Format-PatchEvent {
    param([string]$Label, $Evt)
    if (-not $Evt) { return "<div class=`"pe`"><span class=`"mute`">$Label</span> <span class=`"zero`">No data</span></div>" }
    $cls = 'Healthy'; $txt = 'Success'
    if ($Evt.Failed) { $cls = 'Critical'; $txt = 'Failed' }
    return "<div class=`"pe`"><span class=`"mute`">$Label</span> $($Evt.Time.ToString('yyyy-MM-dd HH:mm')) <span class=`"pill $cls`">$txt</span></div>"
}

function Get-SevClass([int]$rank) {
    switch ($rank) { 4 {'critical'} 3 {'critical'} 2 {'warning'} 1 {'minor'} default {'minor'} }
}
#endregion

#region Page builder
function Build-Page {
    param($Rows, [string]$ErrorMessage)

    $total   = $Rows.Count
    $online  = @($Rows | Where-Object Online).Count
    $offline = $total - $online
    $healthy = @($Rows | Where-Object { $_.Health -eq 'Healthy' }).Count
    $unhealthy = $total - $healthy

    $order = @{ 'Unknown' = 0; 'Unhealthy' = 1; 'Needs attention' = 2; 'Healthy' = 3 }
    $sorted = $Rows | Sort-Object @{ e = { $order[$_.NinjaHealth] } }, Name

    $sb = New-Object Text.StringBuilder
    $now = Get-Date
    [void]$sb.AppendLine(@"
<!DOCTYPE html>
<html lang="en"><head><meta charset="utf-8">
<meta http-equiv="refresh" content="$RefreshSeconds">
<meta name="viewport" content="width=device-width, initial-scale=1">
<title>NinjaOne Dashboard Monitor</title>
<style>
 :root{--bg:#0f131a;--card:#171d27;--line:#2a3140;--text:#e6eaf2;--mute:#8b95a8;
       --ok:#2fbf71;--warn:#f0b429;--crit:#ef4e4e;--off:#7a8497;--info:#4c9aff}
 *{box-sizing:border-box} body{margin:0;background:var(--bg);color:var(--text);font:14px/1.45 Segoe UI,Roboto,Arial,sans-serif;padding:20px 28px}
 header{display:flex;justify-content:space-between;align-items:baseline;margin-bottom:16px}
 h1{font-size:20px;margin:0} .mute{color:var(--mute);font-size:12px}
 .banner{background:#3a1d1d;border:1px solid var(--crit);color:#ffb4b4;padding:8px 12px;border-radius:6px;margin-bottom:14px}
 .card{background:var(--card);border:1px solid var(--line);border-radius:10px;padding:14px 16px;margin-bottom:20px;overflow:auto}
 table{width:100%;border-collapse:collapse} th,td{text-align:left;padding:8px 10px;border-bottom:1px solid var(--line);vertical-align:top}
 th{color:var(--mute);font-weight:600;font-size:12px;text-transform:uppercase;letter-spacing:.04em;position:sticky;top:0;background:var(--card)}
 tr:last-child td{border-bottom:none}
 .pill{display:inline-block;padding:2px 9px;border-radius:999px;font-size:12px;font-weight:600}
 .Healthy{background:#12372a;color:var(--ok)} .Warning{background:#3b3010;color:var(--warn)}
 .Critical{background:#401b1b;color:var(--crit)} .Offline{background:#262c38;color:var(--off)}
 .badge{display:inline-block;min-width:22px;text-align:center;padding:1px 7px;border-radius:5px;font-size:12px;font-weight:700;margin-right:4px}
 .badge.critical{background:var(--crit);color:#fff} .badge.warning{background:var(--warn);color:#222} .badge.minor{background:var(--info);color:#fff}
 .alertmsg{color:var(--mute);font-size:12px;margin-top:3px;max-width:520px;white-space:nowrap;overflow:hidden;text-overflow:ellipsis}
 .yes{color:var(--warn);font-weight:700} .no{color:var(--mute)} .zero{color:var(--mute)} .bad{color:var(--crit);font-weight:700} .pend{color:var(--warn);font-weight:700}
 .pe{white-space:nowrap;font-size:13px;line-height:1.7} .pe .mute{display:inline-block;width:42px} .pe .pill{font-size:11px;padding:1px 7px;margin-left:4px}
 .banner.w{background:#3b3010;border-color:var(--warn);color:#ffe08a}
 details.card>summary{cursor:pointer;font-size:15px;font-weight:700;list-style:none;display:flex;align-items:center;gap:8px;user-select:none}
 details.card>summary::-webkit-details-marker{display:none}
 details.card>summary::before{content:"\25BE";color:var(--mute);display:inline-block;transition:transform .15s}
 details.card:not([open])>summary::before{transform:rotate(-90deg)}
 details.card[open]>summary{margin-bottom:10px}
 #board{display:flex;flex-wrap:wrap;gap:20px;align-items:flex-start;margin-bottom:20px;position:relative}
 #board.free{display:block}
 #board.free>details.card{position:absolute;margin:0}
 .rbtn{display:block;margin:0 0 10px auto;background:transparent;color:var(--mute);border:1px solid var(--line);border-radius:6px;padding:3px 10px;font-size:12px;cursor:pointer}
 .rbtn:hover{color:var(--text)}
 #board>details.card{width:100%;margin-bottom:0}
 #board>details.card.half{width:calc(50% - 10px);min-width:340px}
 .grip{margin-left:auto;cursor:grab;color:var(--mute);font-size:18px;line-height:1;padding:0 4px}
 .grip:hover{color:var(--text)}
 details.card.dragging{opacity:.45;outline:2px dashed var(--info)}
 details.card{resize:both;overflow:auto;max-width:100%;min-width:260px}
 .patchrows{margin-top:14px}
 .prow{display:flex;justify-content:space-between;gap:16px;padding:8px 2px;border-top:1px solid var(--line)}
 .prow b{white-space:nowrap}
 .tstats{display:flex;flex-wrap:wrap;gap:12px 0;margin:4px 0 18px}
 .tstat{border-left:3px solid;padding:0 44px 0 12px;min-width:170px}
 .tnum{font-size:26px;font-weight:600;line-height:1.2}
 .tgrid{display:grid;grid-template-columns:repeat(auto-fill,minmax(250px,1fr));gap:16px}
 .tgroup{border-left:3px solid;padding:2px 0 2px 12px}
 .thead{display:flex;justify-content:space-between;gap:12px;font-weight:700;font-size:14px;margin-bottom:4px}
 .tdot{display:inline-block;width:8px;height:8px;border-radius:50%;margin-right:8px}
 .trow{display:flex;justify-content:space-between;gap:12px;padding:3px 0;font-size:13px;border-top:1px solid var(--line)}
 .afilter{display:flex;flex-wrap:wrap;gap:6px 18px;align-items:center;margin:0 0 10px;font-size:13px}
 .afilter label{cursor:pointer;display:inline-flex;align-items:center;gap:6px}
 .al{display:grid;grid-template-columns:125px 130px 1fr;column-gap:18px;align-items:start;padding:3px 0;line-height:1.5}
 .al .d{color:var(--mute);white-space:nowrap}
 .al .pill{display:block;text-align:center}
 .charts{align-items:start}
 .charts{display:grid;grid-template-columns:repeat(auto-fit,minmax(340px,1fr));gap:20px}
 .charts h2{font-size:15px;margin:0 0 10px}
 .donutwrap{display:flex;align-items:center;gap:26px;flex-wrap:nowrap}
 html{color-scheme:dark}
 .donut{width:220px;height:220px;transform:rotate(0deg)} .c1{font-size:7px;font-weight:700;fill:var(--text)} .c2{font-size:2.6px;fill:var(--mute)}
 .legend{list-style:none;margin:0;padding:0} .legend li{display:flex;align-items:center;gap:8px;padding:4px 0}
 .legend b{margin-left:auto;padding-left:24px} .dot{width:12px;height:12px;border-radius:3px;display:inline-block}
</style></head><body>
<header><h1>NinjaOne Dashboard Monitor</h1>
<div class="mute">Updated $($now.ToString('yyyy-MM-dd HH:mm:ss')) &middot; refreshes every $RefreshSeconds s &middot; next in <span id="cd">$RefreshSeconds</span>s</div></header>
"@)

    if ($ErrorMessage) {
        [void]$sb.AppendLine("<div class=`"banner`">Last refresh failed &mdash; showing previous data if any. $(Enc $ErrorMessage)</div>")
    }

    foreach ($w in $script:Warnings) {
        [void]$sb.AppendLine("<div class=`"banner w`">Partial data: $(Enc $w)</div>")
    }

    [void]$sb.AppendLine('<div id="board"><details class="card" id="sec-servers" open><summary>Servers</summary><table><thead><tr><th>Server</th><th>Health</th><th>Status</th><th>Alerts</th><th>Reboot Pending</th><th>Uptime</th><th>Disk Used</th><th>Backup</th><th>Pending Patches</th><th>Failed Patches</th><th>Last Patch Scan / Apply</th></tr></thead><tbody>')

    foreach ($r in $sorted) {
        # Alerts cell
        $alCell = '<span class="zero">None</span>'
        if ($r.Alerts.Count -gt 0) {
            $c = @{ critical = 0; warning = 0; minor = 0 }
            foreach ($a in $r.Alerts) { $gi = Get-AlertGroup $a; if ($gi -eq 0) { $c['critical']++ } elseif ($gi -eq 1) { $c['warning']++ } else { $c['minor']++ } }
            $alCell = ''
            foreach ($k in 'critical','warning','minor') {
                if ($c[$k] -gt 0) { $alCell += "<span class=`"badge $k`" title=`"$k`">$($c[$k])</span>" }
            }
        }

        $rb = if ($r.Reboot) { '<span class="yes">Yes</span>' } else { '<span class="no">No</span>' }
        $pc = if ($r.Pending -gt 0) { "<span class=`"pend`">$($r.Pending)</span>" } else { '<span class="zero">0</span>' }
        $fc = if ($r.Failed  -gt 0) { "<span class=`"bad`">$($r.Failed)</span>"  } else { '<span class="zero">0</span>' }

        [void]$sb.AppendLine("<tr><td><b>$(Enc $r.Name)</b><div class=`"mute`">$(Enc $r.Class)</div></td><td><span class=`"pill $(Get-HealthClass $r.NinjaHealth)`">$($r.NinjaHealth)</span></td><td>$(if ($r.Online) { '<span class="pill Healthy">Online</span>' } else { '<span class="pill Critical">Offline</span>' })</td><td>$alCell</td><td>$rb</td><td>$(Format-Uptime $r.LastBoot $r.Online)</td><td>$(Format-Disk $r.Disk)</td><td>$(Format-Backup $r.Backup)</td><td>$pc</td><td>$fc</td><td>$(Format-PatchEvent 'Scan' $r.LastScan)$(Format-PatchEvent 'Apply' $r.LastApply)</td></tr>")
    }
    if ($total -eq 0) { [void]$sb.AppendLine('<tr><td colspan="11" class="mute">No servers returned.</td></tr>') }
    [void]$sb.AppendLine('</tbody></table><div class="mute" style="margin-top:10px"><b>Alert badges:</b> <span class="dot" style="background:#ef4e4e;width:9px;height:9px;vertical-align:middle"></span> Unhealthy/Critical/Major &nbsp; <span class="dot" style="background:#f0b429;width:9px;height:9px;vertical-align:middle"></span> Needs attention/Moderate/Minor &nbsp; <span class="dot" style="background:#4c9aff;width:9px;height:9px;vertical-align:middle"></span> Other</div></details>')

    $d1 = New-Donut -Segments @(
            @{ Label = 'Healthy';         Value = @($Rows | Where-Object { $_.NinjaHealth -eq 'Healthy' }).Count;         Color = '#2fbf71' },
            @{ Label = 'Needs attention'; Value = @($Rows | Where-Object { $_.NinjaHealth -eq 'Needs attention' }).Count; Color = '#f0b429' },
            @{ Label = 'Unhealthy';       Value = @($Rows | Where-Object { $_.NinjaHealth -eq 'Unhealthy' }).Count;       Color = '#ef4e4e' },
            @{ Label = 'Unknown';         Value = @($Rows | Where-Object { $_.NinjaHealth -eq 'Unknown' }).Count;         Color = '#7a8497' }
          ) -CenterTop "$(@($Rows | Where-Object { $_.NinjaHealth -eq 'Healthy' }).Count)/$total" -CenterBottom 'healthy'

    $d2 = New-Donut -Segments @(
            @{ Label = 'Online servers';  Value = $online;  Color = '#4c9aff' },
            @{ Label = 'Offline servers'; Value = $offline; Color = '#7a8497' },
            @{ Label = 'Total servers';   Value = $total;   Color = '#2a3140' }
          ) -CenterTop "$total" -CenterBottom 'total servers'
    # The "Total" legend row is informational; strip it from the ring itself:
    $d2 = New-Donut -Segments @(
            @{ Label = 'Online servers';  Value = $online;  Color = '#2fbf71' },
            @{ Label = 'Offline servers'; Value = $offline; Color = '#ef4e4e' }
          ) -CenterTop "$total" -CenterBottom 'total servers'
    $d2 = $d2.Replace('</ul>', "<li><span class=`"dot`" style=`"background:transparent;border:1px solid #8b95a8`"></span>Total servers<b>$total</b></li></ul>")

    [void]$sb.AppendLine("<details class=`"card half`" id=`"sec-health`" open><summary>All Servers vs Healthy Servers</summary>$d1</details><details class=`"card half`" id=`"sec-online`" open><summary>Online vs Offline Servers</summary>$d2</details>")

    # ---- Patching compliance: patch-enabled servers that are fully patched (rejected patches are ignored) ----
    $enabledRows = @($Rows | Where-Object { $_.PatchEnabled })
    $nEn         = $enabledRows.Count
    $fullyN      = @($enabledRows | Where-Object { $_.Pending -eq 0 -and $_.Failed -eq 0 }).Count
    $compPct     = 0.0; if ($nEn -gt 0)  { $compPct   = 100.0 * $fullyN / $nEn }
    $enablePct   = 0.0; if ($total -gt 0) { $enablePct = 100.0 * $nEn / $total }
    $cut         = (Get-Date).AddDays(-30)
    $scanFailN   = @($enabledRows | Where-Object { $_.LastScan -and $_.LastScan.Failed }).Count
    $noScanN     = @($enabledRows | Where-Object { (-not $_.LastScan) -or ($_.LastScan.Time -lt $cut) }).Count
    $failPatchN  = @($enabledRows | Where-Object { $_.Failed -gt 0 }).Count
    $dash        = { param($n) if ($n -gt 0) { "$n" } else { '-' } }

    $d3 = New-Donut -Segments @(
            @{ Label = 'Fully patched';     Value = $fullyN;          Color = '#2fbf71' },
            @{ Label = 'Not fully patched'; Value = ($nEn - $fullyN); Color = '#ef4e4e' }
          ) -CenterTop ($compPct.ToString('0.00', $Inv) + '%') -CenterBottom 'compliant'

    $prows = "<div class=`"prow`"><span>Device patch enablement</span><b>$($enablePct.ToString('0.00', $Inv))% <span class=`"mute`">($nEn of $total)</span></b></div>" +
             "<div class=`"prow`"><span>Devices with last scan failed</span><b>$(& $dash $scanFailN)</b></div>" +
             "<div class=`"prow`"><span>Devices with no scans in 30 days</span><b>$(& $dash $noScanN)</b></div>" +
             "<div class=`"prow`"><span>Devices with failed patches</span><b>$(& $dash $failPatchN)</b></div>"

    [void]$sb.AppendLine("<details class=`"card half`" id=`"sec-patch`" open><summary>OS Patching Compliance</summary>$d3<div class=`"patchrows`">$prows</div><div class=`"mute`" style=`"margin-top:8px;font-size:12px`">Patch-enabled = servers with patch records or scan/apply activity. Rejected patches are ignored.</div></details>")

    # ---- Patches installed: Installed / Approved / Failed / Pending across all servers ----
    $instN     = [int](($Rows | Measure-Object -Property InstalledN -Sum).Sum)
    $apprN     = [int](($Rows | Measure-Object -Property ApprovedN -Sum).Sum)
    $failN     = [int](($Rows | Measure-Object -Property Failed -Sum).Sum)
    $pendTotal = [int](($Rows | Measure-Object -Property Pending -Sum).Sum)
    $pendN     = [Math]::Max(0, $pendTotal - $apprN)
    $allN      = $instN + $apprN + $failN + $pendN
    $instPct   = 0.0; if ($allN -gt 0) { $instPct = 100.0 * $instN / $allN }

    $d4 = New-Donut -Segments @(
            @{ Label = 'Installed'; Value = $instN;       Color = '#0f7a3d' },
            @{ Label = 'Approved';  Value = $apprN;       Color = '#5b6470' },
            @{ Label = 'Failed';    Value = $failN;       Color = '#c13b45' },
            @{ Label = 'Pending';   Value = $pendN;       Color = '#a8860b' }
          ) -CenterTop ($instPct.ToString('0.00', $Inv) + '%') -CenterBottom 'installed'

    [void]$sb.AppendLine("<details class=`"card half`" id=`"sec-installed`" open><summary>OS Patches Installed</summary>$d4<div class=`"mute`" style=`"margin-top:8px;font-size:12px`">Installed = patches installed in the last 90 days. Rejected patches are ignored.</div></details>")

    # ---- Software patches: Installed / Approved / Failed / Pending ----
    $sw = $script:SwPatches
    if (-not $sw) {
        [void]$sb.AppendLine("<details class=`"card half`" id=`"sec-swpatch`" open><summary>Software Patches Installed</summary><div class=`"mute`">Software patch data is unavailable (see the Partial data banner at the top).</div></details>")
    } else {
        $swAll = $sw.Installed + $sw.Approved + $sw.Failed + $sw.Pending
        $swPct = 0.0; if ($swAll -gt 0) { $swPct = 100.0 * $sw.Installed / $swAll }
        $d5 = New-Donut -Segments @(
                @{ Label = 'Installed'; Value = $sw.Installed; Color = '#0f7a3d' },
                @{ Label = 'Approved';  Value = $sw.Approved;  Color = '#5b6470' },
                @{ Label = 'Failed';    Value = $sw.Failed;    Color = '#c13b45' },
                @{ Label = 'Pending';   Value = $sw.Pending;   Color = '#a8860b' }
              ) -CenterTop ($swPct.ToString('0.00', $Inv) + '%') -CenterBottom 'installed'
        [void]$sb.AppendLine("<details class=`"card half`" id=`"sec-swpatch`" open><summary>Software Patches Installed</summary>$d5<div class=`"mute`" style=`"margin-top:8px;font-size:12px`">Installed = software patches installed in the last 90 days. Rejected patches are ignored.</div></details>")
    }

    # ---- NMS devices: up/down donut + grid of devices that are down ----
    $nms     = @($script:Nms | Where-Object { $_ })
    $nmsUp   = @($nms | Where-Object { $_.Online }).Count
    $nmsDown = @($nms | Where-Object { -not $_.Online })
    $nmsTot  = $nms.Count
    $nmsPct  = 0.0; if ($nmsTot -gt 0) { $nmsPct = 100.0 * $nmsUp / $nmsTot }
    $dn = New-Donut -Segments @(
            @{ Label = 'NMS devices up';   Value = $nmsUp;           Color = '#2fbf71' },
            @{ Label = 'NMS devices down'; Value = $nmsDown.Count;   Color = '#ef4e4e' }
          ) -CenterTop "$nmsTot" -CenterBottom 'NMS devices'
    $dn = $dn.Replace('</ul>', "<li><span class=`"dot`" style=`"background:transparent;border:1px solid #8b95a8`"></span>Total NMS devices<b>$nmsTot</b></li></ul>")
    [void]$sb.AppendLine("<details class=`"card half`" id=`"sec-nmsdonut`" open><summary>NMS Devices Up vs Down</summary>$dn</details>")

    [void]$sb.AppendLine("<details class=`"card half`" id=`"sec-nmsdown`" open><summary>NMS Devices Down ($($nmsDown.Count))</summary>")
    if ($nmsDown.Count -eq 0) {
        [void]$sb.AppendLine("<div class=`"mute`">All NMS devices are up.</div>")
    } else {
        [void]$sb.AppendLine("<table><thead><tr><th>Device</th><th>Type</th><th>Organization</th><th>Last uptime</th></tr></thead><tbody>")
        foreach ($n in ($nmsDown | Sort-Object @{ Expression = { if ($_.LastSeen) { $_.LastSeen } else { [datetime]::MinValue } } }, Name)) {
            $lu = '<span class="zero">Unknown</span>'
            if ($n.LastSeen) {
                $ago = (Get-Date) - $n.LastSeen
                if     ($ago.TotalDays -ge 1)  { $agoTxt = '{0}d {1}h ago' -f [int][math]::Floor($ago.TotalDays), $ago.Hours }
                elseif ($ago.TotalHours -ge 1) { $agoTxt = '{0}h {1}m ago' -f $ago.Hours, $ago.Minutes }
                else                           { $agoTxt = '{0}m ago' -f [int][math]::Max(0, $ago.Minutes) }
                $lu = "$($n.LastSeen.ToString('yyyy-MM-dd HH:mm')) <span class=`"mute`">($agoTxt)</span>"
            }
            [void]$sb.AppendLine("<tr><td><b>$(Enc $n.Name)</b></td><td>$(Enc $n.Class)</td><td>$(Enc $n.Org)</td><td>$lu</td></tr>")
        }
        [void]$sb.AppendLine("</tbody></table>")
    }
    [void]$sb.AppendLine("<div class=`"mute`" style=`"margin-top:8px;font-size:12px`">Last uptime = last time NinjaOne heard from the device. NMS devices are discovered automatically from every NMS device class.</div></details>")

    # ---- Cloud monitors: up/down donut + grid of monitors in an alert state ----
    $cmDev = @(); $cmAl = @()
    if ($script:Cloud) { $cmDev = @($script:Cloud.Devices | Where-Object { $_ }); $cmAl = @($script:Cloud.Alerts | Where-Object { $_ }) }
    $cmUp   = @($cmDev | Where-Object { $_.Online }).Count
    $cmDown = $cmDev.Count - $cmUp
    $dc = New-Donut -Segments @(
            @{ Label = 'Cloud monitors up';   Value = $cmUp;   Color = '#2fbf71' },
            @{ Label = 'Cloud monitors down'; Value = $cmDown; Color = '#ef4e4e' }
          ) -CenterTop "$($cmDev.Count)" -CenterBottom 'cloud monitors'
    $dc = $dc.Replace('</ul>', "<li><span class=`"dot`" style=`"background:transparent;border:1px solid #8b95a8`"></span>Total cloud monitors<b>$($cmDev.Count)</b></li></ul>")
    $cmNote = ''; if (-not $script:Cloud) { $cmNote = "<div class=`"mute`">Cloud monitor data is unavailable (see the Partial data banner at the top).</div>" }
    [void]$sb.AppendLine("<details class=`"card half`" id=`"sec-cloudmon`" open><summary>Cloud Monitors Up vs Down</summary>$dc$cmNote</details>")

    [void]$sb.AppendLine("<details class=`"card half`" id=`"sec-cloudalerts`" open><summary>Cloud Monitors in Alert State ($($cmAl.Count))</summary>")
    if ($cmAl.Count -eq 0) {
        [void]$sb.AppendLine("<div class=`"mute`">No cloud monitors are in an alert state.</div>")
    } else {
        [void]$sb.AppendLine("<table><thead><tr><th>Monitor</th><th>Condition health</th><th>Alert</th><th>Triggered</th></tr></thead><tbody>")
        foreach ($c in ($cmAl | Sort-Object @{ Expression = { $_.Group } }, @{ Expression = { if ($_.When) { $_.When } else { [datetime]::MinValue } }; Descending = $true })) {
            $cc = 'Offline'; if ($c.Group -eq 0) { $cc = 'Critical' } elseif ($c.Group -eq 1) { $cc = 'Warning' }
            $tw = 'N/A'; if ($c.When) { $tw = $c.When.ToString('yyyy-MM-dd HH:mm') }
            [void]$sb.AppendLine("<tr><td><b>$(Enc $c.Name)</b><div class=`"mute`">$(Enc $c.Org)</div></td><td><span class=`"pill $cc`">$(Enc $c.Health)</span></td><td>$(Enc $c.Message)</td><td>$tw</td></tr>")
        }
        [void]$sb.AppendLine("</tbody></table>")
    }
    [void]$sb.AppendLine("</details>")

    # ---- Backup overview: last result per device, integrity checks and storage ----
    $bo = $script:BackupOv
    if ($bo) {
        $tile = { param($label, $num, $color) "<div class=`"tstat`" style=`"border-color:$color;min-width:90px;padding-right:28px`"><div class=`"mute`">$label</div><div class=`"tnum`">$num</div></div>" }
        $jt = (& $tile 'Success' $bo.Jobs.Success '#19c37d') + (& $tile 'Warnings' $bo.Jobs.Warnings '#c9a227') + (& $tile 'Failed' $bo.Jobs.Failed '#c13b45') + (& $tile 'Running' $bo.Jobs.Running '#19b5d4')
        $it = (& $tile 'Success' $bo.Integrity.Success '#19c37d') + (& $tile 'Warnings' $bo.Integrity.Warnings '#c9a227') + (& $tile 'Failed' $bo.Integrity.Failed '#c13b45')
        $ut = (& $tile 'All devices' $bo.Devices '#7a8497') + (& $tile 'Local storage' (Format-Bytes $bo.Local) '#7a8497') + (& $tile 'Cloud storage' (Format-Bytes $bo.Cloud) '#7a8497')
        $itSub = 'unavailable'; if ($bo.IntegrityOk) { $itSub = "$($bo.IntegrityN) plans" }
        [void]$sb.AppendLine("<details class=`"card`" id=`"sec-backupov`" open><summary>Backup Overview</summary><div style=`"display:flex;flex-wrap:wrap;gap:20px 36px`"><div><div class=`"thead`"><span>Backup jobs</span><span class=`"mute`">Last result</span></div><div class=`"tstats`" style=`"margin:4px 0 0`">$jt</div></div><div><div class=`"thead`"><span>Integrity check last result</span><span class=`"mute`">$itSub</span></div><div class=`"tstats`" style=`"margin:4px 0 0`">$it</div></div><div><div class=`"thead`"><span>Backup usage</span><span class=`"mute`">All devices</span></div><div class=`"tstats`" style=`"margin:4px 0 0`">$ut</div></div></div><div class=`"mute`" style=`"margin-top:10px;font-size:12px`">Backup jobs = most recent result per device and backup plan, across all devices. Storage = total backup usage across all organizations (as in NinjaOne Backup &gt; Device overview). The API's job list only returns completed cloud jobs from the last several days, so job counts can be lower than the console's.</div></details>")
    } else {
        [void]$sb.AppendLine("<details class=`"card`" id=`"sec-backupov`" open><summary>Backup Overview</summary><div class=`"mute`">Backup data is unavailable.</div></details>")
    }

    # ---- BitLocker: Windows desktops and laptops with BitLocker enabled vs total ----
    $bl = $script:BitLocker
    if ($bl) {
        $blPct = 0.0; if ($bl.Total -gt 0) { $blPct = 100.0 * $bl.Enabled / $bl.Total }
        $dbl = New-Donut -Segments @(
                @{ Label = 'BitLocker enabled';  Value = $bl.Enabled;  Color = '#2fbf71' },
                @{ Label = 'BitLocker disabled'; Value = $bl.Disabled; Color = '#ef4e4e' },
                @{ Label = 'No volume data';     Value = $bl.NoData;   Color = '#7a8497' }
              ) -CenterTop ($blPct.ToString('0.0', $Inv) + '%') -CenterBottom 'encrypted'
        $dbl = $dbl.Replace('</ul>', "<li><span class=`"dot`" style=`"background:transparent;border:1px solid #8b95a8`"></span>Total desktops &amp; laptops<b>$($bl.Total)</b></li></ul>")
        [void]$sb.AppendLine("<details class=`"card half`" id=`"sec-bitlocker`" open><summary>BitLocker: Windows Desktops &amp; Laptops</summary>$dbl<div class=`"mute`" style=`"margin-top:8px;font-size:12px`">Enabled = system drive protected by BitLocker. No volume data = device has not reported volumes.</div></details>")
    } else {
        [void]$sb.AppendLine("<details class=`"card half`" id=`"sec-bitlocker`" open><summary>BitLocker: Windows Desktops &amp; Laptops</summary><div class=`"mute`">BitLocker data is unavailable (see the Partial data banner at the top).</div></details>")
    }

    # ---- Threats: active / blocked / quarantined (most urgent first, newest first within each) ----
    $thRows = @(foreach ($r in $Rows) { foreach ($t in @($r.Threats)) { if ($t) { [pscustomobject]@{ Server = $r.Name; T = $t } } } })
    $nAct = @($thRows | Where-Object { $_.T.Status -eq 'Active' }).Count
    $nQua = @($thRows | Where-Object { $_.T.Status -eq 'Quarantined' }).Count
    $nBlk = @($thRows | Where-Object { $_.T.Status -eq 'Blocked' }).Count
    $thOrder  = @{ 'Active' = 0; 'Quarantined' = 1; 'Blocked' = 2 }
    $thSorted = $thRows | Sort-Object @{ e = { $thOrder[$_.T.Status] } }, @{ e = { if ($_.T.When) { $_.T.When.Ticks } else { 0 } }; Descending = $true }
    $thFailed = (@($script:Warnings | Where-Object { $_ -like 'Antivirus threats*' }).Count -gt 0)
    $thChips  = "<span class=`"pill Critical`">Active $nAct</span> <span class=`"pill Warning`">Quarantined $nQua</span> <span class=`"pill Offline`">Blocked $nBlk</span>"

    [void]$sb.AppendLine("<details class=`"card`" id=`"sec-threats`" open><summary>Threats ($($thRows.Count))</summary><div class=`"afilter`">$thChips</div><table><thead><tr><th>Server</th><th>Status</th><th>Vendor</th><th>Category</th><th>Threat</th><th>Detected</th></tr></thead><tbody>")
    foreach ($x in $thSorted) {
        $tc = 'Offline'; if ($x.T.Status -eq 'Active') { $tc = 'Critical' } elseif ($x.T.Status -eq 'Quarantined') { $tc = 'Warning' }
        $tw = 'N/A'; if ($x.T.When) { $tw = $x.T.When.ToString('yyyy-MM-dd HH:mm') }
        [void]$sb.AppendLine("<tr><td><b>$(Enc $x.Server)</b></td><td><span class=`"pill $tc`" title=`"$(Enc $x.T.Raw)`">$($x.T.Status)</span></td><td>$(Enc $x.T.Vendor)</td><td>$(Enc $x.T.Category)</td><td>$(Enc $x.T.Name)</td><td style=`"white-space:nowrap`">$tw</td></tr>")
    }
    if ($thRows.Count -eq 0) {
        $emptyMsg = 'No active, blocked or quarantined threats.'
        if ($thFailed) { $emptyMsg = 'Threat data could not be retrieved (see the Partial data banner at the top).' }
        [void]$sb.AppendLine("<tr><td colspan=`"6`" class=`"mute`">$emptyMsg</td></tr>")
    }
    [void]$sb.AppendLine('</tbody></table></details>')

    # ---- Threat health chart: servers with an active threat = unhealthy ----
    $withActive = @($Rows | Where-Object { @($_.Threats | Where-Object { $_ -and $_.Status -eq 'Active' }).Count -gt 0 }).Count
    $thHealthy  = $total - $withActive
    $d5 = New-Donut -Segments @(
            @{ Label = 'Healthy (no active threats)'; Value = $thHealthy;  Color = '#2fbf71' },
            @{ Label = 'Unhealthy (active threats)';  Value = $withActive; Color = '#ef4e4e' }
          ) -CenterTop "$thHealthy/$total" -CenterBottom 'healthy'
    $thNote = 'Unhealthy = server has at least one Active threat. Blocked and quarantined threats are not counted against health.'
    if ($thFailed) { $thNote = 'Threat data could not be retrieved, so every server shows as healthy. See the Partial data banner at the top.' }
    [void]$sb.AppendLine("<details class=`"card half`" id=`"sec-threatchart`" open><summary>Healthy vs Unhealthy Servers by Threat</summary>$d5<div class=`"mute`" style=`"margin-top:8px;font-size:12px`">$thNote</div></details>")

    # ---- Devices in maintenance mode (all device types, not just servers) ----
    $mRows   = @($script:Maintenance | Where-Object { $_ })
    $mFailed = (@($script:Warnings | Where-Object { $_ -like 'Maintenance mode*' }).Count -gt 0)
    $mSorted = $mRows | Sort-Object @{ e = { if ($_.Status -match 'IN_MAINT|ACTIVE') { 0 } else { 1 } } }, Name
    [void]$sb.AppendLine("<details class=`"card`" id=`"sec-maint`" open><summary>Devices in Maintenance Mode ($($mRows.Count))</summary><table><thead><tr><th>Device</th><th>Organization</th><th>Status</th><th>Start</th><th>End</th><th>Features suppressed</th></tr></thead><tbody>")
    foreach ($m in $mSorted) {
        $ms = 'N/A';   if ($m.Start) { $ms = $m.Start.ToString('yyyy-MM-dd HH:mm') }
        $me = 'Never'; if ($m.End)   { $me = $m.End.ToString('yyyy-MM-dd HH:mm') }
        $mc = 'Offline'; $mt = 'Unknown'
        if     ("$($m.Status)" -match 'IN_MAINT|ACTIVE') { $mc = 'Warning'; $mt = 'Active' }
        elseif ("$($m.Status)" -match 'PEND|SCHED')      { $mt = 'Scheduled' }
        elseif ("$($m.Status)")                           { $mt = (Get-Culture).TextInfo.ToTitleCase(("$($m.Status)").ToLower().Replace('_',' ')) }
        $mf = ''
        foreach ($f in @($m.Features)) {
            if ($f) { $mf += "<span class=`"pill Offline`" style=`"font-size:11px;padding:1px 7px;margin:0 4px 2px 0`">$(Enc (Format-Feature "$f"))</span>" }
        }
        if (-not $mf) { $mf = '<span class="zero">-</span>' }
        [void]$sb.AppendLine("<tr><td><b>$(Enc $m.Name)</b><div class=`"mute`">$(Enc $m.Class)</div></td><td>$(Enc $m.Org)</td><td><span class=`"pill $mc`">$mt</span></td><td style=`"white-space:nowrap`">$ms</td><td style=`"white-space:nowrap`">$me</td><td>$mf</td></tr>")
    }
    if ($mRows.Count -eq 0) {
        $mMsg = 'No devices are in maintenance mode.'
        if ($mFailed) { $mMsg = 'Maintenance data could not be retrieved (see the Partial data banner at the top).' }
        [void]$sb.AppendLine("<tr><td colspan=`"6`" class=`"mute`">$mMsg</td></tr>")
    }
    [void]$sb.AppendLine('</tbody></table></details>')

    # ---- Ticketing: Open / Waiting / Unassigned plus every status and child status ----
    $tk      = $script:Tickets
    $tFailed = (@($script:Warnings | Where-Object { $_ -like 'Ticketing*' }).Count -gt 0)
    if (-not $tk) {
        $tMsg = 'No ticket data yet.'
        if ($tFailed) { $tMsg = 'Ticket data could not be retrieved (see the Partial data banner at the top).' }
        [void]$sb.AppendLine("<details class=`"card`" id=`"sec-tickets`" open><summary>Ticketing</summary><div class=`"mute`">$tMsg</div></details>")
    } else {
        $palette = @('#4c9aff','#a77bf3','#e05fa0','#e8833a','#7a8497')
        $tstats = "<div class=`"tstats`">"
        $gi = 0
        foreach ($g in $tk.Groups) {
            if (-not $g.Closed) {
                $c = Get-TicketColor $g.Name $g.ApiColor $palette[$gi % $palette.Count]
                $tstats += "<div class=`"tstat`" style=`"border-color:$c`"><div class=`"mute`">$(Enc $g.Name)</div><div class=`"tnum`">$($g.Total.ToString('N0', $Inv))</div></div>"
            }
            $gi++
        }
        $tstats += "<div class=`"tstat`" style=`"border-color:#7a8497`"><div class=`"mute`">Unassigned</div><div class=`"tnum`">$($tk.Unassigned.ToString('N0', $Inv))</div></div></div>"
        $gh = ''
        $gi = 0
        foreach ($g in $tk.Groups) {
            $col = Get-TicketColor $g.Name $g.ApiColor $palette[$gi % $palette.Count]; $gi++
            $items = ''
            if ($g.Children.Count -gt 0) {
                $items += "<div class=`"trow`"><span class=`"mute`"><i class=`"tdot`" style=`"background:$col`"></i>$(Enc $g.Name) (no sub-status)</span><b>$($g.Own.ToString('N0', $Inv))</b></div>"
                foreach ($c in $g.Children) { $cc = Get-TicketColor $c.Name $c.ApiColor $col; $items += "<div class=`"trow`"><span><i class=`"tdot`" style=`"background:$cc`"></i>$(Enc $c.Name)</span><b>$($c.Count.ToString('N0', $Inv))</b></div>" }
            }
            $gh += "<div class=`"tgroup`" style=`"border-color:$col`"><div class=`"thead`"><span>$(Enc $g.Name)</span><b>$($g.Total.ToString('N0', $Inv))</b></div>$items</div>"
        }
        [void]$sb.AppendLine("<details class=`"card`" id=`"sec-tickets`" open><summary>Ticketing ($($tk.Total.ToString('N0', $Inv)) tickets)</summary>$tstats<div class=`"tgrid`">$gh</div><div class=`"mute`" style=`"margin-top:10px;font-size:12px`">Statuses are read live from NinjaOne, including custom ones; parent totals include their child statuses. Unassigned excludes Resolved/Closed-type tickets. Counts refresh every 5 minutes (last read $($tk.Time.ToString('HH:mm:ss'))).</div></details>")
    }

    # ---- Active alerts section: red (Critical/Major) first, then yellow (Moderate/Minor), then the rest; newest first within each ----
    $sevGroup  = @{ CRITICAL = 0; MAJOR = 0; MODERATE = 1; MINOR = 1 }
    $alertRows = @(foreach ($r in $Rows) {
        foreach ($a in $r.Alerts) {
            $sev = "$($a.severity)".ToUpper()
            $g = Get-AlertGroup $a
            $msg = "$($a.message)"; if (-not $msg) { $msg = "$($a.subject)" }
            [pscustomobject]@{ Group = $g; Sev = (Get-AlertLabel $a); When = (ConvertFrom-NinjaTime $(if ($a.createTime) { $a.createTime } else { $a.createdTime })); Server = $r.Name; Message = $msg }
        }
    })
    # One entry per server, listing all of its alerts (worst first, then newest first)
    $alertSorted = @(foreach ($sg in ($alertRows | Group-Object Server)) {
        $items = @($sg.Group | Sort-Object @{ e = { $_.Group }; Descending = $false }, @{ e = { if ($_.When) { $_.When.Ticks } else { 0 } }; Descending = $true })
        $ahtml = ''
        foreach ($it in $items) {
            $ac = 'Offline'; if ($it.Group -eq 0) { $ac = 'Critical' } elseif ($it.Group -eq 1) { $ac = 'Warning' }
            $aw = 'N/A'; if ($it.When) { $aw = $it.When.ToString('yyyy-MM-dd HH:mm') }
            $as = (Get-Culture).TextInfo.ToTitleCase($it.Sev.ToLower()); if (-not $as) { $as = 'None' }
            $ahtml += "<div class=`"al`" data-g=`"$($it.Group)`"><span class=`"d`">$aw</span><span class=`"pill $ac`">$(Enc $as)</span><span>$(Enc $it.Message)</span></div>"
        }
        [pscustomobject]@{ Server = $sg.Name; Group = $items[0].Group; Sev = $items[0].Sev; When = $items[0].When; Items = $items; Html = $ahtml }
    })
    $alertSorted = $alertSorted | Sort-Object @{ e = { $_.Group }; Descending = $false }, @{ e = { if ($_.When) { $_.When.Ticks } else { 0 } }; Descending = $true }

    [void]$sb.AppendLine("<details class=`"card`" id=`"sec-alerts`" open><summary>Active Alerts ($($alertRows.Count))</summary><div class=`"afilter`"><b>Show:</b> <label><input type=`"checkbox`" class=`"afc`" data-g=`"0`" checked> <span class=`"dot`" style=`"background:#ef4e4e;width:9px;height:9px`"></span> Unhealthy / Critical</label> <label><input type=`"checkbox`" class=`"afc`" data-g=`"1`" checked> <span class=`"dot`" style=`"background:#f0b429;width:9px;height:9px`"></span> Needs attention</label> <label><input type=`"checkbox`" class=`"afc`" data-g=`"2`" checked> <span class=`"dot`" style=`"background:#7a8497;width:9px;height:9px`"></span> Other</label> <span class=`"mute`" id=`"acnt`"></span></div><table><thead><tr><th>Server</th><th style=`"text-align:center`">Worst Severity</th><th><div class=`"al`"><span>Date Triggered</span><span style=`"text-align:center`">Severity</span><span>Description</span></div></th></tr></thead><tbody>")
    foreach ($al in $alertSorted) {
        $cls = 'Offline'; if ($al.Group -eq 0) { $cls = 'Critical' } elseif ($al.Group -eq 1) { $cls = 'Warning' }
        $ts  = 'N/A'; if ($al.When) { $ts = $al.When.ToString('yyyy-MM-dd HH:mm:ss') }
        $sevTxt = (Get-Culture).TextInfo.ToTitleCase($al.Sev.ToLower()); if (-not $sevTxt) { $sevTxt = 'None' }
        [void]$sb.AppendLine("<tr class=`"arow`"><td style=`"vertical-align:top`"><b>$(Enc $al.Server)</b> <span class=`"mute`">($($al.Items.Count))</span></td><td style=`"vertical-align:top;text-align:center`"><span class=`"pill $cls`">$(Enc $sevTxt)</span></td><td>$($al.Html)</td></tr>")
    }
    if ($alertRows.Count -eq 0) { [void]$sb.AppendLine('<tr><td colspan="3" class="mute">No active alerts.</td></tr>') }
    [void]$sb.AppendLine('</tbody></table></details></div>')

    
    [void]$sb.AppendLine("<script>var s=$RefreshSeconds,e=document.getElementById('cd');setInterval(function(){s=Math.max(0,s-1);e.textContent=s;},1000);(function(){var K='nj-dash-open:';document.querySelectorAll('details.card[id]').forEach(function(d){try{if(localStorage.getItem(K+d.id)==='0'){d.open=false;}}catch(x){}d.addEventListener('toggle',function(){try{localStorage.setItem(K+d.id,d.open?'1':'0');}catch(x){}});});})();(function(){var K='nj-dash-af:';var cbs=document.querySelectorAll('.afc');function apply(){var on={};cbs.forEach(function(c){on[c.getAttribute('data-g')]=c.checked;});var shown=0,total=0;document.querySelectorAll('.arow').forEach(function(r){var any=false;r.querySelectorAll('.al[data-g]').forEach(function(a){total++;var v=on[a.getAttribute('data-g')];a.style.display=v?'':'none';if(v){any=true;shown++;}});r.style.display=any?'':'none';});var n=document.getElementById('acnt');if(n){n.textContent='Showing '+shown+' of '+total+' alerts';}}cbs.forEach(function(c){try{if(localStorage.getItem(K+c.getAttribute('data-g'))==='0'){c.checked=false;}}catch(x){}c.addEventListener('change',function(){try{localStorage.setItem(K+c.getAttribute('data-g'),c.checked?'1':'0');}catch(x){}apply();});});apply();})();(function(){var K='nj-dash-size:';document.querySelectorAll('details.card[id]').forEach(function(d){var saved=null;try{saved=JSON.parse(localStorage.getItem(K+d.id)||'null');}catch(x){}function restore(){if(saved&&d.open){if(saved.w){d.style.width=saved.w;}if(saved.h){d.style.height=saved.h;}}}restore();d.addEventListener('toggle',function(){if(d.open){restore();}else{d.style.height='';}});if(window.ResizeObserver){new ResizeObserver(function(){if(!d.open){return;}if(!d.style.width&&!d.style.height){return;}saved={w:d.style.width,h:d.style.height};try{localStorage.setItem(K+d.id,JSON.stringify(saved));}catch(x){}}).observe(d);}});})();(function(){var board=document.getElementById('board');if(!board){return;}var K='nj-dash-pos';var GAP=20;var z=10;function cards(){return Array.prototype.slice.call(board.children).filter(function(c){return c.tagName==='DETAILS';});}var saved={};try{saved=JSON.parse(localStorage.getItem(K)||'{}');}catch(x){}function save(){var o={};cards().forEach(function(c){o[c.id]={x:parseFloat(c.style.left)||0,y:parseFloat(c.style.top)||0,z:parseInt(c.style.zIndex)||0};});try{localStorage.setItem(K,JSON.stringify(o));}catch(x){}}function fitBoard(){var h=0;cards().forEach(function(c){h=Math.max(h,(parseFloat(c.style.top)||0)+c.offsetHeight);});board.style.height=(h+GAP)+'px';}var nat={};cards().forEach(function(c){nat[c.id]={x:c.offsetLeft,y:c.offsetTop,w:c.offsetWidth};});board.classList.add('free');cards().forEach(function(c){var n=nat[c.id];var p=saved[c.id]||{x:n.x,y:n.y,z:0};if(!c.style.width){c.style.width=n.w+'px';}var maxX=Math.max(0,board.clientWidth-n.w);c.style.left=Math.max(0,Math.min(p.x,maxX))+'px';c.style.top=Math.max(0,p.y)+'px';if(p.z){c.style.zIndex=p.z;z=Math.max(z,p.z);}var s=c.querySelector('summary');var g=document.createElement('span');g.className='grip';g.title='Drag to move this section';g.textContent='\u283F';s.appendChild(g);g.addEventListener('click',function(e){e.preventDefault();e.stopPropagation();});g.addEventListener('pointerdown',function(e){if(e.button!==0){return;}e.preventDefault();var sx=e.clientX,sy=e.clientY,ox=parseFloat(c.style.left)||0,oy=parseFloat(c.style.top)||0;z++;c.style.zIndex=z;c.classList.add('dragging');g.setPointerCapture(e.pointerId);function mv(ev){var maxX2=Math.max(0,board.clientWidth-c.offsetWidth);c.style.left=Math.max(0,Math.min(maxX2,ox+ev.clientX-sx))+'px';c.style.top=Math.max(0,oy+ev.clientY-sy)+'px';fitBoard();}function up(){g.removeEventListener('pointermove',mv);g.removeEventListener('pointerup',up);c.classList.remove('dragging');save();}g.addEventListener('pointermove',mv);g.addEventListener('pointerup',up);});if(window.ResizeObserver){new ResizeObserver(fitBoard).observe(c);}c.addEventListener('toggle',fitBoard);});fitBoard();window.addEventListener('resize',fitBoard);var b=document.createElement('button');b.textContent='Reset layout';b.className='rbtn';b.addEventListener('click',function(){try{Object.keys(localStorage).forEach(function(k){if(k.indexOf('nj-dash-pos')===0||k.indexOf('nj-dash-size:')===0||k.indexOf('nj-dash-order')===0){localStorage.removeItem(k);}});}catch(x){}location.reload();});board.parentNode.insertBefore(b,board);})();(function(){document.querySelectorAll('#sec-health,#sec-online,#sec-patch,#sec-installed,#sec-swpatch,#sec-nmsdonut,#sec-cloudmon,#sec-bitlocker,#sec-threatchart').forEach(function(d){var w=d.querySelector('.donutwrap');if(!w){return;}var svg=w.querySelector('svg');var lg=w.querySelector('.legend');function fit(){if(!d.open){return;}var S=220;if(d.style.height||d.style.width){var cs=getComputedStyle(d);var padX=parseFloat(cs.paddingLeft)+parseFloat(cs.paddingRight);var padY=parseFloat(cs.paddingTop)+parseFloat(cs.paddingBottom);var sum=d.querySelector('summary');var extra=0;Array.prototype.forEach.call(d.children,function(ch){if(ch!==sum&&ch!==w){extra+=ch.offsetHeight+(parseFloat(getComputedStyle(ch).marginTop)||0);}});var availH=d.style.height?(d.clientHeight-padY-sum.offsetHeight-extra-24):9999;var availW=d.clientWidth-padX-(lg?lg.offsetWidth:0)-30;S=Math.max(60,Math.min(220,availH,availW));}svg.style.width=S+'px';svg.style.height=S+'px';}if(window.ResizeObserver){new ResizeObserver(fit).observe(d);}d.addEventListener('toggle',fit);fit();});})();</script></body></html>")
    return $sb.ToString()
}
#endregion

#region Main loop
$opened   = $false
$lastRows = @()
Write-Host "NinjaOne dashboard -> $OutputPath  (Ctrl+C to stop)"

do {
    $err = $null
    $script:Warnings.Clear()
    try {
        $lastRows = Get-DashboardData
    } catch {
        $err = $_.Exception.Message
        $script:Token = $null   # force re-auth next cycle
        Write-Warning "Refresh failed: $err"
    }

    $html = Build-Page -Rows $lastRows -ErrorMessage $err
    $tmp  = "$OutputPath.tmp"
    Set-Content -Path $tmp -Value $html -Encoding UTF8
    Move-Item -Path $tmp -Destination $OutputPath -Force
    Write-Host "$(Get-Date -Format 'HH:mm:ss')  wrote $($lastRows.Count) servers"

    if ($OpenBrowser -and -not $opened) { Start-Process $OutputPath; $opened = $true }
    if (-not $Once) { Start-Sleep -Seconds $RefreshSeconds }
} while (-not $Once)
#endregion