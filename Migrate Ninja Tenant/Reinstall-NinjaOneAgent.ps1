<#
.SYNOPSIS
  Reinstalls or migrates the NinjaOne agent — used for repairing a broken agent or moving
  a device to a different NinjaOne instance/organization.

.DESCRIPTION
  Uses NinjaOne's official Agent removal logic (see the Dojo article below), with reinstall
  logic added on afterward. The uninstall + reinstall runs in a separate PowerShell session
  so this script can hand off and exit cleanly.
  Removal guide reference: https://ninjarmm.zendesk.com/hc/en-us/articles/115001836286-NinjaOne-Removal-Guide

    # ---------------------------------------------------------------
    # Author: Mark Giordano
    # Date: 08/22/2024
    # Updated: 1/14/2025  - Include Ninja Remote Removal
    # Updated: 12/30/2025 - Removed reliance on scheduled task. Supports tokenization.
    # ---------------------------------------------------------------

.NOTES
  =====================================================================================
  HOW TO PROVIDE INPUTS TO THIS SCRIPT
  =====================================================================================
  Fill in the values under "USER CONFIGURATION" below, directly in this file, before
  running the script. You do not need to set up anything in NinjaOne's Script
  Variables — just edit the three variables in this file and run it.

  1) $MsiUrl (required, always)
     The download URL of the NinjaOne installer you want to install.
     - If this is a "generic" installer URL (not tied to one specific organization),
       you MUST also fill in $Token below.
     - If this is an organization-specific generated installer URL (NinjaOne
       generates these per-org and they contain a GUID in the URL), leave $Token
       blank.

  2) $Token (required for generic installers, leave blank for org-specific URLs)
     - For a GENERIC installer URL: this must be a valid installer token GUID,
       formatted like 8-4-4-4-12 hex characters (e.g. 1a2b3c4d-1234-5678-9abc-1234567890ab).
     - For a FedRAMP migration specifically: put the target FedRAMP instance's
       ClientUID here instead of a normal token.

  3) $FedRampHostUrl (only required for FedRAMP migrations — otherwise leave blank)
     - Only fill this in if you are migrating this device to a FedRAMP NinjaOne
       instance. In that case, this MUST be the FedRAMP instance's Host URL, AND
       $Token above MUST contain that instance's ClientUID.
     - If you are not doing a FedRAMP migration, leave this blank ('').

  QUICK REFERENCE — WHICH COMBINATION DO I NEED?
  ---------------------------------------------------------------------------------
   Scenario                           | $MsiUrl               | $Token         | $FedRampHostUrl
  ---------------------------------------------------------------------------------
   Normal repair/reinstall             | org-specific install  | leave blank    | leave blank
   (org-generated URL)                 | URL                    |                |
  ---------------------------------------------------------------------------------
   Migrate using a generic installer   | generic installer URL | installer      | leave blank
                                        |                       | token GUID     |
  ---------------------------------------------------------------------------------
   Migrate to a FedRAMP instance       | generic installer URL | FedRAMP        | FedRAMP
                                        |                       | ClientUID      | instance URL
  ---------------------------------------------------------------------------------

  The script will validate your inputs and stop with a clear log message if an
  invalid or conflicting combination is detected (e.g. a Token provided alongside an
  org-specific URL, or a Host URL provided without a matching ClientUID Token).

  All progress and errors are logged to: C:\Windows\Temp\NinjaOneAgentReinstall.log
#>

<#==========================================================================================
  USER CONFIGURATION — EDIT THE VALUES BELOW BEFORE RUNNING THE SCRIPT
==========================================================================================#>

$MsiUrl         = ''   # <-- REQUIRED: paste your installer download URL here (see Section 1 above)
$Token          = ''   # <-- Leave blank for a normal reinstall. Fill in only for a generic installer or FedRAMP ClientUID.
$FedRampHostUrl = ''   # <-- Leave blank unless migrating to a FedRAMP instance.

<#--------------------------------------------------------------------------------------
  SCRIPT FUNCTIONS — no need to edit anything below this line
--------------------------------------------------------------------------------------#>

function New-FileDownload {
  param(
    [Parameter(Mandatory = $true)]
    [ValidateNotNullOrEmpty()]
    [string]$Url,
    [Parameter(Mandatory = $true)]
    [ValidateNotNullOrEmpty()]
    [string]$Destination
  )
  [Net.ServicePointManager]::SecurityProtocol = [Net.SecurityProtocolType]'Tls12,Tls13'
  $webClient = New-Object System.Net.WebClient
  $webClient.DownloadFile($Url, $Destination)
  if (Test-Path -LiteralPath $Destination) {
    Write-Verbose "File downloaded Successfully"
    return $true
  }
  else {
    Write-Verbose "File download Failed"
    return $false
  }
}

function Write-LogEntry {
  param (
    [Parameter(Mandatory = $true)]
    [string]$Message
  )

  $LogPath = "$env:windir\temp\NinjaOneAgentReinstall.log"
  $TimeStamp = Get-Date -Format 'yyyy-MM-dd HH:mm:ss'
  Add-Content -Path $LogPath -Value "$TimeStamp - $Message"
  Write-Host "$TimeStamp - $Message"
}

#Get current user context
$CurrentUser = New-Object Security.Principal.WindowsPrincipal $([Security.Principal.WindowsIdentity]::GetCurrent())
#Check user that is running the script is a member of Administrator Group
if (!($CurrentUser.IsInRole([Security.Principal.WindowsBuiltinRole]::Administrator))) {
  #UAC Prompt will occur for the user to input Administrator credentials and relaunch the powershell session
  Write-LogEntry 'This script must be ran with administrative privileges'
  Start-Process powershell.exe "-NoProfile -ExecutionPolicy Bypass -File `"$PSCommandPath`"" -Verb RunAs; Exit
}

<#--------------------------------------------------------------------------------------
  DERIVED VALUES — built automatically from what you entered above. No editing needed.
--------------------------------------------------------------------------------------#>

$InstallerName  = [System.IO.Path]::GetFileName($MsiUrl)
$MSIDestination = "$env:windir\temp\$InstallerName"
$UsingGenericInstaller = '$false'   # kept as a string — gets written into the reinstall script below as literal PowerShell code
$GenericInstallerUrlPattern = '/[0-9a-fA-F]{8}(-[0-9a-fA-F]{4}){3}-[0-9a-fA-F]{12}/'   # matches org-specific generated installer URLs

<#--------------------------------------------------------------------------------------
  VALIDATE INPUTS — checks the combination of MSI URL / Token / Host URL makes sense
  before downloading or changing anything on this device.
--------------------------------------------------------------------------------------#>

if ([string]::IsNullOrWhiteSpace($MsiUrl)) {
  Write-LogEntry 'No installer URL provided. Fill in $MsiUrl under USER CONFIGURATION at the top of this script. Cannot continue.'
  exit 1
}

if (!([string]::IsNullOrWhiteSpace($Token))) {
  # A Token was provided — validate it looks like a real GUID.
  if ($Token -notmatch '^[0-9a-fA-F]{8}-([0-9a-fA-F]{4}-){3}[0-9a-fA-F]{12}$') {
    Write-LogEntry 'An invalid token was provided. Please ensure it was entered correctly. Exiting.'
    exit 1
  }

  # A Token should only be used with a GENERIC installer URL, not an org-specific generated URL.
  if ($MsiUrl -match $GenericInstallerUrlPattern) {
    Write-LogEntry 'A token was provided, but the URL appears to be for a generated (org-specific) installer, not the generic installer.'
    Write-LogEntry 'Script will not continue. Please use either a generic installer URL with a Token, or an org-specific URL with no Token — not both.'
    exit 1
  }

  Write-LogEntry 'Generic installer being used and valid token detected. Continuing...'
  $UsingGenericInstaller = '$true'
}
else {
  # No Token was provided — the MSI URL must be an org-specific generated installer URL.
  if ($MsiUrl -notmatch $GenericInstallerUrlPattern) {
    Write-LogEntry 'A generic install URL was provided with no token. Please provide a token to use the generic installer. Exiting.'
    exit 1
  }

  if (![string]::IsNullOrWhiteSpace($FedRampHostUrl)) {
    Write-LogEntry 'A FedRAMP host instance was provided but no ClientUID (Token).'
    Write-LogEntry 'A FedRAMP migration requires a generic installer URL, a Host URL, AND a ClientUID entered in the Token variable. Exiting.'
    exit 1
  }
}

<#--------------------------------------------------------------------------------------
  DOWNLOAD THE INSTALLER
--------------------------------------------------------------------------------------#>

$FileDownload = New-FileDownload -Url $MsiUrl -Destination $MSIDestination

if (!($FileDownload)) {
  Write-LogEntry 'Failed to download file. Exiting.'
  Exit 1
}

Write-LogEntry 'Installer downloaded. Continuing to reinstallation...'

<#--------------------------------------------------------------------------------------
  UNINSTALL + REINSTALL — this code block is written out to its own script file and run
  in a separate PowerShell session, so this script can exit cleanly while the agent
  (which is currently running this script) gets removed and replaced.
--------------------------------------------------------------------------------------#>

$ReinstallCode = @'
Start-Sleep 30
function Write-LogEntry {
  param (
    [Parameter(Mandatory = $true)]
    [string]$Message
  )

  $LogPath = "$env:windir\temp\NinjaOneAgentReinstall.log"
  $TimeStamp = Get-Date -Format 'yyyy-MM-dd HH:mm:ss'
  Add-Content -Path $LogPath -Value "$TimeStamp - $Message"
  Write-Host "$TimeStamp - $Message"
}

function Install-MSI {
    Param (
      [Parameter(Mandatory = $true, ValueFromPipeline = $true)]
      [ValidateNotNullOrEmpty()]
      [System.IO.FileInfo]$File,
      [String[]]$AdditionalParams,
      [Switch]$OutputLog
    )
    $dateStamp = Get-Date -Format yyyyMMddTHHmmss
    $logFile = "$env:windir\temp\NinjaOneInstallLog_$DateStamp.log"
    $MSIArguments = @(
      "/i",
      ('"{0}"' -f $file.fullname),
      "/qn",
      "/norestart",
      "/L*v",
      $logFile
    )
    if ($AdditionalParams) {
      $MSIArguments += $AdditionalParams
    }
    Start-Process "msiexec.exe" -ArgumentList $MSIArguments -Wait -NoNewWindow 
    if ($OutputLog.IsPresent) {
      $logContents = get-content $logFile
      Write-Output $logContents
    }
  }
  
  function Get-NinjaInstallStatus {
    $CheckApp = Get-ItemProperty 'HKLM:\Software\Wow6432Node\Microsoft\Windows\CurrentVersion\Uninstall\*' | Where-Object { $_.DisplayName -eq 'NinjaRMMAgent' }
    if ($CheckApp) {
      return $true
    }
    else {
      return $false
    }
  }

function Uninstall-NinjaMSI {
  $Arguments = @(
    "/x$($UninstallString)"
    '/quiet'
    '/L*V'
    "$env:windir\temp\NinjaRMMAgent_uninstall.log"
    "WRAPPED_ARGUMENTS=`"--mode unattended`""
  )

  Start-Process "msiexec.exe" -ArgumentList $Arguments -Wait -NoNewWindow
  Write-LogEntry 'Finished running uninstaller. Continuing to clean up...'
  Start-Sleep 30
}

#Get current user context
$CurrentUser = New-Object Security.Principal.WindowsPrincipal $([Security.Principal.WindowsIdentity]::GetCurrent())
#Check user that is running the script is a member of Administrator Group
if (!($CurrentUser.IsInRole([Security.Principal.WindowsBuiltinRole]::Administrator))) {
  #UAC Prompt will occur for the user to input Administrator credentials and relaunch the powershell session
  Write-LogEntry 'This script must be ran with administrative privileges. Script will relaunch and request elevation...'
  Start-Process powershell.exe "-NoProfile -ExecutionPolicy Bypass -File `"$PSCommandPath`"" -Verb RunAs; Exit
}

$ErrorActionPreference = "SilentlyContinue"

Write-LogEntry 'Beginning NinjaRMM Agent removal...'
Write-LogEntry 'Path to log file: C:\Windows\Temp\NinjaRMMAgent_uninstall.log'

$NinjaRegPath = 'HKLM:\SOFTWARE\WOW6432Node\NinjaRMM LLC\NinjaRMMAgent'
$NinjaDataDirectory = "$($env:ProgramData)\NinjaRMMAgent"
$UninstallRegPath = 'HKLM:\SOFTWARE\WOW6432Node\Microsoft\Windows\CurrentVersion\Uninstall\*'
$NinjaModulePath = "$env:ProgramFiles\WindowsPowerShell\Modules\NJCliPSh"

if (!([System.Environment]::Is64BitOperatingSystem)) {
  $NinjaRegPath = 'HKLM:\SOFTWARE\NinjaRMM LLC\NinjaRMMAgent'
  $UninstallRegPath = 'HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\Uninstall\*'
}

$NinjaInstallLocation = (Get-ItemPropertyValue $NinjaRegPath -Name Location).Replace('/', '\')

if (!(Test-Path "$($NinjaInstallLocation)\NinjaRMMAgent.exe")) {
  $NinjaServicePath = ((Get-WMIObject Win32_Service | Where-Object { $_.Name -eq 'NinjaRMMAgent' }).PathName).Trim('"')
  if (!(Test-Path $NinjaServicePath)) {
    Write-LogEntry 'Unable to locate Ninja installation path. Continuing with cleanup...'
  }
  else {
    $NinjaInstallLocation = $NinjaServicePath | Split-Path
  }
}

Start-Process "$NinjaInstallLocation\NinjaRMMAgent.exe" -ArgumentList "-disableUninstallPrevention NOUI" -Wait -NoNewWindow
$UninstallString = (Get-ItemProperty $UninstallRegPath | Where-Object { ($_.DisplayName -eq 'NinjaRMMAgent') -and ($_.UninstallString -match 'msiexec') }).UninstallString

if (!($UninstallString)) {
  Write-LogEntry 'Unable to to determine uninstall string. Continuing with cleanup...' 
}
else {
  $UninstallString = $UninstallString.Split('X')[1]
  Uninstall-NinjaMSI
}

$NinjaServices = @('NinjaRMMAgent', 'nmsmanager', 'lockhart')
$Processes = @("NinjaRMMAgent", "NinjaRMMAgentPatcher", "njbar", "NinjaRMMProxyProcess64")

foreach ($Process in $Processes) {
  if ($GetP = Get-Process $Process) {
    try {
      Stop-Process $GetP -Force -ErrorAction Stop
      Write-LogEntry "Successfully stopped process: $($GetP.Name)"
    }
    catch {
      Write-LogEntry "Unable to stop $($GetP.Name) for the following reason:"
      Write-LogEntry "$($_.Exception.Message). Continuing..."
    }
  }
}

foreach ($NS in $NinjaServices) {
  if ($NS -eq 'lockhart' -and !(Test-Path "$NinjaInstallLocation\lockhart\bin\lockhart.exe")) {
    continue
  }
  if (Get-Service $NS) {
    try {
      Write-LogEntry "Stopping service $($NS)..."
      Stop-Service $NS -Force -ErrorAction Stop
    }
    catch {
      Write-LogEntry "Unable to stop $($NS) service..."
      Write-LogEntry "$($_.Exception.Message)"
      Write-LogEntry 'Attempting to remove service...'
    }
  
    & sc.exe DELETE $NS
    Start-Sleep 5
    if (Get-Service $NS) {
      Write-LogEntry "Failed to remove $($NS) service. Continuing with remaining removal steps..."
    }
    else {
      Write-LogEntry "Successfully removed $($NS) service."
    }
  }
}

if (Test-Path $NinjaInstallLocation) {
  Write-LogEntry 'Removing Ninja installation directory:'
  Write-LogEntry "$($NinjaInstallLocation)"
  try {
    Remove-Item $NinjaInstallLocation -Recurse -Force -ErrorAction Stop
    Write-LogEntry 'Successfully removed.'
  }
  catch {
    Write-LogEntry 'Failed to remove Ninja installation directory.'
    Write-LogEntry "$($_.Exception.Message)"
    Write-LogEntry 'Continuing with removal attempt...'
  }
}

if (Test-Path $NinjaDataDirectory) {
  Write-LogEntry 'Removing Ninja data directory:'
  Write-LogEntry "$($NinjaDataDirectory)"
  try {
    Remove-Item $NinjaDataDirectory -Recurse -Force -ErrorAction Stop
    Write-LogEntry 'Successfully removed.'
  }
  catch {
    Write-LogEntry 'Failed to remove Ninja data directory.'
    Write-LogEntry "$($_.Exception.Message)"
    Write-LogEntry 'Continuing with removal attempt...'
  }
}

if (Test-Path $NinjaModulePath) {
  Write-LogEntry 'Removing Ninja data directory:'
  Write-LogEntry "$($NinjaModulePath)"
  try {
    Remove-Item $NinjaModulePath -Recurse -Force -ErrorAction Stop
    Write-LogEntry 'Successfully removed.'
  }
  catch {
    Write-LogEntry 'Failed to remove Ninja PowerShell module directory.'
    Write-LogEntry "$($_.Exception.Message)"
    Write-LogEntry 'Continuing with removal attempt...'
  }
}

$MSIWrapperReg = 'HKLM:\SOFTWARE\WOW6432Node\EXEMSI.COM\MSI Wrapper\Installed'
$ProductInstallerReg = 'HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\Installer\UserData\S-1-5-18\Products'
$HKCRInstallerReg = 'Registry::\HKEY_CLASSES_ROOT\Installer\Products'

$RegKeysToRemove = [System.Collections.Generic.List[object]]::New()

(Get-ItemProperty $UninstallRegPath | Where-Object { $_.DisplayName -eq 'NinjaRMMAgent' }).PSPath | ForEach-Object { $RegKeysToRemove.Add($_) }
(Get-ItemProperty $ProductInstallerReg | Where-Object { $_.ProductName -eq 'NinjaRMMAgent' }).PSPath | ForEach-Object { $RegKeysToRemove.Add($_) }
(Get-ChildItem $MSIWrapperReg | Where-Object { $_.Name -match 'NinjaRMMAgent' }).PSPath | ForEach-Object { $RegKeysToRemove.Add($_) }
Get-ChildItem $HKCRInstallerReg | ForEach-Object { if ((Get-ItemPropertyValue $_.PSPath -Name 'ProductName') -eq 'NinjaRMMAgent') { $RegKeysToRemove.Add($_.PSPath) } }

$ProductInstallerKeys = Get-ChildItem $ProductInstallerReg | Select-Object *
foreach ($Key in $ProductInstallerKeys) {
  $KeyName = $($Key.Name).Replace('HKEY_LOCAL_MACHINE', 'HKLM:') + "\InstallProperties"
  if (Get-ItemProperty $KeyName | Where-Object { $_.DisplayName -eq 'NinjaRMMAgent' }) {
    $RegKeysToRemove.Add($Key.PSPath)
  }
}

Write-LogEntry 'Removing registry items if found...'
if (($RegKeysToRemove | Measure-Object).Count -gt 0 ) {
  foreach ($RegKey in $RegKeysToRemove) {
    if (!([string]::IsNullOrWhiteSpace($RegKey))) {
      Write-LogEntry "Attempting to remove: $($RegKey)"
      try { 
        Remove-Item $RegKey -Recurse -Force -ErrorAction Stop
        Write-LogEntry 'Successfully removed.'
      }
      catch {
        Write-LogEntry 'Failed to remove registry key.'
        Write-LogEntry "$($_.Exception.Message)"
        Write-LogEntry "Continuing with removal..."
      }
    }
  }
}

if (Test-Path $NinjaRegPath) {
  try {
    Write-LogEntry "Removing: $($NinjaRegPath)"
    Get-Item ($NinjaRegPath | Split-Path -ErrorAction Stop) | Remove-Item -Recurse -Force -ErrorAction Stop
    Write-LogEntry 'Successfully removed.'
  }
  catch {
    Write-LogEntry 'Failed to remove key.'
    Write-LogEntry "$($_.Exception.Message)"
    Write-LogEntry "Continuing with removal..."
  }
}

#Checks for rogue reg entry from older installations where ProductName was missing
#Filters out a Windows Common GUID that doesn't have a ProductName
$Child = Get-ChildItem 'HKLM:\Software\Classes\Installer\Products'
$MissingPNs = [System.Collections.Generic.List[object]]::New()

foreach ($C in $Child) {
  if ($C.Name -match '99E80CA9B0328e74791254777B1F42AE') {
    continue
  }
  try {
    Get-ItemPropertyValue $C.PSPath -Name 'ProductName' -ErrorAction Stop | Out-Null
  }
  catch {
    $MissingPNs.Add($($C.Name))
  } 
}

##Begin Ninja Remote Removal##
$NR = 'ncstreamer'

if (Get-Process $NR) {
  Write-LogEntry 'Stopping Ninja Remote process...'
  try {
    Get-Process $NR | Stop-Process -Force
  }
  catch {
    Write-LogEntry 'Unable to stop the Ninja Remote process...'
    Write-LogEntry "$($_.Exception.Message)"
    Write-LogEntry 'Continuing to Ninja Remote service...'
  }
}

if (Get-Service $NR) {
  try {
    Stop-Service $NR -Force
  }
  catch {
    Write-LogEntry 'Unable to stop the Ninja Remote service...'
    Write-LogEntry "$($_.Exception.Message)"
    Write-LogEntry 'Attempting to remove service...'
  }

  & sc.exe DELETE $NR
  Start-Sleep 5
  if (Get-Service $NR) {
    Write-LogEntry 'Failed to remove Ninja Remote service. Continuing with remaining removal steps...'
  }
}

$NRDriver = 'nrvirtualdisplay.inf'
$DriverCheck = pnputil /enum-drivers | Where-Object { $_ -match "$NRDriver" }
if ($DriverCheck) {
  Write-LogEntry 'Ninja Remote Virtual Driver found. Removing...'
  $DriverBreakdown = pnputil /enum-drivers | Where-Object { $_ -ne 'Microsoft PnP Utility' }

  $DriversArray = [System.Collections.Generic.List[object]]::New()
  $CurrentDriver = @{}
    
  foreach ($Line in $DriverBreakdown) {
    if ($Line -ne "") {
      $ObjectName = $Line.Split(':').Trim()[0]
      $ObjectValue = $Line.Split(':').Trim()[1]
      $CurrentDriver[$ObjectName] = $ObjectValue
    }
    else {
      if ($CurrentDriver.Count -gt 0) {
        $DriversArray.Add([PSCustomObject]$CurrentDriver)
        $CurrentDriver = @{}
      }
    }
  }

  $DriverToRemove = ($DriversArray | Where-Object { $_.'Provider Name' -eq 'NinjaOne' }).'Published Name'
  pnputil /delete-driver "$DriverToRemove" /force
}

$NRDirectory = "$($env:ProgramFiles)\NinjaRemote"
if (Test-Path $NRDirectory) {
  Write-LogEntry "Removing directory: $NRDirectory"
  Remove-Item $NRDirectory -Recurse -Force
  if (Test-Path $NRDirectory) {
    Write-LogEntry 'Failed to completely remove Ninja Remote directory at:'
    Write-LogEntry "$NRDirectory"
    Write-LogEntry 'Continuing to registry removal...'
  }
}

$NRHKUReg = 'Registry::\HKEY_USERS\S-1-5-18\Software\NinjaRMM LLC'
if (Test-Path $NRHKUReg) {
  Remove-Item $NRHKUReg -Recurse -Force
}

function Remove-NRRegistryItems {
  param (
    [Parameter(Mandatory = $true)]
    [string]$SID
  )
  $NRRunReg = "Registry::\HKEY_USERS\$SID\SOFTWARE\Microsoft\Windows\CurrentVersion\Run"
  $NRRegLocation = "Registry::\HKEY_USERS\$SID\Software\NinjaRMM LLC"
  if (Test-Path $NRRunReg) {
    $RunRegValues = Get-ItemProperty -Path $NRRunReg
    $PropertyNames = $RunRegValues.PSObject.Properties | Where-Object { $_.Name -match "NinjaRMM|NinjaOne" } 
    foreach ($PName in $PropertyNames) {    
      Write-LogEntry "Removing item..."
      Write-LogEntry "$($PName.Name): $($PName.Value)"
      Remove-ItemProperty $NRRunReg -Name $PName.Name -Force
    }
  }
  if (Test-Path $NRRegLocation) {
    Write-LogEntry "Removing $NRRegLocation..."
    Remove-Item $NRRegLocation -Recurse -Force
  }
  Write-LogEntry 'Registry removal completed.'
}

$AllProfiles = Get-CimInstance Win32_UserProfile | Select-Object LocalPath, SID, Loaded, Special | 
Where-Object { $_.SID -like "S-1-5-21-*" }
$Mounted = $AllProfiles | Where-Object { $_.Loaded -eq $true }
$Unmounted = $AllProfiles | Where-Object { $_.Loaded -eq $false }

$Mounted | Foreach-Object {
  Write-LogEntry "Removing registry items for $($_.LocalPath)"
  Remove-NRRegistryItems -SID "$($_.SID)"
}

$Unmounted | ForEach-Object {
  $Hive = "$($_.LocalPath)\NTUSER.DAT"
  if (Test-Path $Hive) {      
    Write-LogEntry "Loading hive and removing Ninja Remote registry items for $($_.LocalPath)..."

    REG LOAD HKU\$($_.SID) $Hive 2>&1>$null

    Remove-NRRegistryItems -SID "$($_.SID)"
        
    [GC]::Collect()
    [GC]::WaitForPendingFinalizers()
          
    REG UNLOAD HKU\$($_.SID) 2>&1>$null
  } 
}

$NRPrinter = Get-Printer | Where-Object { $_.Name -eq 'NinjaRemote' }

if ($NRPrinter) {
  Write-LogEntry 'Removing Ninja Remote printer...'
  Remove-Printer -InputObject $NRPrinter
}

$NRPrintDriverPath = "$env:SystemDrive\Users\Public\Documents\NrSpool\NrPdfPrint"
if (Test-Path $NRPrintDriverPath) {
  Write-LogEntry 'Removing Ninja Remote printer driver...'
  Remove-Item $NRPrintDriverPath -Force
}

Write-LogEntry 'Removal of Ninja Remote complete.'
##End Ninja Remote Removal##

if ($MissingPNs) {
  Write-LogEntry '############################# !!! WARNING !!! ####################################'
  Write-LogEntry 'Some registry keys are missing the Product Name.'
  Write-LogEntry 'This could be an indicator of a corrupt Ninja install key.'
  Write-LogEntry 'If you are still unable to install the NinjaOne Agent after running this script...'
  Write-LogEntry 'Please make a backup of the following keys and then remove them from the registry:'
  Write-LogEntry ( $MissingPNs | Out-String )
  Write-LogEntry '##################################################################################'
}

Write-LogEntry 'Removal script completed. Please review if any errors displayed.'

###Reinstall Section###

Start-Sleep 30

Write-LogEntry 'Confirming NinjaOne was fully removed...'
  
if (Get-NinjaInstallStatus) {
  Write-LogEntry 'Cannot continue as NinjaOne is already installed.'
  exit 0
}
  
Write-LogEntry 'NinjaOne not found. Continuing with install script...'

if ($UsingGeneric) {
  if ($HostURL) {
    Install-MSI -File $MSIDestination -AdditionalParams "CLIENTUID=$Token HOST=$HostURL"
  }
  else {
    Install-MSI -File $MSIDestination -AdditionalParams TOKENID=$Token
  }
}
else {
  Install-MSI -File $MSIDestination
}

if (!(Get-NinjaInstallStatus)) {
  Write-LogEntry 'Failed to install NinjaOne agent.'
  Remove-Item $MSIDestination -Force
  exit 1
}
  
Write-LogEntry 'Successfully installed NinjaOne agent.'
Remove-Item $MSIDestination -Force
exit 0
'@

<#--------------------------------------------------------------------------------------
  HAND OFF TO THE SEPARATE PROCESS — writes the block above to its own .ps1 file, along
  with the specific values gathered from this run, then launches it detached from this
  script so it can keep running after the agent (and this script's own process) stops.
--------------------------------------------------------------------------------------#>

$VariablesToPass = @"
`$InstallerName = '$InstallerName'
`$MSIDestination = '$MSIDestination'
`$Token = '$Token'
`$HostURL = '$FedRampHostUrl'
`$UsingGeneric = $UsingGenericInstaller`n
"@

$ReinstallCodePSDestination = "$env:windir\temp\NinjaOneAgentReinstall.ps1"

try {
  New-Item "$ReinstallCodePSDestination" -Force
  Set-Content "$ReinstallCodePSDestination" ($VariablesToPass + $ReinstallCode)
}
catch {
  Write-LogEntry "$($_.Exception.Message)"
  Write-LogEntry 'Failed to create or update removal PS1 file. Exiting.'
  exit 1
}

$GrabProcess = Start-Process powershell.exe -ArgumentList '-NoProfile', '-ExecutionPolicy', 'Bypass', '-File', "$ReinstallCodePSDestination" -WindowStyle Hidden -PassThru
Write-LogEntry "The PowerShell uninstall process ID is: $($GrabProcess.ID)"
Write-LogEntry 'The uninstall process will begin in a moment. The device will go offline in the current Ninja instance shortly.'
Write-LogEntry 'You can check C:\Windows\temp\NinjaOneAgentReinstall.log for additional reinstall details.'
exit 0
