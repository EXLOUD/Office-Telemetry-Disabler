# Office Privacy and Telemetry Disabler
# This script disables telemetry, logging, feedback, connected experiences, and other privacy-related settings in Microsoft Office
# Supports Office 2010, 2013, 2016/2019/365, 2021, 2024, including Click-to-Run installations
# Compatible with Windows 7 and newer
# by EXLOUD
# >> https://github.com/EXLOUD <<
# aka Vladyslav Bober

# Color scheme for consistent output
$Colors = @{
    Title    = 'Cyan'
    Section  = 'Yellow'
    Success  = 'Green'
    Info     = 'Blue'
    Warning  = 'Yellow'
    Error    = 'Red'
    Gray     = 'Gray'
    Found    = 'Green'
    Changed  = 'Magenta'
    NotFound = 'Gray'
    OK       = 'Green'
    Skip     = 'DarkGray'
}

function Get-WindowsVersion {
    $version = [Environment]::OSVersion.Version
    $major = $version.Major
    $minor = $version.Minor
    
    if ($major -eq 6 -and $minor -eq 1) { return "Windows 7" }
    elseif ($major -eq 6 -and $minor -eq 2) { return "Windows 8" }
    elseif ($major -eq 6 -and $minor -eq 3) { return "Windows 8.1" }
    elseif ($major -eq 10) { return "Windows 10/11" }
    else { return "Windows $major.$minor" }
}

Write-Host "--- Checking for admin rights ---" -ForegroundColor $Colors.Section

$currentPrincipal = $null
$isAdmin = $false

try {
    $currentPrincipal = New-Object Security.Principal.WindowsPrincipal([Security.Principal.WindowsIdentity]::GetCurrent())
    $isAdmin = $currentPrincipal.IsInRole([Security.Principal.WindowsBuiltInRole]::Administrator)
} catch {
    try {
        $identity = [Security.Principal.WindowsIdentity]::GetCurrent()
        $principal = New-Object Security.Principal.WindowsPrincipal($identity)
        $isAdmin = $principal.IsInRole("Administrators")
    } catch {
        Write-Host "  [!] Cannot determine admin rights, proceeding with caution" -ForegroundColor $Colors.Warning
        $isAdmin = $false
    }
}

if (-not $isAdmin) {
    Write-Host "  [X] Administrator privileges required." -ForegroundColor $Colors.Error
    Write-Host "    Please run the script as Administrator to use this script." -ForegroundColor $Colors.Warning
    Write-Host ""
    Read-Host "Press Enter to exit"
    exit 1
} else {
    Write-Host "  [OK] Running with administrator privileges." -ForegroundColor $Colors.Success
}

$windowsVersion = Get-WindowsVersion
Write-Host "  [INFO] Detected: $windowsVersion" -ForegroundColor $Colors.Info

# Define supported Office versions
$OfficeVersions = @{
    "14.0" = "Office 2010"
    "15.0" = "Office 2013"
    "16.0" = "Office 2016/365/2019/2021/2024"
}

# Function to set registry value with logging
function Set-RegistryValueWithLogging {
    param (
        [string]$Path,
        [string]$Name,
        [string]$Type,
        [Parameter(Mandatory=$true)]
        $Value,
        [string]$Description
    )
    
    try {
        if (-not (Test-Path $Path)) {
            New-Item -Path $Path -Force | Out-Null
        }
        
        $currentValue = Get-ItemProperty -Path $Path -Name $Name -ErrorAction SilentlyContinue
        if ($currentValue -and $currentValue.$Name -eq $Value) {
            Write-Host "  [SKIP] Already set: $Name -> $Value" -ForegroundColor $Colors.Skip
            Write-Host "    Description: $Description" -ForegroundColor $Colors.Skip
            return
        }
        
        Set-ItemProperty -Path $Path -Name $Name -Type $Type -Value $Value -Force
        Write-Host "  [OK] Found and changed: $Name -> $Value" -ForegroundColor $Colors.OK
        Write-Host "    Description: $Description" -ForegroundColor $Colors.OK
    }
    catch {
        Write-Host "  [ERROR] Failed to set: $Name -> $Value" -ForegroundColor $Colors.Error
        Write-Host "    Error: $_" -ForegroundColor $Colors.Error
    }
}

# Function to disable scheduled task with logging (compatible with Windows 7)
function Disable-ScheduledTaskWithLogging {
    param (
        [string]$TaskName,
        [string]$Description = ""
    )
    
    try {
        if (Get-Command "Get-ScheduledTask" -ErrorAction SilentlyContinue) {
            $task = Get-ScheduledTask -TaskName $TaskName -ErrorAction SilentlyContinue
            if ($task) {
                $taskState = $task.State
                if ($taskState -eq 'Disabled') {
                    Write-Host "  [OK] Task already disabled: $TaskName" -ForegroundColor $Colors.Success
                    $script:telemetryTasksDisabled++
                } else {
                    Disable-ScheduledTask -TaskName $TaskName -ErrorAction Stop | Out-Null
                    Write-Host "  [OK] Found and disabled: $TaskName" -ForegroundColor $Colors.Changed
                    $script:telemetryTasksDisabled++
                }
                
                if ($Description) {
                    Write-Host "    Description: $Description" -ForegroundColor $Colors.Info
                }
            } else {
                Write-Host "  [SKIP] Task not found: $TaskName" -ForegroundColor $Colors.NotFound
                $script:telemetryTasksNotFound++
            }
        } else {
            $taskPath = $TaskName -replace "Microsoft\\", "\"
            $queryResult = & schtasks.exe /Query /TN $taskPath 2>$null
            if ($LASTEXITCODE -eq 0) {
                # Використовуємо $queryResult для отримання інформації про завдання
                $taskInfo = $queryResult | ConvertFrom-Csv -ErrorAction SilentlyContinue
                if ($taskInfo -and $taskInfo.Status -eq "Disabled") {
                    Write-Host "  [OK] Task already disabled: $TaskName" -ForegroundColor $Colors.Success
                    $script:telemetryTasksDisabled++
                } else {
                    & schtasks.exe /Change /TN $taskPath /DISABLE >$null 2>&1
                    if ($LASTEXITCODE -eq 0) {
                        Write-Host "  [OK] Found and disabled: $TaskName" -ForegroundColor $Colors.Changed
                        $script:telemetryTasksDisabled++
                    } else {
                        Write-Host "  [ERROR] Failed to disable: $TaskName" -ForegroundColor $Colors.Error
                    }
                }
                
                if ($Description) {
                    Write-Host "    Description: $Description" -ForegroundColor $Colors.Info
                }
            } else {
                Write-Host "  [SKIP] Task not found: $TaskName" -ForegroundColor $Colors.NotFound
                $script:telemetryTasksNotFound++
            }
        }
    } catch {
        Write-Host "  [ERROR] Error disabling $TaskName : $($_.Exception.Message)" -ForegroundColor $Colors.Error
    }
}

# Detect installed Office versions
$installedVersions = @()
if (Test-Path "HKCU:\SOFTWARE\Microsoft\Office") {
    $hkcuVersions = Get-ChildItem -Path "HKCU:\SOFTWARE\Microsoft\Office" -ErrorAction SilentlyContinue | 
        Where-Object { $_.PSChildName -match "^\d+\.\d+$" -and $OfficeVersions.ContainsKey($_.PSChildName) } | 
        Select-Object -ExpandProperty PSChildName
    if ($hkcuVersions) { 
        $installedVersions += $hkcuVersions 
        Write-Host "  [INFO] Found in HKCU: $($hkcuVersions -join ', ')" -ForegroundColor $Colors.Info
    }
}
if (Test-Path "HKLM:\SOFTWARE\Microsoft\Office") {
    $hklmVersions = Get-ChildItem -Path "HKLM:\SOFTWARE\Microsoft\Office" -ErrorAction SilentlyContinue | 
        Where-Object { $_.PSChildName -match "^\d+\.\d+$" -and $OfficeVersions.ContainsKey($_.PSChildName) } | 
        Select-Object -ExpandProperty PSChildName
    if ($hklmVersions) { 
        $installedVersions += $hklmVersions 
        Write-Host "  [INFO] Found in HKLM: $($hklmVersions -join ', ')" -ForegroundColor $Colors.Info
    }
}
if (Test-Path "HKLM:\SOFTWARE\Microsoft\Office\ClickToRun") {
    $ctrVersions = Get-ItemProperty -Path "HKLM:\SOFTWARE\Microsoft\Office\ClickToRun\Configuration" -ErrorAction SilentlyContinue | 
        Select-Object -ExpandProperty VersionToReport
    if ($ctrVersions -and $ctrVersions -match "^\d+\.\d+\.\d+\.\d+$" -and "16.0" -notin $installedVersions) {
        $installedVersions += "16.0"
        Write-Host "  [INFO] Found Click-to-Run: 16.0" -ForegroundColor $Colors.Info
    }
}
$installedVersions = $installedVersions | Sort-Object -Unique
Write-Host "  [INFO] Total unique versions: $($installedVersions -join ', ')" -ForegroundColor $Colors.Info

if (-not $installedVersions) {
    Write-Host "No supported Office versions found in registry. Please ensure Microsoft Office is installed." -ForegroundColor $Colors.Error
    exit
}

$modernVersions = $installedVersions | Where-Object { $_ -ge "16.0" }
$updateVersions = $installedVersions | Where-Object { $_ -ge "15.0" }

# Initialize counters for tasks
$script:telemetryTasksProcessed = 0
$script:telemetryTasksDisabled = 0
$script:telemetryTasksNotFound = 0

# Disable client telemetry
Write-Host "`n--- Disabling client telemetry ---" -ForegroundColor $Colors.Section

Set-RegistryValueWithLogging -Path "HKCU:\SOFTWARE\Microsoft\Office\Common\ClientTelemetry" -Name "DisableTelemetry" -Type "DWord" -Value 1 -Description "Disable common telemetry"
Set-RegistryValueWithLogging -Path "HKCU:\SOFTWARE\Microsoft\Office\Common\ClientTelemetry" -Name "VerboseLogging" -Type "DWord" -Value 0 -Description "Disable verbose logging"
Set-RegistryValueWithLogging -Path "HKCU:\SOFTWARE\Microsoft\Office\Common\ClientTelemetry" -Name "SendTelemetry" -Type "DWord" -Value 3 -Description "Set telemetry level to minimum"
Set-RegistryValueWithLogging -Path "HKCU:\SOFTWARE\Policies\Microsoft\Office\Common\ClientTelemetry" -Name "SendTelemetry" -Type "DWord" -Value 3 -Description "Set telemetry level to minimum (Policies)"

if (Test-Path "HKLM:\SOFTWARE\Microsoft\Office\ClickToRun\Configuration") {
    Set-RegistryValueWithLogging -Path "HKLM:\SOFTWARE\Microsoft\Office\ClickToRun\Configuration" -Name "SendTelemetry" -Type "DWord" -Value 3 -Description "Set telemetry level to minimum (ClickToRun)"
}

foreach ($version in $installedVersions) {
    Write-Host "`nProcessing $($OfficeVersions[$version]) telemetry..." -ForegroundColor $Colors.Info
    
    Set-RegistryValueWithLogging -Path "HKCU:\SOFTWARE\Microsoft\Office\$version\Common\ClientTelemetry" -Name "DisableTelemetry" -Type "DWord" -Value 1 -Description "Disable telemetry"
    Set-RegistryValueWithLogging -Path "HKCU:\SOFTWARE\Microsoft\Office\$version\Common\ClientTelemetry" -Name "VerboseLogging" -Type "DWord" -Value 0 -Description "Disable verbose logging"
}

# Disable Customer Experience Improvement Program
Write-Host "`n--- Disabling Customer Experience Improvement Program ---" -ForegroundColor $Colors.Section

foreach ($version in $installedVersions) {
    Write-Host "`nProcessing $($OfficeVersions[$version]) CEIP..." -ForegroundColor $Colors.Info
    
    Set-RegistryValueWithLogging -Path "HKCU:\SOFTWARE\Microsoft\Office\$version\Common" -Name "QMEnable" -Type "DWord" -Value 0 -Description "Disable CEIP"
    Set-RegistryValueWithLogging -Path "HKCU:\SOFTWARE\Microsoft\Office\$version\Common" -Name "sendcustomerdata" -Type "DWord" -Value 0 -Description "Disable customer data collection"
    Set-RegistryValueWithLogging -Path "HKCU:\SOFTWARE\Microsoft\Office\$version\Common" -Name "updatereliabilitydata" -Type "DWord" -Value 0 -Description "Disable update reliability data"
    Set-RegistryValueWithLogging -Path "HKCU:\SOFTWARE\Microsoft\Office\$version\Common" -Name "linkedin" -Type "DWord" -Value 0 -Description "Disable LinkedIn integration"
    Set-RegistryValueWithLogging -Path "HKCU:\SOFTWARE\Policies\Microsoft\Office\$version\Common" -Name "sendcustomerdata" -Type "DWord" -Value 0 -Description "Disable customer data collection (Policies)"
    Set-RegistryValueWithLogging -Path "HKCU:\SOFTWARE\Policies\Microsoft\Office\$version\Common" -Name "updatereliabilitydata" -Type "DWord" -Value 0 -Description "Disable update reliability data (Policies)"
    Set-RegistryValueWithLogging -Path "HKCU:\SOFTWARE\Policies\Microsoft\Office\$version\Common" -Name "linkedin" -Type "DWord" -Value 0 -Description "Disable LinkedIn integration (Policies)"
}

# Disable feedback
Write-Host "`n--- Disabling feedback ---" -ForegroundColor $Colors.Section

foreach ($version in $installedVersions) {
    Write-Host "`nProcessing $($OfficeVersions[$version]) feedback..." -ForegroundColor $Colors.Info
    
    Set-RegistryValueWithLogging -Path "HKCU:\SOFTWARE\Microsoft\Office\$version\Common\Feedback" -Name "Enabled" -Type "DWord" -Value 0 -Description "Disable feedback"
    Set-RegistryValueWithLogging -Path "HKCU:\SOFTWARE\Microsoft\Office\$version\Common\Feedback" -Name "includescreenshot" -Type "DWord" -Value 0 -Description "Disable screenshot in feedback"
    Set-RegistryValueWithLogging -Path "HKCU:\SOFTWARE\Microsoft\Office\$version\Common\Feedback" -Name "includeemail" -Type "DWord" -Value 0 -Description "Disable email in feedback"
    Set-RegistryValueWithLogging -Path "HKCU:\SOFTWARE\Microsoft\Office\$version\Common\Feedback" -Name "surveyenabled" -Type "DWord" -Value 0 -Description "Disable surveys"
    Set-RegistryValueWithLogging -Path "HKCU:\SOFTWARE\Policies\Microsoft\Office\$version\Common\Feedback" -Name "includescreenshot" -Type "DWord" -Value 0 -Description "Disable screenshot in feedback (Policies)"
    Set-RegistryValueWithLogging -Path "HKCU:\SOFTWARE\Policies\Microsoft\Office\$version\Common\Feedback" -Name "includeemail" -Type "DWord" -Value 0 -Description "Disable email in feedback (Policies)"
    Set-RegistryValueWithLogging -Path "HKCU:\SOFTWARE\Policies\Microsoft\Office\$version\Common\Feedback" -Name "surveyenabled" -Type "DWord" -Value 0 -Description "Disable surveys (Policies)"
}

# Disable Connected Experiences
Write-Host "`n--- Disabling Connected Experiences ---" -ForegroundColor $Colors.Section

if (Test-Path "HKLM:\SOFTWARE\Microsoft\Office\ClickToRun\Configuration") {
    Set-RegistryValueWithLogging -Path "HKLM:\SOFTWARE\Microsoft\Office\ClickToRun\Configuration" -Name "UserContentDisabled" -Type "DWord" -Value 2 -Description "Disable content analysis (ClickToRun)"
    Set-RegistryValueWithLogging -Path "HKLM:\SOFTWARE\Microsoft\Office\ClickToRun\Configuration" -Name "DownloadContentDisabled" -Type "DWord" -Value 2 -Description "Disable online content download (ClickToRun)"
    Set-RegistryValueWithLogging -Path "HKLM:\SOFTWARE\Microsoft\Office\ClickToRun\Configuration" -Name "DisconnectedState" -Type "DWord" -Value 2 -Description "Set disconnected state (ClickToRun)"
    Set-RegistryValueWithLogging -Path "HKLM:\SOFTWARE\Microsoft\Office\ClickToRun\Configuration" -Name "ControllerConnectedServicesEnabled" -Type "DWord" -Value 2 -Description "Disable controller connected services (ClickToRun)"
    Set-RegistryValueWithLogging -Path "HKLM:\SOFTWARE\Microsoft\Office\ClickToRun\Configuration" -Name "skydrivesigninoption" -Type "DWord" -Value 0 -Description "Disable SkyDrive/OneDrive sign-in (ClickToRun)"
}

foreach ($version in $modernVersions) {
    Write-Host "`nProcessing $($OfficeVersions[$version]) connected experiences..." -ForegroundColor $Colors.Info
    
    Set-RegistryValueWithLogging -Path "HKCU:\SOFTWARE\Microsoft\Office\$version\Common\Privacy" -Name "UserContentDisabled" -Type "DWord" -Value 2 -Description "Disable content analysis"
    Set-RegistryValueWithLogging -Path "HKCU:\SOFTWARE\Microsoft\Office\$version\Common\Privacy" -Name "DownloadContentDisabled" -Type "DWord" -Value 2 -Description "Disable online content download"
    Set-RegistryValueWithLogging -Path "HKCU:\SOFTWARE\Microsoft\Office\$version\Common\Privacy" -Name "disconnectedstate" -Type "DWord" -Value 2 -Description "Set disconnected state"
    Set-RegistryValueWithLogging -Path "HKCU:\SOFTWARE\Microsoft\Office\$version\Common\Privacy" -Name "controllerconnectedservicesenabled" -Type "DWord" -Value 2 -Description "Disable controller connected services"
    Set-RegistryValueWithLogging -Path "HKCU:\SOFTWARE\Policies\Microsoft\Office\$version\Common\Privacy" -Name "disconnectedstate" -Type "DWord" -Value 2 -Description "Set disconnected state (Policies)"
    Set-RegistryValueWithLogging -Path "HKCU:\SOFTWARE\Policies\Microsoft\Office\$version\Common\Privacy" -Name "controllerconnectedservicesenabled" -Type "DWord" -Value 2 -Description "Disable controller connected services (Policies)"
    Set-RegistryValueWithLogging -Path "HKCU:\SOFTWARE\Microsoft\Office\$version\Common\Internet" -Name "UseOnlineContent" -Type "DWord" -Value 0 -Description "Disable online content"
    Set-RegistryValueWithLogging -Path "HKCU:\SOFTWARE\Microsoft\Office\$version\Common\Internet" -Name "serviceleveloptions" -Type "DWord" -Value 0 -Description "Disable service level options"
    Set-RegistryValueWithLogging -Path "HKCU:\SOFTWARE\Policies\Microsoft\Office\$version\Common\Internet" -Name "serviceleveloptions" -Type "DWord" -Value 0 -Description "Disable service level options (Policies)"
}

# Disable Office updates and notifications
Write-Host "`n--- Disabling Office updates and notifications ---" -ForegroundColor $Colors.Section

if (Test-Path "HKLM:\SOFTWARE\Microsoft\Office\ClickToRun\Configuration") {
    Set-RegistryValueWithLogging -Path "HKLM:\SOFTWARE\Microsoft\Office\ClickToRun\Configuration" -Name "UpdatesEnabled" -Type "String" -Value "False" -Description "Disable automatic updates (ClickToRun)"
}

foreach ($version in $updateVersions) {
    Write-Host "`nProcessing $($OfficeVersions[$version]) updates and notifications..." -ForegroundColor $Colors.Info
    
    Set-RegistryValueWithLogging -Path "HKCU:\SOFTWARE\Microsoft\Office\$version\Update" -Name "OfficeMgmtCOM" -Type "DWord" -Value 0 -Description "Disable Office management COM"
    Set-RegistryValueWithLogging -Path "HKCU:\SOFTWARE\Microsoft\Office\$version\Update" -Name "EnableAutomaticUpdates" -Type "DWord" -Value 0 -Description "Disable automatic updates"
    Set-RegistryValueWithLogging -Path "HKCU:\SOFTWARE\Microsoft\Office\$version\Common\PTWatson" -Name "PTWOptIn" -Type "DWord" -Value 0 -Description "Disable PTWatson opt-in"
}

# Disable Microsoft Office logging
Write-Host "`n--- Disabling Microsoft Office logging ---" -ForegroundColor $Colors.Section

foreach ($version in $installedVersions) {
    Write-Host "`nProcessing $($OfficeVersions[$version]) logging..." -ForegroundColor $Colors.Info
    
    Set-RegistryValueWithLogging -Path "HKCU:\SOFTWARE\Microsoft\Office\$version\Outlook\Options\Mail" -Name "EnableLogging" -Type "DWord" -Value 0 -Description "Disable Outlook mail logging"
    Set-RegistryValueWithLogging -Path "HKCU:\SOFTWARE\Microsoft\Office\$version\Word\Options" -Name "EnableLogging" -Type "DWord" -Value 0 -Description "Disable Word logging"
    Set-RegistryValueWithLogging -Path "HKCU:\SOFTWARE\Microsoft\Office\$version\OSM" -Name "EnableLogging" -Type "DWord" -Value 0 -Description "Disable OSM logging"
    Set-RegistryValueWithLogging -Path "HKCU:\SOFTWARE\Microsoft\Office\$version\OSM" -Name "EnableUpload" -Type "DWord" -Value 0 -Description "Disable OSM upload"
    Set-RegistryValueWithLogging -Path "HKCU:\SOFTWARE\Policies\Microsoft\Office\$version\OSM" -Name "EnableLogging" -Type "DWord" -Value 0 -Description "Disable OSM logging (Policies)"
    Set-RegistryValueWithLogging -Path "HKCU:\SOFTWARE\Policies\Microsoft\Office\$version\OSM" -Name "EnableUpload" -Type "DWord" -Value 0 -Description "Disable OSM upload (Policies)"
}

# Disable First Run Settings
Write-Host "`n--- Disabling First Run Settings ---" -ForegroundColor $Colors.Section

foreach ($version in $installedVersions) {
    Write-Host "`nProcessing $($OfficeVersions[$version]) First Run settings..." -ForegroundColor $Colors.Info
    
    Set-RegistryValueWithLogging -Path "HKCU:\SOFTWARE\Microsoft\Office\$version\Firstrun" -Name "BootedRTM" -Type "DWord" -Value 1 -Description "Mark Office as booted"
    Set-RegistryValueWithLogging -Path "HKCU:\SOFTWARE\Microsoft\Office\$version\Firstrun" -Name "disablemovie" -Type "DWord" -Value 1 -Description "Disable first-run movie"
    Set-RegistryValueWithLogging -Path "HKCU:\SOFTWARE\Policies\Microsoft\Office\$version\Firstrun" -Name "BootedRTM" -Type "DWord" -Value 1 -Description "Mark Office as booted (Policies)"
    Set-RegistryValueWithLogging -Path "HKCU:\SOFTWARE\Policies\Microsoft\Office\$version\Firstrun" -Name "disablemovie" -Type "DWord" -Value 1 -Description "Disable first-run movie (Policies)"
}

# Disable Lync Telemetry
Write-Host "`n--- Disabling Lync Telemetry ---" -ForegroundColor $Colors.Section

foreach ($version in $installedVersions) {
    Write-Host "`nProcessing $($OfficeVersions[$version]) Lync telemetry..." -ForegroundColor $Colors.Info
    
    Set-RegistryValueWithLogging -Path "HKCU:\SOFTWARE\Microsoft\Office\$version\Lync" -Name "disableautomaticsendtracing" -Type "DWord" -Value 1 -Description "Disable Lync automatic tracing"
    Set-RegistryValueWithLogging -Path "HKCU:\SOFTWARE\Policies\Microsoft\Office\$version\Lync" -Name "disableautomaticsendtracing" -Type "DWord" -Value 1 -Description "Disable Lync automatic tracing (Policies)"
}

# Disable Security File Validation Reporting
Write-Host "`n--- Disabling Security File Validation Reporting ---" -ForegroundColor $Colors.Section

foreach ($version in $installedVersions) {
    Write-Host "`nProcessing $($OfficeVersions[$version]) file validation reporting..." -ForegroundColor $Colors.Info
    
    Set-RegistryValueWithLogging -Path "HKCU:\SOFTWARE\Microsoft\Office\$version\Common\Security\FileValidation" -Name "disablereporting" -Type "DWord" -Value 1 -Description "Disable file validation reporting"
    Set-RegistryValueWithLogging -Path "HKCU:\SOFTWARE\Policies\Microsoft\Office\$version\Common\Security\FileValidation" -Name "disablereporting" -Type "DWord" -Value 1 -Description "Disable file validation reporting (Policies)"
}

# Disable OSM Prevented Applications and Solution Types
Write-Host "`n--- Disabling OSM Prevented Applications and Solution Types ---" -ForegroundColor $Colors.Section

foreach ($version in $installedVersions) {
    Write-Host "`nProcessing $($OfficeVersions[$version]) OSM prevented settings..." -ForegroundColor $Colors.Info
    
    $preventedApps = @("accesssolution", "olksolution", "onenotesolution", "pptsolution", "projectsolution", "publishersolution", "visiosolution", "wdsolution", "xlsolution")
    $preventedTypes = @("agave", "appaddins", "comaddins", "documentfiles", "templatefiles")
    
    foreach ($app in $preventedApps) {
        Set-RegistryValueWithLogging -Path "HKCU:\SOFTWARE\Microsoft\Office\$version\OSM\PreventedApplications" -Name $app -Type "DWord" -Value 1 -Description "Disable $app telemetry"
        Set-RegistryValueWithLogging -Path "HKCU:\SOFTWARE\Policies\Microsoft\Office\$version\OSM\PreventedApplications" -Name $app -Type "DWord" -Value 1 -Description "Disable $app telemetry (Policies)"
    }
    foreach ($type in $preventedTypes) {
        Set-RegistryValueWithLogging -Path "HKCU:\SOFTWARE\Microsoft\Office\$version\OSM\PreventedSolutiontypes" -Name $type -Type "DWord" -Value 1 -Description "Disable $type telemetry"
        Set-RegistryValueWithLogging -Path "HKCU:\SOFTWARE\Policies\Microsoft\Office\$version\OSM\PreventedSolutiontypes" -Name $type -Type "DWord" -Value 1 -Description "Disable $type telemetry (Policies)"
    }
}

# Disable General Settings
Write-Host "`n--- Disabling General Settings ---" -ForegroundColor $Colors.Section

foreach ($version in $installedVersions) {
    Write-Host "`nProcessing $($OfficeVersions[$version]) General settings..." -ForegroundColor $Colors.Info
    
    Set-RegistryValueWithLogging -Path "HKCU:\SOFTWARE\Microsoft\Office\$version\Common\General" -Name "disablecomingsoon" -Type "DWord" -Value 1 -Description "Disable coming soon prompts"
    Set-RegistryValueWithLogging -Path "HKCU:\SOFTWARE\Microsoft\Office\$version\Common\General" -Name "optindisable" -Type "DWord" -Value 1 -Description "Disable opt-in prompts"
    Set-RegistryValueWithLogging -Path "HKCU:\SOFTWARE\Microsoft\Office\$version\Common\General" -Name "shownfirstrunoptin" -Type "DWord" -Value 1 -Description "Mark first-run opt-in as shown"
    Set-RegistryValueWithLogging -Path "HKCU:\SOFTWARE\Microsoft\Office\$version\Common\General" -Name "ShownFileFmtPrompt" -Type "DWord" -Value 1 -Description "Mark file format prompt as shown"
    Set-RegistryValueWithLogging -Path "HKCU:\SOFTWARE\Policies\Microsoft\Office\$version\Common\General" -Name "disablecomingsoon" -Type "DWord" -Value 1 -Description "Disable coming soon prompts (Policies)"
    Set-RegistryValueWithLogging -Path "HKCU:\SOFTWARE\Policies\Microsoft\Office\$version\Common\General" -Name "optindisable" -Type "DWord" -Value 1 -Description "Disable opt-in prompts (Policies)"
    Set-RegistryValueWithLogging -Path "HKCU:\SOFTWARE\Policies\Microsoft\Office\$version\Common\General" -Name "shownfirstrunoptin" -Type "DWord" -Value 1 -Description "Mark first-run opt-in as shown (Policies)"
    Set-RegistryValueWithLogging -Path "HKCU:\SOFTWARE\Policies\Microsoft\Office\$version\Common\General" -Name "ShownFileFmtPrompt" -Type "DWord" -Value 1 -Description "Mark file format prompt as shown (Policies)"
}

# Disable Office telemetry scheduled tasks
Write-Host "`n--- Disabling Office telemetry scheduled tasks ---" -ForegroundColor $Colors.Section

$telemetryTasks = @(
    @{ Name = "Microsoft\Office\OfficeTelemetryAgentFallBack"; Description = "Office 2013 Telemetry Agent Fallback" },
    @{ Name = "Microsoft\Office\OfficeTelemetryAgentLogOn"; Description = "Office 2013 Telemetry Agent Logon" },
    @{ Name = "Microsoft\Office\OfficeTelemetryAgentFallBack2016"; Description = "Office 2016/2019 Telemetry Agent Fallback" },
    @{ Name = "Microsoft\Office\OfficeTelemetryAgentLogOn2016"; Description = "Office 2016/2019 Telemetry Agent Logon" },
    @{ Name = "Microsoft\Office\OfficeBackgroundTaskHandlerRegistration"; Description = "Office Background Task Handler Registration" },
    @{ Name = "Microsoft\Office\OfficeBackgroundTaskHandlerLogon"; Description = "Office Background Task Handler Logon" },
    @{ Name = "Microsoft\Office\Office Automatic Updates"; Description = "Office Automatic Updates" },
    @{ Name = "Microsoft\Office\Office Automatic Updates 2.0"; Description = "Office Automatic Updates 2.0" },
    @{ Name = "Microsoft\Office\Office Feature Updates"; Description = "Office Feature Updates" },
    @{ Name = "Microsoft\Office\Office Feature Updates Logon"; Description = "Office Feature Updates Logon" },
    @{ Name = "Microsoft\Office\OfficeTelemetryAgent"; Description = "Office Telemetry Agent" },
    @{ Name = "Microsoft\Office\Office 15 Subscription Heartbeat"; Description = "Office 2013 Subscription Heartbeat" },
    @{ Name = "Microsoft\Office\Office 16 Subscription Heartbeat"; Description = "Office 2016/2019 Subscription Heartbeat" },
    @{ Name = "Microsoft\Office\Office Subscription Maintenance"; Description = "Office Subscription Maintenance" },
    @{ Name = "Microsoft\Office\Office ClickToRun Service Monitor"; Description = "Office ClickToRun Service Monitor" }
)

Write-Host "`nProcessing telemetry agent scheduled tasks..." -ForegroundColor $Colors.Info
$script:telemetryTasksProcessed = 0
$script:telemetryTasksDisabled = 0
$script:telemetryTasksNotFound = 0

foreach ($task in $telemetryTasks) {
    $script:telemetryTasksProcessed++
    Disable-ScheduledTaskWithLogging -TaskName $task.Name -Description $task.Description
}

# Block telemetry domains via hosts file
Write-Host "`n--- Blocking telemetry domains ---" -ForegroundColor $Colors.Section

$hostsFile = "$env:windir\System32\drivers\etc\hosts"
$backupFile = "$hostsFile.backup.$(Get-Date -Format 'yyyyMMdd_HHmmss')"
$telemetryDomains = @(
    "vortex.data.microsoft.com",
    "vortex-win.data.microsoft.com",
    "telecommand.telemetry.microsoft.com",
    "telecommand.telemetry.microsoft.com.nsatc.net",
    "oca.telemetry.microsoft.com",
    "oca.telemetry.microsoft.com.nsatc.net",
    "sqm.telemetry.microsoft.com",
    "sqm.telemetry.microsoft.com.nsatc.net",
    "watson.telemetry.microsoft.com",
    "watson.telemetry.microsoft.com.nsatc.net",
    "redir.metaservices.microsoft.com",
    "settings-sandbox.data.microsoft.com",
    "vortex-sandbox.data.microsoft.com",
    "survey.watson.microsoft.com",
    "watson.live.com",
    "watson.microsoft.com",
    "feedback.search.microsoft.com",
    "i1.services.social.microsoft.com",
    "i1.services.social.microsoft.com.nsatc.net",
    "corp.sts.microsoft.com",
    "diagnostics.support.microsoft.com",
    "statsfe2.ws.microsoft.com",
    "pre.footprintpredict.com",
    "i.s1.social.ms.akadns.net",
    "settings-win.data.microsoft.com",
    "diagnostics.office.com",
    "officeclient.microsoft.com",
    "wer.microsoft.com",
    "v10c.events.data.microsoft.com",
    "v10.events.data.microsoft.com",
    "v20.events.data.microsoft.com",
    "client.wns.windows.com",
    "appexsin.trafficmanager.net",
    "appex-rf.msn.com"
)

$applyHostsBlock = Read-Host "Do you want to block telemetry domains in the hosts file? (y/n)"
if ($applyHostsBlock -eq 'y' -or $applyHostsBlock -eq 'Y') {
    try {
        try {
            if (Get-Command "Add-MpPreference" -ErrorAction SilentlyContinue) {
                Add-MpPreference -ExclusionPath "$env:SystemRoot\System32\drivers\etc\hosts" -ErrorAction Stop
                Write-Host "  [OK] Hosts file added to Windows Defender exclusions" -ForegroundColor $Colors.Success
            } else {
                Write-Host "  [WARN] Maybe defender deleted :)" -ForegroundColor $Colors.Skip
            }
        } catch {
            Write-Host "  [ERROR] Unable to add $($_.Exception.Message)" -ForegroundColor $Colors.Error
        }
        
        Copy-Item -Path $hostsFile -Destination $backupFile -Force
        Write-Host "  [OK] Created backup of hosts file: $backupFile" -ForegroundColor $Colors.OK

        $hostsFileInfo = Get-Item $hostsFile -ErrorAction Stop
        $originalAttributes = $hostsFileInfo.Attributes
        $wasReadOnly = $hostsFileInfo.IsReadOnly
        
        if ($wasReadOnly) {
            Write-Host "  [INFO] Hosts file is read-only, temporarily removing read-only attribute..." -ForegroundColor $Colors.Warning
            Set-ItemProperty -Path $hostsFile -Name IsReadOnly -Value $false
        }
        
        $hostsContent = Get-Content -Path $hostsFile -Raw -ErrorAction SilentlyContinue
        if (-not $hostsContent) { $hostsContent = "" }
        
        $addedCount = 0
        $skippedCount = 0
        foreach ($domain in $telemetryDomains) {
            if ($hostsContent -notmatch "0\.0\.0\.0\s+$([regex]::Escape($domain))" -and
                $hostsContent -notmatch "127\.0\.0\.1\s+$([regex]::Escape($domain))") {
                Add-Content -Path $hostsFile -Value "0.0.0.0 $domain" -Force
                Write-Host "  [OK] Blocked domain: $domain" -ForegroundColor $Colors.OK
                $addedCount++
            } else {
                Write-Host "  [SKIP] Domain already blocked: $domain" -ForegroundColor $Colors.Skip
                $skippedCount++
            }
        }
        
        if ($addedCount -eq 0) {
            Write-Host "  [SKIP] All telemetry domains already blocked" -ForegroundColor $Colors.Skip
        } else {
            Write-Host "  [OK] Added $addedCount new entries to hosts file" -ForegroundColor $Colors.Success
        }

        Write-Host "`n  Hosts blocking summary:" -ForegroundColor $Colors.Info
        Write-Host "    Total hosts: $($telemetryDomains.Count)" -ForegroundColor $Colors.Info
        Write-Host "    Newly blocked: $addedCount" -ForegroundColor $Colors.Changed
        Write-Host "    Already blocked: $skippedCount" -ForegroundColor $Colors.Success
        
        # Відновлення оригінальних атрибутів файлу
        try {
            Set-ItemProperty -Path $hostsFile -Name Attributes -Value $originalAttributes
            Write-Host "  [OK] Restored original file attributes" -ForegroundColor $Colors.Success
        } catch {
            Write-Host "  [WARN] Could not restore original attributes: $($_.Exception.Message)" -ForegroundColor $Colors.Warning
            # Fallback - принаймні встановлюємо IsReadOnly якщо воно було
            if ($wasReadOnly) {
                try {
                    Set-ItemProperty -Path $hostsFile -Name IsReadOnly -Value $true
                    Write-Host "  [OK] Restored read-only attribute as fallback" -ForegroundColor $Colors.Success
                } catch {
                    Write-Host "  [ERROR] Could not restore read-only attribute: $($_.Exception.Message)" -ForegroundColor $Colors.Error
                }
            }
        }
        
        try {
            $flushSuccess = $false
            try {
                & "$env:SystemRoot\System32\ipconfig.exe" /flushdns | Out-Null
                $flushSuccess = $true
            } catch {
                try {
                    if (Get-Command "Clear-DnsClientCache" -ErrorAction SilentlyContinue) {
                        Clear-DnsClientCache -ErrorAction Stop
                        $flushSuccess = $true
                    }
                } catch { }
            }
            
            if ($flushSuccess) {
                Write-Host "  [OK] DNS cache flushed" -ForegroundColor $Colors.Success
            } else {
                Write-Host "  [WARN] Could not flush DNS cache (not critical)" -ForegroundColor $Colors.Warning
            }
        } catch {
            Write-Host "  [WARN] Could not flush DNS cache (not critical)" -ForegroundColor $Colors.Warning
        }
    }
    catch {
        Write-Host "  [ERROR] Failed to update hosts file" -ForegroundColor $Colors.Error
        Write-Host "    Error: $_" -ForegroundColor $Colors.Error
        if (Test-Path $hostsFile) {
            try {
                Set-ItemProperty -Path $hostsFile -Name Attributes -Value $originalAttributes
                Write-Host "  [OK] Restored original attributes after error" -ForegroundColor $Colors.Warning
            } catch {
                if ($wasReadOnly) {
                    try {
                        Set-ItemProperty -Path $hostsFile -Name IsReadOnly -Value $true
                        Write-Host "  [OK] Restored read-only attribute after error" -ForegroundColor $Colors.Warning
                    } catch {
                        Write-Host "  [ERROR] Could not restore attributes after error: $($_.Exception.Message)" -ForegroundColor $Colors.Error
                    }
                }
            }
        }
    }
} else {
    Write-Host "  [SKIP] Hosts file modification skipped by user" -ForegroundColor $Colors.Skip
}

# Summary
Write-Host ("`n" + ("=" * 75)) -ForegroundColor $Colors.Title
Write-Host ("Script Execution Summary".PadLeft(50)) -ForegroundColor $Colors.Title
Write-Host ("=" * 75) -ForegroundColor $Colors.Title
Write-Host "`n  > Office versions processed: $($installedVersions -join ', ')" -ForegroundColor $Colors.Success
Write-Host "  > Registry settings applied or skipped: (Check output above for details)" -ForegroundColor $Colors.Success
Write-Host "  > Scheduled tasks processed: $($script:telemetryTasksProcessed)" -ForegroundColor $Colors.Success
Write-Host "    - Disabled: $($script:telemetryTasksDisabled)" -ForegroundColor $Colors.Changed
Write-Host "    - Not found: $($script:telemetryTasksNotFound)" -ForegroundColor $Colors.NotFound
Write-Host "  > Hosts file: $(if ($applyHostsBlock -eq 'y' -or $applyHostsBlock -eq 'Y') { 'Modified' } else { 'Skipped' })" -ForegroundColor $Colors.Success
Write-Host "`n Note: Some changes may require restarting Office applications to take effect." -ForegroundColor $Colors.Warning
Write-Host "`n Script completed successfully!" -ForegroundColor $Colors.Success
Write-Host ("`n" + ("=" * 75)) -ForegroundColor $Colors.Title
