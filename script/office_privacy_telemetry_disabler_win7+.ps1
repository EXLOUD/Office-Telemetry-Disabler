# Office Privacy and Telemetry Disabler (Windows 7+ compatibility)
# PowerShell script to disable Microsoft Office logging, telemetry, and privacy features
# by EXLOUD
# >> https://github.com/EXLOUD <<

# Color scheme for consistent output
$Colors = @{
	Title    = 'Cyan'
	Section  = 'Yellow'
	Success  = 'Green'
	Info	 = 'Blue'
	Warning  = 'Yellow'
	Error    = 'Red'
	Gray	 = 'Gray'
	Found    = 'Green'
	Changed  = 'Magenta'
	NotFound = 'Gray'
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

$OfficeVersions = @{
	'14.0' = 'Office 2010'
	'15.0' = 'Office 2013'
	'16.0' = 'Office 2016/2019'
	'17.0' = 'Office 2021'
	'18.0' = 'Office 2024'
}

# Function to write colored output (без емоджі)
function Write-ColoredOutput {
	param (
		[string]$Message,
		[string]$Color = 'White'
	)
	
	# Map custom colors to PowerShell colors
	$psColor = switch ($Color) {
		'Warning' { $Colors.Warning }
		'Found' { $Colors.Found }
		'Header' { $Colors.Section }
		'Changed' { $Colors.Changed }
		'NotFound' { $Colors.NotFound }
		'Info' { $Colors.Info }
		'Success' { $Colors.Success }
		'Error' { $Colors.Error }
		default { 'White' }
	}
	
	Write-Host $Message -ForegroundColor $psColor
}

# Function to check if Office version is installed
function Test-OfficeVersion {
	param ([string]$Version)
	
	$registryPath = "HKCU:\SOFTWARE\Microsoft\Office\$Version"
	return Test-Path $registryPath
}

# Function to get installed Office version
function Get-InstalledOfficeVersion {
	$installedVersions = @()
	foreach ($version in $OfficeVersions.Keys) {
		if (Test-OfficeVersion $version) {
			$installedVersions += $version
		}
	}
	# Sort versions in ascending order
	return $installedVersions | Sort-Object
}

# Function to set registry value with logging
function Set-RegistryValueWithLogging {
	[CmdletBinding(SupportsShouldProcess)]
	param (
		[string]$Path,
		[string]$Name,
		[string]$Type,
		[object]$Value,
		[string]$Description = ""
	)
	
	if ($PSCmdlet.ShouldProcess("Registry: $Path\$Name", "Set value to $Value")) {
		try {
			# Check if the registry key exists
			if (-not (Test-Path $Path)) {
				Write-Host "  [SKIP] Registry path not found: $Path" -ForegroundColor $Colors.NotFound
				return
			}
			
			# Create the registry key if it doesn't exist
			if (-not (Test-Path $Path)) {
				New-Item -Path $Path -Force | Out-Null
			}
			
			# Set the registry value
			Set-ItemProperty -Path $Path -Name $Name -Value $Value -Type $Type -Force
			
			Write-Host "  [OK] Found and changed: $Name -> $Value" -ForegroundColor $Colors.Changed
			if ($Description) {
				Write-Host "    Description: $Description" -ForegroundColor $Colors.Info
			}
		} catch {
			Write-Host "  [ERROR] Error setting $Name : $($_.Exception.Message)" -ForegroundColor $Colors.Error
		}
	}
}

# Function to disable scheduled task with logging (сумісна з Windows 7)
function Disable-ScheduledTaskWithLogging {
	param (
		[string]$TaskName,
		[string]$Description = ""
	)
	
	try {
		# Перевіряємо чи є Get-ScheduledTask (Windows 8+)
		if (Get-Command "Get-ScheduledTask" -ErrorAction SilentlyContinue) {
			# Windows 8+
			$task = Get-ScheduledTask -TaskName $TaskName -ErrorAction SilentlyContinue
			if ($task) {
				$taskState = $task.State
				if ($taskState -eq 'Disabled') {
					Write-Host "  [OK] Task already disabled: $TaskName" -ForegroundColor $Colors.Success
				} else {
					Disable-ScheduledTask -TaskName $TaskName -ErrorAction Stop | Out-Null
					Write-Host "  [OK] Found and disabled: $TaskName" -ForegroundColor $Colors.Changed
				}
				
				if ($Description) {
					Write-Host "    Description: $Description" -ForegroundColor $Colors.Info
				}
			} else {
				Write-Host "  [SKIP] Task not found: $TaskName" -ForegroundColor $Colors.NotFound
			}
		} else {
			# Windows 7 - використовуємо schtasks.exe
			$taskPath = $TaskName -replace "Microsoft\\", "\"
			
			# Перевіряємо чи існує завдання
			$queryResult = & schtasks.exe /Query /TN $taskPath 2>$null
			if ($LASTEXITCODE -eq 0) {
				# Завдання існує, перевіряємо статус
				$taskInfo = & schtasks.exe /Query /TN $taskPath /FO CSV 2>$null | ConvertFrom-Csv
				if ($taskInfo -and $taskInfo.Status -eq "Disabled") {
					Write-Host "  [OK] Task already disabled: $TaskName" -ForegroundColor $Colors.Success
				} else {
					# Відключаємо завдання
					& schtasks.exe /Change /TN $taskPath /DISABLE >$null 2>&1
					if ($LASTEXITCODE -eq 0) {
						Write-Host "  [OK] Found and disabled: $TaskName" -ForegroundColor $Colors.Changed
					} else {
						Write-Host "  [ERROR] Failed to disable: $TaskName" -ForegroundColor $Colors.Error
					}
				}
				
				if ($Description) {
					Write-Host "    Description: $Description" -ForegroundColor $Colors.Info
				}
			} else {
				Write-Host "  [SKIP] Task not found: $TaskName" -ForegroundColor $Colors.NotFound
			}
		}
	} catch {
		Write-Host "  [ERROR] Error disabling $TaskName : $($_.Exception.Message)" -ForegroundColor $Colors.Error
	}
}

# Main script execution
Write-Host "`n==========================================" -ForegroundColor $Colors.Title
Write-Host "   Office Privacy and Telemetry Disabler" -ForegroundColor $Colors.Title
Write-Host "                by EXLOUD" -ForegroundColor $Colors.Title
Write-Host "     >> https://github.com/EXLOUD <<" -ForegroundColor $Colors.Title
Write-Host "==========================================" -ForegroundColor $Colors.Title

# Check for installed Office versions
Write-Host "`n--- Checking for installed Office versions ---" -ForegroundColor $Colors.Section
$installedVersions = Get-InstalledOfficeVersion

if ($installedVersions.Count -eq 0) {
	Write-Host "[ERROR] No Office installations found in registry." -ForegroundColor $Colors.Error
	Read-Host "Press Enter to exit"
	exit
}

Write-Host "Found Office versions:" -ForegroundColor $Colors.Found
foreach ($version in $installedVersions) {
	Write-Host "  [OK] $($OfficeVersions[$version]) (Version $version)" -ForegroundColor $Colors.Found
}

# ----------------------------------------------------------
# Disable Microsoft Office logging
# ----------------------------------------------------------
Write-Host "`n--- Disabling Microsoft Office logging ---" -ForegroundColor $Colors.Section

foreach ($version in $installedVersions) {
	Write-Host "`nProcessing $($OfficeVersions[$version])..." -ForegroundColor $Colors.Info
	
	# Outlook logging
	Set-RegistryValueWithLogging -Path "HKCU:\SOFTWARE\Microsoft\Office\$version\Outlook\Options\Mail" -Name "EnableLogging" -Type "DWord" -Value 0 -Description "Disable Outlook mail logging"
	Set-RegistryValueWithLogging -Path "HKCU:\SOFTWARE\Microsoft\Office\$version\Outlook\Options\Calendar" -Name "EnableCalendarLogging" -Type "DWord" -Value 0 -Description "Disable Outlook calendar logging"
	
	# Word logging
	Set-RegistryValueWithLogging -Path "HKCU:\SOFTWARE\Microsoft\Office\$version\Word\Options" -Name "EnableLogging" -Type "DWord" -Value 0 -Description "Disable Word logging"
	
	# OSM (Office Service Manager)
	Set-RegistryValueWithLogging -Path "HKCU:\SOFTWARE\Policies\Microsoft\Office\$version\OSM" -Name "EnableLogging" -Type "DWord" -Value 0 -Description "Disable OSM logging"
	Set-RegistryValueWithLogging -Path "HKCU:\SOFTWARE\Policies\Microsoft\Office\$version\OSM" -Name "EnableUpload" -Type "DWord" -Value 0 -Description "Disable OSM upload"
}

# ----------------------------------------------------------
# Disable client telemetry
# ----------------------------------------------------------
Write-Host "`n--- Disabling client telemetry ---" -ForegroundColor $Colors.Section

# Common telemetry settings
Set-RegistryValueWithLogging -Path "HKCU:\SOFTWARE\Microsoft\Office\Common\ClientTelemetry" -Name "DisableTelemetry" -Type "DWord" -Value 1 -Description "Disable common telemetry"
Set-RegistryValueWithLogging -Path "HKCU:\SOFTWARE\Microsoft\Office\Common\ClientTelemetry" -Name "VerboseLogging" -Type "DWord" -Value 0 -Description "Disable verbose logging"

foreach ($version in $installedVersions) {
	Write-Host "`nProcessing $($OfficeVersions[$version]) telemetry..." -ForegroundColor $Colors.Info
	
	Set-RegistryValueWithLogging -Path "HKCU:\SOFTWARE\Microsoft\Office\$version\Common\ClientTelemetry" -Name "DisableTelemetry" -Type "DWord" -Value 1 -Description "Disable telemetry"
	Set-RegistryValueWithLogging -Path "HKCU:\SOFTWARE\Microsoft\Office\$version\Common\ClientTelemetry" -Name "VerboseLogging" -Type "DWord" -Value 0 -Description "Disable verbose logging"
}

# ----------------------------------------------------------
# Disable Customer Experience Improvement Program
# ----------------------------------------------------------
Write-Host "`n--- Disabling Customer Experience Improvement Program ---" -ForegroundColor $Colors.Section

# Common CEIP settings
Set-RegistryValueWithLogging -Path "HKCU:\SOFTWARE\Microsoft\Office\Common" -Name "QMEnable" -Type "DWord" -Value 0 -Description "Disable common CEIP"

foreach ($version in $installedVersions) {
	Write-Host "`nProcessing $($OfficeVersions[$version]) CEIP..." -ForegroundColor $Colors.Info
	
	Set-RegistryValueWithLogging -Path "HKCU:\SOFTWARE\Microsoft\Office\$version\Common" -Name "QMEnable" -Type "DWord" -Value 0 -Description "Disable CEIP"
}

# ----------------------------------------------------------
# Disable feedback
# ----------------------------------------------------------
Write-Host "`n--- Disabling feedback ---" -ForegroundColor $Colors.Section

# Common feedback settings
Set-RegistryValueWithLogging -Path "HKCU:\SOFTWARE\Microsoft\Office\Common\Feedback" -Name "Enabled" -Type "DWord" -Value 0 -Description "Disable common feedback"

foreach ($version in $installedVersions) {
	Write-Host "`nProcessing $($OfficeVersions[$version]) feedback..." -ForegroundColor $Colors.Info
	
	Set-RegistryValueWithLogging -Path "HKCU:\SOFTWARE\Microsoft\Office\$version\Common\Feedback" -Name "Enabled" -Type "DWord" -Value 0 -Description "Disable feedback"
}

# ----------------------------------------------------------
# Disable Connected Experiences (Office 2016 and later)
# ----------------------------------------------------------
Write-Host "`n--- Disabling Connected Experiences ---" -ForegroundColor $Colors.Section

$modernVersions = $installedVersions | Where-Object { $_ -in @('16.0', '17.0', '18.0') }

if ($modernVersions.Count -gt 0) {
	foreach ($version in $modernVersions) {
		Write-Host "`nProcessing $($OfficeVersions[$version]) Connected Experiences..." -ForegroundColor $Colors.Info
		
		Set-RegistryValueWithLogging -Path "HKCU:\SOFTWARE\Microsoft\Office\$version\Common\Privacy" -Name "UserContentDisabled" -Type "DWord" -Value 2 -Description "Disable user content experiences"
		Set-RegistryValueWithLogging -Path "HKCU:\SOFTWARE\Microsoft\Office\$version\Common\Privacy" -Name "DownloadContentDisabled" -Type "DWord" -Value 2 -Description "Disable download content experiences"
		Set-RegistryValueWithLogging -Path "HKCU:\SOFTWARE\Microsoft\Office\$version\Common\Internet" -Name "UseOnlineContent" -Type "DWord" -Value 0 -Description "Disable online content"
	}
} else {
	Write-Host "[SKIP] No modern Office versions found for Connected Experiences settings." -ForegroundColor $Colors.NotFound
}

# ----------------------------------------------------------
# Disable telemetry agent scheduled tasks
# ----------------------------------------------------------
Write-Host "`n--- Disabling telemetry agent scheduled tasks ---" -ForegroundColor $Colors.Section

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
	@{ Name = "Microsoft\Office\Office Feature Updates Logon"; Description = "Office Feature Updates Logon" }
)

Write-Host "`nProcessing telemetry agent scheduled tasks..." -ForegroundColor $Colors.Info
$telemetryTasksProcessed = 0
$telemetryTasksDisabled = 0
$telemetryTasksNotFound = 0

foreach ($task in $telemetryTasks) {
	$telemetryTasksProcessed++
	Disable-ScheduledTaskWithLogging -TaskName $task.Name -Description $task.Description
}

# ----------------------------------------------------------
# Disable Subscription Heartbeat
# ----------------------------------------------------------
Write-Host "`n--- Disabling Subscription Heartbeat ---" -ForegroundColor $Colors.Section

$heartbeatTasks = @(
	@{ Name = "Microsoft\Office\Office 15 Subscription Heartbeat"; Description = "Office 2013 Subscription Heartbeat" },
	@{ Name = "Microsoft\Office\Office 16 Subscription Heartbeat"; Description = "Office 2016/2019 Subscription Heartbeat" },
	@{ Name = "Microsoft\Office\Office Subscription Maintenance"; Description = "Office Subscription Maintenance" },
	@{ Name = "Microsoft\Office\Office ClickToRun Service Monitor"; Description = "Office ClickToRun Service Monitor" }
)

Write-Host "`nProcessing subscription heartbeat tasks..." -ForegroundColor $Colors.Info
$heartbeatTasksProcessed = 0
$heartbeatTasksDisabled = 0
$heartbeatTasksNotFound = 0

foreach ($task in $heartbeatTasks) {
	$heartbeatTasksProcessed++
	Disable-ScheduledTaskWithLogging -TaskName $task.Name -Description $task.Description
}

# ----------------------------------------------------------
# Disable Office updates and notifications
# ----------------------------------------------------------
Write-Host "`n--- Disabling Office updates and notifications ---" -ForegroundColor $Colors.Section

$updateVersions = $installedVersions | Where-Object { $_ -in @('16.0', '17.0', '18.0') }

if ($updateVersions.Count -gt 0) {
	foreach ($version in $updateVersions) {
		Write-Host "`nProcessing $($OfficeVersions[$version]) updates..." -ForegroundColor $Colors.Info
		
		Set-RegistryValueWithLogging -Path "HKCU:\SOFTWARE\Microsoft\Office\$version\Common\OfficeUpdate" -Name "OfficeMgmtCOM" -Type "DWord" -Value 0 -Description "Disable Office management COM"
		Set-RegistryValueWithLogging -Path "HKCU:\SOFTWARE\Microsoft\Office\$version\Common\OfficeUpdate" -Name "EnableAutomaticUpdates" -Type "DWord" -Value 0 -Description "Disable automatic updates"
		Set-RegistryValueWithLogging -Path "HKCU:\SOFTWARE\Microsoft\Office\$version\Common\PTWatson" -Name "PTWOptIn" -Type "DWord" -Value 0 -Description "Disable Watson error reporting"
	}
} else {
	Write-Host "[SKIP] No modern Office versions found for update settings." -ForegroundColor $Colors.NotFound
}

# ----------------------------------------------------------
# Summary
# ----------------------------------------------------------
Write-Host "`n==========================================" -ForegroundColor $Colors.Title
Write-Host "                  Summary" -ForegroundColor $Colors.Title
Write-Host "==========================================" -ForegroundColor $Colors.Title

Write-Host "`nOffice privacy and telemetry settings have been applied to:" -ForegroundColor $Colors.Info
foreach ($version in $installedVersions) {
	Write-Host "  [OK] $($OfficeVersions[$version])" -ForegroundColor $Colors.Found
}

Write-Host "`nScheduled tasks processed:" -ForegroundColor $Colors.Info
Write-Host "  > Telemetry tasks: $($telemetryTasks.Count) processed" -ForegroundColor $Colors.Success
Write-Host "  > Heartbeat tasks: $($heartbeatTasks.Count) processed" -ForegroundColor $Colors.Success

Write-Host "`nLegend:" -ForegroundColor White
Write-Host "[OK] " -NoNewline -ForegroundColor $Colors.Success; Write-Host "Action completed successfully"
Write-Host "[OK] " -NoNewline -ForegroundColor $Colors.Changed; Write-Host "Setting changed"
Write-Host "[INFO] " -NoNewline -ForegroundColor $Colors.Info; Write-Host "Information or preparatory action"
Write-Host "[SKIP] " -NoNewline -ForegroundColor $Colors.NotFound; Write-Host "Component not found, skipped"
Write-Host "[ERROR] " -NoNewline -ForegroundColor $Colors.Error; Write-Host "Error occurred"

Write-Host "`nNote: Some changes may require restarting Office applications to take effect." -ForegroundColor $Colors.Warning

# ----------------------------------------------------------
# Additional option: Block Microsoft Office telemetry hosts
# ----------------------------------------------------------
Write-Host "`n--- Additional Option: Block Telemetry Hosts ---" -ForegroundColor $Colors.Section
$blockHosts = Read-Host "Do you want to block Microsoft Office telemetry hosts in the hosts file? (y/n)"
if ($blockHosts -eq 'y' -or $blockHosts -eq 'Y' -or $blockHosts -eq 'yes' -or $blockHosts -eq 'Yes') {
	Write-Host "`nBlocking Microsoft Office telemetry hosts..." -ForegroundColor $Colors.Info
	
	# Спробуємо додати винятки для Windows Defender (якщо доступно)
	try {
		if (Get-Command "Add-MpPreference" -ErrorAction SilentlyContinue) {
			Add-MpPreference -ExclusionPath "$env:SystemRoot\System32\drivers\etc\hosts" -ErrorAction Stop
			Write-Host "  [OK] Hosts file added to Windows Defender exclusions" -ForegroundColor $Colors.Success
		}
	} catch {
		Write-Host "  [WARN] Could not add hosts file to Defender exclusions (not critical)" -ForegroundColor $Colors.Warning
	}
	
	$hostsToBlock = @(
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
		"choice.microsoft.com",
		"choice.microsoft.com.nsatc.net",
		"df.telemetry.microsoft.com",
		"reports.wes.df.telemetry.microsoft.com",
		"wes.df.telemetry.microsoft.com",
		"services.wes.df.telemetry.microsoft.com",
		"sqm.df.telemetry.microsoft.com",
		"telemetry.microsoft.com",
		"watson.ppe.telemetry.microsoft.com",
		"telemetry.appex.bing.net",
		"telemetry.urs.microsoft.com",
		"telemetry.appex.bing.net",
		"settings-sandbox.data.microsoft.com",
		"vortex-sandbox.data.microsoft.com",
		"survey.watson.microsoft.com",
		"watson.live.com",
		"watson.microsoft.com",
		"statsfe2.ws.microsoft.com",
		"corpext.msitadfs.glbdns2.microsoft.com",
		"compatexchange.cloudapp.net",
		"cs1.wpc.v0cdn.net",
		"a-0001.a-msedge.net",
		"statsfe2.update.microsoft.com.akadns.net",
		"sls.update.microsoft.com.akadns.net",
		"fe2.update.microsoft.com.akadns.net",
		"diagnostics.support.microsoft.com",
		"corp.sts.microsoft.com",
		"statsfe1.ws.microsoft.com",
		"pre.footprintpredict.com",
		"i1.services.social.microsoft.com",
		"i1.services.social.microsoft.com.nsatc.net",
		"feedback.windows.com",
		"feedback.microsoft-hohm.com",
		"feedback.search.microsoft.com"
	)
	
	$hostsFilePath = "$env:SystemRoot\System32\drivers\etc\hosts"
	
	try {
		$hostsFileInfo = Get-Item $hostsFilePath -ErrorAction Stop
		$originalAttributes = $hostsFileInfo.Attributes
		$wasReadOnly = $hostsFileInfo.IsReadOnly
		
		if ($wasReadOnly) {
			Write-Host "  [INFO] Hosts file is read-only, temporarily removing read-only attribute..." -ForegroundColor $Colors.Warning
			Set-ItemProperty -Path $hostsFilePath -Name IsReadOnly -Value $false
		}
		
		$currentHosts = Get-Content $hostsFilePath -ErrorAction Stop
		$backupPath = "$hostsFilePath.backup.$(Get-Date -Format 'yyyyMMdd_HHmmss')"
		Copy-Item $hostsFilePath $backupPath
		Write-Host "  [OK] Hosts file backed up to: $backupPath" -ForegroundColor $Colors.Success
		
		$hostsAdded = 0
		$hostsSkipped = 0
		$newEntries = @()
		
		foreach ($hostname in $hostsToBlock) {
			$hostEntry = "0.0.0.0 $hostname"
			if ($currentHosts -notcontains $hostEntry -and
				$currentHosts -notmatch "127\.0\.0\.1\s+$([regex]::Escape($hostname))" -and
				$currentHosts -notmatch "0\.0\.0\.0\s+$([regex]::Escape($hostname))") {
				$newEntries += $hostEntry
				$hostsAdded++
				Write-Host "  [OK] Will block: $hostname" -ForegroundColor $Colors.Changed
			} else {
				$hostsSkipped++
				Write-Host "  [INFO] Already blocked: $hostname" -ForegroundColor $Colors.Success
			}
		}
		
		if ($newEntries.Count -gt 0) {
			$newEntries = @("`n# Microsoft Office Telemetry Hosts - Blocked by Office Privacy Disabler") + $newEntries + @("# End of Office Telemetry Hosts`n")
			Add-Content $hostsFilePath $newEntries
			Write-Host "`n  [OK] Added $hostsAdded new entries to hosts file" -ForegroundColor $Colors.Success
		}
		
		if ($wasReadOnly) {
			Set-ItemProperty -Path $hostsFilePath -Name IsReadOnly -Value $true
			Write-Host "  [OK] Restored read-only attribute to hosts file" -ForegroundColor $Colors.Success
		}
		
		Write-Host "`n  Hosts blocking summary:" -ForegroundColor $Colors.Info
		Write-Host "    Total hosts: $($hostsToBlock.Count)" -ForegroundColor $Colors.Info
		Write-Host "    Newly blocked: $hostsAdded" -ForegroundColor $Colors.Changed
		Write-Host "    Already blocked: $hostsSkipped" -ForegroundColor $Colors.Success
		
		# Очищення DNS кешу (сумісно з Windows 7)
		try {
			$flushSuccess = $false
			# Спробуємо ipconfig (доступний на всіх версіях Windows)
			try {
				& "$env:SystemRoot\System32\ipconfig.exe" /flushdns | Out-Null
				$flushSuccess = $true
			} catch {
				# Для Windows 8+ спробуємо PowerShell команду
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
		
	} catch {
		Write-Host "  [ERROR] Error modifying hosts file: $($_.Exception.Message)" -ForegroundColor $Colors.Error
		if ($wasReadOnly -and (Test-Path $hostsFilePath)) {
			try {
				Set-ItemProperty -Path $hostsFilePath -Name IsReadOnly -Value $true
				Write-Host "  [OK] Restored read-only attribute after error" -ForegroundColor $Colors.Warning
			} catch {
				Write-Host "  [ERROR] Could not restore read-only attribute: $($_.Exception.Message)" -ForegroundColor $Colors.Error
			}
		}
	}
} else {
	Write-Host "Hosts blocking skipped." -ForegroundColor $Colors.Info
}

Write-Host "`nScript completed successfully!" -ForegroundColor $Colors.Success
Read-Host "Press Enter to exit"
