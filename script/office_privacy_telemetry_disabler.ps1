<#
.SYNOPSIS
    Office Privacy & Telemetry Disabler (PowerShell 7 edition)
.DESCRIPTION
    Disables logging, telemetry, CEIP, feedback, Connected Experiences,
    scheduled telemetry tasks, updates and optionally blocks telemetry hosts
    in the HOSTS file (with Windows-Defender exclusion).
.NOTES
    Author : EXLOUD
    License: MIT
    GitHub : https://github.com/EXLOUD
#>

[CmdletBinding(SupportsShouldProcess)]
param (
	[switch]$BlockHosts,
	[switch]$Force
)

# ── Color scheme ─────────────────────────────────────────────────────────────
$Colors = @{
	Title    = 'Cyan'
	Section  = 'Yellow'
	Success  = 'Green'
	Info	 = 'Blue'
	Warning  = 'Yellow'
	Error    = 'Red'
	NotFound = 'Gray'
	Changed  = 'Magenta'
}

function Write-C
{
	param (
		$t,
		$c = 'White'
	)
	$color = if ($Colors.ContainsKey($c)) { $Colors[$c] }
	else { 'White' }
	Write-Host $t -ForegroundColor $color
}
function Write-S { param ($t) Write-C "`n--- $t ---" Section }
function Write-O { param ($t) Write-C "  ✓ $t" Success }
function Write-W { param ($t) Write-C "  ⚠ $t" Warning }
function Write-E { param ($t) Write-C "  ✗ $t" Error }
function Write-N { param ($t) Write-C "  → $t" NotFound }
function Write-X { param ($t) Write-C "  ✓ $t" Changed }

# ── Admin check ───────────────────────────────────────────────────────────────
$isAdmin = ([Security.Principal.WindowsPrincipal]`
	[Security.Principal.WindowsIdentity]::GetCurrent()).IsInRole(
	[Security.Principal.WindowsBuiltInRole]::Administrator)
if (-not $isAdmin -and -not $Force)
{
	Write-E "Administrator rights required. Re-run as Admin or use -Force."
	exit 1
}

# ── Office versions ───────────────────────────────────────────────────────────
$OfficeMap = @{
	'14.0' = 'Office 2010'
	'15.0' = 'Office 2013'
	'16.0' = 'Office 2016/2019/365'
	'17.0' = 'Office 2021'
	'18.0' = 'Office 2024'
}

$installed = $OfficeMap.Keys | Where-Object { Test-Path "HKCU:\SOFTWARE\Microsoft\Office\$_" } | Sort-Object
if (-not $installed) { Write-E "No Office installations found."; exit }

# ── Registry helper (skip if key missing) ────────────────────────────────────
function Set-Reg
{
	[CmdletBinding(SupportsShouldProcess)]
	param (
		[string]$Path,
		[string]$Name,
		[string]$Type = 'DWord',
		[object]$Value,
		[string]$Desc = ''
	)
	if (-not (Test-Path $Path)) { New-Item $Path -Force | Out-Null }
	if ($PSCmdlet.ShouldProcess("$Path\$Name", "Set $Value"))
	{
		Set-ItemProperty -Path $Path -Name $Name -Type $Type -Value $Value -Force
		Write-X "$Name → $Value ($Desc)"
	}
}

# ── Registry tweaks ───────────────────────────────────────────────────────────
Write-S "Applying registry privacy settings"
foreach ($ver in $installed)
{
	Write-C "`nProcessing $($OfficeMap[$ver])" Info
	Set-Reg -Path "HKCU:\SOFTWARE\Microsoft\Office\$ver\Outlook\Options\Mail" -Name EnableLogging -Value 0 -Desc "Outlook Mail"
	Set-Reg -Path "HKCU:\SOFTWARE\Microsoft\Office\$ver\Outlook\Options\Calendar" -Name EnableCalendarLogging -Value 0 -Desc "Calendar"
	Set-Reg -Path "HKCU:\SOFTWARE\Microsoft\Office\$ver\Word\Options" -Name EnableLogging -Value 0 -Desc "Word"
	Set-Reg -Path "HKCU:\SOFTWARE\Policies\Microsoft\Office\$ver\OSM" -Name EnableLogging -Value 0 -Desc "OSM"
	Set-Reg -Path "HKCU:\SOFTWARE\Policies\Microsoft\Office\$ver\OSM" -Name EnableUpload -Value 0 -Desc "OSM Upload"
	Set-Reg -Path "HKCU:\SOFTWARE\Microsoft\Office\$ver\Common\ClientTelemetry" -Name DisableTelemetry -Value 1 -Desc "Telemetry"
	Set-Reg -Path "HKCU:\SOFTWARE\Microsoft\Office\$ver\Common\ClientTelemetry" -Name VerboseLogging -Value 0 -Desc "Verbose"
	Set-Reg -Path "HKCU:\SOFTWARE\Microsoft\Office\$ver\Common" -Name QMEnable -Value 0 -Desc "CEIP"
	Set-Reg -Path "HKCU:\SOFTWARE\Microsoft\Office\$ver\Common\Feedback" -Name Enabled -Value 0 -Desc "Feedback"
	if ($ver -ge '16.0')
	{
		Set-Reg -Path "HKCU:\SOFTWARE\Microsoft\Office\$ver\Common\Privacy" -Name UserContentDisabled -Value 2 -Desc "User content"
		Set-Reg -Path "HKCU:\SOFTWARE\Microsoft\Office\$ver\Common\Privacy" -Name DownloadContentDisabled -Value 2 -Desc "Download content"
		Set-Reg -Path "HKCU:\SOFTWARE\Microsoft\Office\$ver\Common\Internet" -Name UseOnlineContent -Value 0 -Desc "Online content"
	}
}

# ── Scheduled-task handler (PowerShell 7) ─────────────────────────────────────
function Disable-Task
{
	param (
		[string]$Name,
		[string]$Description
	)
	
	$task = Get-ScheduledTask -TaskName $Name -ErrorAction SilentlyContinue
	if (-not $task)
	{
		Write-O "Task not found (OK): $Name"
		return @{ Status = 'NotFound' }
	}
	
	if ($task.State -eq 'Disabled')
	{
		Write-O "Already disabled: $Name"
		return @{ Status = 'AlreadyDisabled' }
	}
	
	try
	{
		Disable-ScheduledTask -TaskName $Name -ErrorAction Stop | Out-Null
		Write-X "Disabled: $Name ($Description)"
		return @{ Status = 'Disabled' }
	}
	catch
	{
		Write-E "Failed to disable $Name : $($_.Exception.Message)"
		return @{ Status = 'Error' }
	}
}

# ── Task list (same as before) ────────────────────────────────────────────────
$tasks = @(
	@{ Name = 'Microsoft\Office\OfficeTelemetryAgentFallBack'; Desc = 'Telemetry Agent 2013' },
	@{ Name = 'Microsoft\Office\OfficeTelemetryAgentLogOn'; Desc = 'Telemetry Logon 2013' },
	@{ Name = 'Microsoft\Office\OfficeTelemetryAgentFallBack2016'; Desc = 'Telemetry Agent 2016+' },
	@{ Name = 'Microsoft\Office\OfficeTelemetryAgentLogOn2016'; Desc = 'Telemetry Logon 2016+' },
	@{ Name = 'Microsoft\Office\OfficeBackgroundTaskHandlerRegistration'; Desc = 'Background registration' },
	@{ Name = 'Microsoft\Office\OfficeBackgroundTaskHandlerLogon'; Desc = 'Background logon' },
	@{ Name = 'Microsoft\Office\Office Automatic Updates'; Desc = 'Auto-updates' },
	@{ Name = 'Microsoft\Office\Office Automatic Updates 2.0'; Desc = 'Auto-updates 2.0' },
	@{ Name = 'Microsoft\Office\Office Feature Updates'; Desc = 'Feature updates' },
	@{ Name = 'Microsoft\Office\Office Feature Updates Logon'; Desc = 'Feature updates logon' },
	@{ Name = 'Microsoft\Office\Office 15 Subscription Heartbeat'; Desc = 'Subscription heartbeat 2013' },
	@{ Name = 'Microsoft\Office\Office 16 Subscription Heartbeat'; Desc = 'Subscription heartbeat 2016+' },
	@{ Name = 'Microsoft\Office\Office Subscription Maintenance'; Desc = 'Subscription maintenance' },
	@{ Name = 'Microsoft\Office\Office ClickToRun Service Monitor'; Desc = 'Click-to-Run monitor' }
)

# ── Processing with counters ───────────────────────────────────────────────────
$stats = @{
	Total		    = $tasks.Count
	Disabled	    = 0
	AlreadyDisabled = 0
	NotFound	    = 0
	Errors		    = 0
}

Write-S "Disabling telemetry scheduled tasks"
foreach ($t in $tasks)
{
	$result = Disable-Task -Name $t.Name -Description $t.Desc
	switch ($result.Status)
	{
		'Disabled'        { $stats.Disabled++ }
		'AlreadyDisabled' { $stats.AlreadyDisabled++ }
		'NotFound'        { $stats.NotFound++ }
		'Error'           { $stats.Errors++ }
	}
}

Write-C ("Tasks processed: {0} | Disabled: {1} | Already-off: {2} | Not-found (OK): {3} | Errors: {4}" `
	-f $stats.Total, $stats.Disabled, $stats.AlreadyDisabled, $stats.NotFound, $stats.Errors) Info

# ── HOSTS file block (with Defender exclusion) ────────────────────────────────
if ($BlockHosts -or ($PSCmdlet.ShouldContinue(
			'Block Microsoft telemetry hosts via hosts file?',
			'Hosts file protection')))
{
	$hostsPath = "$env:SystemRoot\System32\drivers\etc\hosts"
	
	# Defender exclusion
	try
	{
		Add-MpPreference -ExclusionPath $hostsPath -EA Stop
		Write-O "Hosts file added to Windows Defender exclusions"
	}
	catch { Write-W "Could not add Defender exclusion (not critical)" }
	
	$domains = @(
		"vortex.data.microsoft.com", "vortex-win.data.microsoft.com",
		"telecommand.telemetry.microsoft.com", "telecommand.telemetry.microsoft.com.nsatc.net",
		"oca.telemetry.microsoft.com", "oca.telemetry.microsoft.com.nsatc.net",
		"sqm.telemetry.microsoft.com", "sqm.telemetry.microsoft.com.nsatc.net",
		"watson.telemetry.microsoft.com", "watson.telemetry.microsoft.com.nsatc.net",
		"redir.metaservices.microsoft.com", "choice.microsoft.com", "choice.microsoft.com.nsatc.net",
		"df.telemetry.microsoft.com", "reports.wes.df.telemetry.microsoft.com",
		"wes.df.telemetry.microsoft.com", "services.wes.df.telemetry.microsoft.com",
		"sqm.df.telemetry.microsoft.com", "telemetry.microsoft.com",
		"watson.ppe.telemetry.microsoft.com", "telemetry.appex.bing.net", "telemetry.urs.microsoft.com",
		"settings-sandbox.data.microsoft.com", "vortex-sandbox.data.microsoft.com",
		"survey.watson.microsoft.com", "watson.live.com", "watson.microsoft.com",
		"statsfe2.ws.microsoft.com", "corpext.msitadfs.glbdns2.microsoft.com",
		"compatexchange.cloudapp.net", "cs1.wpc.v0cdn.net", "a-0001.a-msedge.net",
		"statsfe2.update.microsoft.com.akadns.net", "sls.update.microsoft.com.akadns.net",
		"fe2.update.microsoft.com.akadns.net", "diagnostics.support.microsoft.com",
		"corp.sts.microsoft.com", "statsfe1.ws.microsoft.com", "pre.footprintpredict.com",
		"i1.services.social.microsoft.com", "i1.services.social.microsoft.com.nsatc.net",
		"feedback.windows.com", "feedback.microsoft-hohm.com", "feedback.search.microsoft.com"
	)
	
	$backup = "$hostsPath.backup.$(Get-Date -Format 'yyyyMMddHHmmss')"
	Copy-Item $hostsPath $backup
	Write-O "Hosts backed up to $backup"
	
	$hosts = Get-Content $hostsPath -Raw
	$new = $domains | Where-Object { $hosts -notmatch [regex]::Escape($_) } |
	ForEach-Object { "0.0.0.0 $_" }
	
	if ($new)
	{
		$new = @("`n# Office-Telemetry Block - $(Get-Date)") + $new + "# End of block`n"
		Add-Content $hostsPath $new
		Write-X "Added $($new.Count - 2) blocked domains"
	}
	else { Write-N "All domains already blocked" }
	
	Clear-DnsClientCache -EA SilentlyContinue
	Write-O "DNS cache flushed"
}

# ── Summary ───────────────────────────────────────────────────────────────────
Write-S "Summary"
$installed | ForEach-Object { Write-C "  ✓ $($OfficeMap[$_])" Found }
Write-C "`nRestart Office applications for changes to take effect." Warning
Write-C "`nScript completed successfully!" Success
