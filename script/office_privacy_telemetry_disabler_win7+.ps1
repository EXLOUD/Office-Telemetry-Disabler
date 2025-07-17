<#
.SYNOPSIS
    Office Privacy & Telemetry Disabler (Windows 7/8/10/11)

.DESCRIPTION
    Disables Microsoft Office logging, telemetry, CEIP, feedback, Connected Experiences,
    scheduled telemetry tasks, updates and optionally blocks known Microsoft telemetry
    hosts via the HOSTS file (with Windows Defender exclusion).

.NOTES
    Author : EXLOUD
    License: MIT
    GitHub : https://github.com/EXLOUD
    Version: 2.0
#>

[CmdletBinding(SupportsShouldProcess)]
param(
    [switch]$BlockHosts,
    [switch]$Force
)

#region ─── Colour & Logging Helpers ───────────────────────────────────────────────
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
    NotFound = 'DarkGray'
}

function Write-C  { param($t,$c='White')  Write-Host $t -ForegroundColor $Colors[$c] }
function Write-S  { param($t)  Write-C "`n--- $t ---" Section }
function Write-O  { param($t)  Write-C "  [OK] $t"   Success }
function Write-W  { param($t)  Write-C "  [WARN] $t" Warning }
function Write-E  { param($t)  Write-C "  [ERROR] $t" Error }
function Write-X  { param($t)  Write-C "  [CHANGED] $t" Changed }
function Write-N  { param($t)  Write-C "  [SKIP] $t" NotFound }

$isAdmin = ([Security.Principal.WindowsPrincipal]`
           [Security.Principal.WindowsIdentity]::GetCurrent()).IsInRole(
           [Security.Principal.WindowsBuiltInRole]::Administrator)
if (-not $isAdmin -and -not $Force) {
    Write-E "Administrator rights required. Re-run as Admin or use -Force."
    exit 1
}

$os = switch -Regex ([Environment]::OSVersion.Version.ToString()) {
    '^6\.1'  { 'Windows 7' }
    '^6\.[2-3]' { 'Windows 8/8.1' }
    '^10\.'  { 'Windows 10/11' }
    default  { "Windows $_" }
}
Write-C "Detected OS: $os" Info
#endregion

#region ─── Office Discovery ───────────────────────────────────────────────────────
$OfficeMap = @{
    '14.0' = 'Office 2010'
    '15.0' = 'Office 2013'
    '16.0' = 'Office 2016/2019/365'
    '17.0' = 'Office 2021'
    '18.0' = 'Office 2024'
}
$installed = $OfficeMap.Keys | Where-Object {
    Test-Path "HKCU:\SOFTWARE\Microsoft\Office\$_"
} | Sort-Object
if (-not $installed) { Write-E "No Office installations detected."; exit }
#endregion

#region ─── Registry Helper with Roll-back ────────────────────────────────────────
$Rollback = @()
function Set-Reg {
    [CmdletBinding(SupportsShouldProcess)]
    param($Path,$Name,$Type='DWord',$Value,$Desc='')
    if (-not (Test-Path $Path)) { New-Item $Path -Force | Out-Null }
    $old = Get-ItemProperty -Path $Path -Name $Name -EA SilentlyContinue
    if ($null -ne $old.$Name) { $script:Rollback += [pscustomobject]@{Path=$Path;Name=$Name;Old=$old.$Name} }
    if ($PSCmdlet.ShouldProcess("$Path\$Name","Set $Value")) {
        Set-ItemProperty -Path $Path -Name $Name -Type $Type -Value $Value -Force
        Write-X "$Name -> $Value ($Desc)"
    }
}
#endregion

#region ─── Registry Tweaks ───────────────────────────────────────────────────────
Write-S "Applying registry privacy settings"
foreach ($ver in $installed) {
    Write-C "`nProcessing $($OfficeMap[$ver])" Info
    Set-Reg "HKCU:\SOFTWARE\Microsoft\Office\$ver\Outlook\Options\Mail" EnableLogging 0 "Outlook Mail logging"
    Set-Reg "HKCU:\SOFTWARE\Microsoft\Office\$ver\Outlook\Options\Calendar" EnableCalendarLogging 0 "Calendar logging"
    Set-Reg "HKCU:\SOFTWARE\Microsoft\Office\$ver\Word\Options" EnableLogging 0 "Word logging"
    Set-Reg "HKCU:\SOFTWARE\Policies\Microsoft\Office\$ver\OSM" EnableLogging 0 "OSM logging"
    Set-Reg "HKCU:\SOFTWARE\Policies\Microsoft\Office\$ver\OSM" EnableUpload 0 "OSM upload"
    Set-Reg "HKCU:\SOFTWARE\Microsoft\Office\$ver\Common\ClientTelemetry" DisableTelemetry 1 "Client telemetry"
    Set-Reg "HKCU:\SOFTWARE\Microsoft\Office\$ver\Common\ClientTelemetry" VerboseLogging 0 "Verbose logging"
    Set-Reg "HKCU:\SOFTWARE\Microsoft\Office\$ver\Common" QMEnable 0 "CEIP"
    Set-Reg "HKCU:\SOFTWARE\Microsoft\Office\$ver\Common\Feedback" Enabled 0 "Feedback"

    if ($ver -ge '16.0') {
        Set-Reg "HKCU:\SOFTWARE\Microsoft\Office\$ver\Common\Privacy" UserContentDisabled 2 "User content"
        Set-Reg "HKCU:\SOFTWARE\Microsoft\Office\$ver\Common\Privacy" DownloadContentDisabled 2 "Download content"
        Set-Reg "HKCU:\SOFTWARE\Microsoft\Office\$ver\Common\Internet" UseOnlineContent 0 "Online content"
    }
}
#endregion

#region ─── Scheduled Tasks ───────────────────────────────────────────────────────
$tasks = @(
    'Microsoft\Office\OfficeTelemetryAgentFallBack',
    'Microsoft\Office\OfficeTelemetryAgentLogOn',
    'Microsoft\Office\OfficeTelemetryAgentFallBack2016',
    'Microsoft\Office\OfficeTelemetryAgentLogOn2016',
    'Microsoft\Office\OfficeBackgroundTaskHandlerRegistration',
    'Microsoft\Office\OfficeBackgroundTaskHandlerLogon',
    'Microsoft\Office\Office Automatic Updates',
    'Microsoft\Office\Office Automatic Updates 2.0',
    'Microsoft\Office\Office Feature Updates',
    'Microsoft\Office\Office Feature Updates Logon',
    'Microsoft\Office\Office 15 Subscription Heartbeat',
    'Microsoft\Office\Office 16 Subscription Heartbeat',
    'Microsoft\Office\Office Subscription Maintenance',
    'Microsoft\Office\Office ClickToRun Service Monitor'
)
function Disable-Task {
    param($Name)
    try {
        if (Get-Command Get-ScheduledTask -EA SilentlyContinue) {
            $t = Get-ScheduledTask -TaskName $Name -EA SilentlyContinue
            if ($t -and $t.State -ne 'Disabled') { Disable-ScheduledTask -TaskName $Name | Out-Null }
            Write-O "Disabled $Name"
        } else {
            & schtasks.exe /Change /TN $Name /DISABLE >$null
            if ($LASTEXITCODE -eq 0) { Write-O "Disabled $Name" } else { Write-N $Name }
        }
    } catch { Write-N $Name }
}
Write-S "Disabling scheduled telemetry tasks"
$tasks | ForEach-Object { Disable-Task $_ }
#endregion

#region ─── HOSTS File Block ──────────────────────────────────────────────────────
if ($BlockHosts -or ($PSCmdlet.ShouldContinue(
    'Block Microsoft telemetry domains via hosts file?',
    'Hosts file protection'))) {

    $hostsPath = "$env:SystemRoot\System32\drivers\etc\hosts"

    # ─── Windows Defender Exclusion ───────────────────────────────────────────────
    try {
        if (Get-Command "Add-MpPreference" -ErrorAction SilentlyContinue) {
            Add-MpPreference -ExclusionPath $hostsPath -ErrorAction Stop
            Write-O "Hosts file added to Windows Defender exclusions"
        }
    } catch {
        Write-W "Could not add hosts file to Defender exclusions (not critical)"
    }

    $domains   = @(
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

    $backup = "$hostsPath.backup.$(Get-Date -Format 'yyyyMMddHHmmss')"
    Copy-Item $hostsPath $backup
    Write-O "Hosts backed up to $backup"

    $hosts = Get-Content $hostsPath -Raw
    $new   = $domains | Where-Object { $hosts -notmatch [regex]::Escape($_) } |
             ForEach-Object { "0.0.0.0 $_" }

    if ($new) {
        $new = "`n# Office-Telemetry Block - $(Get-Date)`n" + ($new -join "`n") + "`n# End of Office Telemetry Hosts`n"
        Add-Content $hostsPath $new
        Write-O "Added $($new.Count) new blocked domains"
    } else {
        Write-N "All domains already blocked"
    }

    & "$env:SystemRoot\System32\ipconfig.exe" /flushdns | Out-Null
}
#endregion

#region ─── Roll-back Helper ───────────────────────────────────────────────────────
function Undo-OfficePrivacyChanges {
    foreach ($r in $Rollback) {
        Set-ItemProperty -Path $r.Path -Name $r.Name -Value $r.Old -Force
    }
    Write-O "Registry rollback complete"
}
#endregion

#region ─── Summary ────────────────────────────────────────────────────────────────
Write-S "Summary"
$installed | ForEach-Object { Write-C "  [OK] $($OfficeMap[$_])" Found }
Write-C "`nRestart Office applications for changes to take effect." Warning
#endregion

Write-C "`nScript completed successfully!" Success
if (-not $WhatIf) { Read-Host "Press Enter to exit" }
