# Setup-TokenRefresh.ps1
# Run this ONCE on EACH TV machine. NO ADMIN NEEDED.
#   1. Creates  %USERPROFILE%\PowerPointTokenRefresh\
#   2. Copies   RestartPowerPointSlideshow.ps1 there
#   3. Registers the "PowerPoint Token Refresh" task (every 50 min, runs as you)
# Re-running it is safe -- it just overwrites the existing task.
#
# HOW TO RUN (no admin required):
#   right-click this file > "Run with PowerShell"
#   -- or from a normal (non-admin) PowerShell window:
#   powershell -ExecutionPolicy Bypass -File ".\Setup-TokenRefresh.ps1"

$ErrorActionPreference = 'Stop'

$taskName   = 'PowerPoint Token Refresh'
$installDir = Join-Path $env:USERPROFILE 'PowerPointTokenRefresh'
$source     = Join-Path $PSScriptRoot 'RestartPowerPointSlideshow.ps1'
$dest       = Join-Path $installDir   'RestartPowerPointSlideshow.ps1'

Write-Host "Setting up '$taskName' on $env:COMPUTERNAME as $env:USERDOMAIN\$env:USERNAME ..."

# 1. Create the install folder under your profile (no admin needed)
if (-not (Test-Path -LiteralPath $installDir)) {
    New-Item -ItemType Directory -Path $installDir | Out-Null
    Write-Host "Created $installDir"
}

# 2. Copy the restart script there
if (-not (Test-Path -LiteralPath $source)) {
    throw "Can't find RestartPowerPointSlideshow.ps1 next to this setup script: $source"
}
Copy-Item -LiteralPath $source -Destination $dest -Force
Write-Host "Copied restart script to $dest"

# 3. Register the task -- runs as the current user at limited rights, so no admin needed
$action = New-ScheduledTaskAction -Execute 'powershell.exe' `
    -Argument ('-ExecutionPolicy Bypass -WindowStyle Hidden -File "{0}"' -f $dest)

$trigger = New-ScheduledTaskTrigger -Once -At (Get-Date).AddMinutes(2) `
    -RepetitionInterval (New-TimeSpan -Minutes 50)

$principal = New-ScheduledTaskPrincipal -UserId "$env:USERDOMAIN\$env:USERNAME" `
    -LogonType Interactive -RunLevel Limited

$settings = New-ScheduledTaskSettingsSet -MultipleInstances IgnoreNew `
    -AllowStartIfOnBatteries -DontStopIfGoingOnBatteries -StartWhenAvailable

try {
    Register-ScheduledTask -TaskName $taskName -Action $action -Trigger $trigger `
        -Principal $principal -Settings $settings -Force | Out-Null
    Write-Host ""
    Write-Host "DONE. '$taskName' is registered; it restarts the slideshow every 50 minutes."
    Write-Host "Verify:  Get-ScheduledTask -TaskName '$taskName' | Get-ScheduledTaskInfo"
    Write-Host "Logs:    Get-Content `"$dest`".Replace('RestartPowerPointSlideshow.ps1','TokenRefresh.log') -Tail 20 -Wait"
} catch {
    Write-Warning "Could not register the task: $($_.Exception.Message)"
    Write-Warning "If this says Access Denied, your machine may block task creation by policy."
    Write-Warning "Tell me and I'll give you a no-task fallback (a logon startup loop instead)."
    exit 1
}
