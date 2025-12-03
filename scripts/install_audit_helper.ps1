$ErrorActionPreference = "Stop"

# Paths
$scriptDir = Split-Path -Parent $MyInvocation.MyCommand.Path
$repoDir   = Split-Path -Parent $scriptDir  # Go up one level to root
$xlStart   = Join-Path $env:APPDATA "Microsoft\Excel\XLSTART"

# Auto-detect XLAM file in repo root
$addinFiles = Get-ChildItem -Path $repoDir -Filter "*.xlam" -File
if ($addinFiles.Count -eq 0) {
    Write-Error "Cannot find any .xlam file in: $repoDir"
} elseif ($addinFiles.Count -gt 1) {
    Write-Host "Multiple .xlam files found. Please select:" -ForegroundColor Yellow
    for ($i = 0; $i -lt $addinFiles.Count; $i++) {
        Write-Host "  [$($i+1)] $($addinFiles[$i].Name)"
    }
    $selection = Read-Host "Enter number (1-$($addinFiles.Count))"
    $addinFile = $addinFiles[[int]$selection - 1]
} else {
    $addinFile = $addinFiles[0]
}

$source = $addinFile.FullName
$addinName = $addinFile.Name
$target = Join-Path $xlStart $addinName

Write-Host "Installing: $addinName" -ForegroundColor Cyan

if (-not (Test-Path $xlStart)) {
    New-Item -ItemType Directory -Path $xlStart | Out-Null
}

# Copy add-in to XLSTART so Excel auto-loads it
Copy-Item -Path $source -Destination $xlStart -Force

# Remove Zone.Identifier if present
$zoneFile = $target + ":Zone.Identifier"
if (Test-Path $zoneFile) { Remove-Item $zoneFile -Force }

Write-Host "Installed successfully at: $xlStart" -ForegroundColor Green

# Setup auto-update (mandatory for keeping tool up-to-date)
Write-Host "`nSetting up automatic updates..." -ForegroundColor Cyan

try {
    $TaskName = "GAFC Audit Helper Auto Update"
    $UpdateScript = Join-Path $scriptDir "update_audit_helper.ps1"
    $UpdateInterval = 12  # Hours

    # Remove existing task if present
    $existingTask = Get-ScheduledTask -TaskName $TaskName -ErrorAction SilentlyContinue
    if ($existingTask) {
        Unregister-ScheduledTask -TaskName $TaskName -Confirm:$false
    }

    # Create scheduled task
    $action = New-ScheduledTaskAction -Execute "powershell.exe" -Argument "-NoProfile -WindowStyle Hidden -ExecutionPolicy Bypass -File `"$UpdateScript`""
    $trigger = New-ScheduledTaskTrigger -Once -At (Get-Date).AddMinutes(5) -RepetitionInterval (New-TimeSpan -Hours $UpdateInterval)
    $settings = New-ScheduledTaskSettingsSet -AllowStartIfOnBatteries -DontStopIfGoingOnBatteries -StartWhenAvailable -RunOnlyIfNetworkAvailable -ExecutionTimeLimit (New-TimeSpan -Minutes 10)
    $principal = New-ScheduledTaskPrincipal -UserId $env:USERNAME -LogonType S4U

    Register-ScheduledTask -TaskName $TaskName -Action $action -Trigger $trigger -Settings $settings -Principal $principal -Description "Automatically checks for and installs GAFC Audit Helper updates when Excel is not running" | Out-Null

    Write-Host "Auto-update enabled! Updates will check every $UpdateInterval hours." -ForegroundColor Green
} catch {
    Write-Host "Warning: Failed to setup auto-update: $_" -ForegroundColor Yellow
}

Write-Host "`nOpening Excel to activate license..." -ForegroundColor Cyan

# Open Excel with the add-in loaded
Start-Process "excel.exe"

Write-Host "`nDone! Please enter your license key when prompted." -ForegroundColor Green
Write-Host "For manual updates, run: update_audit_helper.ps1" -ForegroundColor Gray
exit 0
