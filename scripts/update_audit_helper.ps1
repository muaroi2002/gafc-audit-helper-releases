$ErrorActionPreference = "Stop"

# Configuration
$ManifestUrl = "https://raw.githubusercontent.com/muaroi2002/gafc-audit-helper-releases/main/releases/audit_tool.json"
$AddinName   = "gafc_audit_helper.xlam"  # Fixed filename (version in metadata only)
$SilentMode  = $false  # Set to $true for scheduled task (no output)

# Paths
$XlStart = Join-Path $env:APPDATA "Microsoft\Excel\XLSTART"
$Target  = Join-Path $XlStart $AddinName
$LogFile = Join-Path $env:TEMP "gafc_update.log"

# Helper function for logging
function Write-Log {
    param($Message, $Color = "White")
    $timestamp = Get-Date -Format "yyyy-MM-dd HH:mm:ss"
    $logMsg = "[$timestamp] $Message"
    Add-Content -Path $LogFile -Value $logMsg
    if (-not $SilentMode) {
        Write-Host $Message -ForegroundColor $Color
    }
}

# Check if Excel is running
function Test-ExcelRunning {
    $excelProcesses = Get-Process -Name "EXCEL" -ErrorAction SilentlyContinue
    return ($null -ne $excelProcesses)
}

if (-not (Test-Path $XlStart)) {
    Write-Log "Add-in not installed: XLSTART folder not found ($XlStart). Run install first." "Red"
    exit 1
}

# Exit silently if Excel is running (for scheduled task)
if (Test-ExcelRunning) {
    Write-Log "Excel is running. Skipping update." "Yellow"
    exit 0
}

# Fetch manifest
try {
    $manifest = Invoke-RestMethod -Uri $ManifestUrl -UseBasicParsing
} catch {
    Write-Log "Failed to download manifest: $ManifestUrl. Error: $_" "Red"
    exit 1
}

if (-not $manifest.download_url) {
    Write-Log "Manifest missing download_url" "Red"
    exit 1
}
$downloadUrl = $manifest.download_url
$latest      = $manifest.latest
$shaExpected = $manifest.sha256
$Temp        = Join-Path $env:TEMP "$AddinName.tmp"

Write-Log "Downloading version $latest ..." "Cyan"
try {
    Invoke-WebRequest -Uri $downloadUrl -OutFile $Temp
} catch {
    Write-Log "Failed to download update: $_" "Red"
    exit 1
}

# Verify hash if available
if ($shaExpected) {
    $shaLocal = (Get-FileHash -Path $Temp -Algorithm SHA256).Hash.ToLower()
    if ($shaLocal -ne $shaExpected.ToLower()) {
        Remove-Item $Temp -Force
        Write-Log "SHA256 mismatch. Expected: $shaExpected, Actual: $shaLocal" "Red"
        exit 1
    }
    Write-Log "SHA256 verified successfully." "Green"
}

# Check if target file already exists and has same hash
if (Test-Path $Target -PathType Leaf -ErrorAction SilentlyContinue) {
    if ($shaExpected) {
        $currentSha = (Get-FileHash -Path $Target -Algorithm SHA256).Hash.ToLower()
        if ($currentSha -eq $shaExpected.ToLower()) {
            Write-Log "Already on latest version ($latest). No update needed." "Green"
            Remove-Item $Temp -Force
            exit 0
        }
    }
    # Remove current version (no backup needed - we already verified download)
    Write-Log "Removing current version..." "Cyan"
    Remove-Item $Target -Force
}

# Clean up any old .bak files
$bakFile = $Target + ".bak"
if (Test-Path $bakFile) {
    Write-Log "Removing old backup file..." "Cyan"
    Remove-Item $bakFile -Force -ErrorAction SilentlyContinue
}

# Copy to XLSTART
Write-Log "Installing version $latest ..." "Cyan"
Copy-Item $Temp $Target -Force
Remove-Item $Temp -Force

# Remove Zone.Identifier if present
$zoneFile = $Target + ":Zone.Identifier"
if (Test-Path $zoneFile) { Remove-Item $zoneFile -Force }

Write-Log "GAFC Audit Helper updated successfully to version $latest" "Green"
Write-Log "Location: $Target" "Green"
Write-Log "Next time you open Excel, the new version will be loaded." "Green"
exit 0
