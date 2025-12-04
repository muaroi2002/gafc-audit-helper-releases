$ErrorActionPreference = "Stop"

$xlStart   = Join-Path $env:APPDATA "Microsoft\Excel\XLSTART"
$addinName = "gafc_audit_helper.xlam"
$target    = Join-Path $xlStart $addinName

if (Test-Path $target) {
    # Close Excel if running to avoid file lock
    $excelProcesses = Get-Process -Name "EXCEL" -ErrorAction SilentlyContinue
    if ($excelProcesses) {
        Write-Host "Excel is running. Closing Excel..." -ForegroundColor Yellow
        $excelProcesses | ForEach-Object { $_.CloseMainWindow() | Out-Null }
        Start-Sleep -Seconds 2
        # Force kill if still running
        Stop-Process -Name "EXCEL" -Force -ErrorAction SilentlyContinue
        Start-Sleep -Seconds 1
    }

    Remove-Item $target -Force
    Write-Host "Add-in removed successfully from: $target" -ForegroundColor Green
} else {
    Write-Host "Add-in not found in XLSTART, nothing to uninstall." -ForegroundColor Yellow
}
exit 0
