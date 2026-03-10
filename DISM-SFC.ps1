# Ensure script is running as Administrator
if (-not ([Security.Principal.WindowsPrincipal][Security.Principal.WindowsIdentity]::GetCurrent()).IsInRole([Security.Principal.WindowsBuiltinRole]::Administrator)) {
    Write-Host "ERROR: Please run this script as an Administrator." -ForegroundColor Red
    return
}

Write-Host "Starting System Repair..." -ForegroundColor Cyan

# 1. Run DISM RestoreHealth
Write-Host "`n[1/2] Running DISM RestoreHealth..." -ForegroundColor Yellow
$dismOutput = dism.exe /Online /Cleanup-Image /RestoreHealth 2>&1

if ($LASTEXITCODE -ne 0) {
    Write-Host "ERROR: DISM encountered a problem. Exit Code: $LASTEXITCODE" -ForegroundColor Red
    $dismOutput | Select-String "Error:" | Write-Host -ForegroundColor Red
} else {
    Write-Host "SUCCESS: DISM completed successfully." -ForegroundColor Green
}

# 2. Run SFC Scannow
Write-Host "`n[2/2] Running SFC Scannow..." -ForegroundColor Yellow
$sfcOutput = sfc.exe /scannow 2>&1

# Check for specific SFC failure messages
if ($sfcOutput -match "Windows Resource Protection found corrupt files but was unable to fix some of them") {
    Write-Host "ERROR: SFC found corruption it could NOT repair." -ForegroundColor Red
}
elseif ($sfcOutput -match "Windows Resource Protection could not perform the requested operation") {
    Write-Host "ERROR: SFC failed to run." -ForegroundColor Red
}
elseif ($sfcOutput -match "Windows Resource Protection found corrupt files and successfully repaired them") {
    Write-Host "SUCCESS: SFC found and repaired corrupted files." -ForegroundColor Green
}
else {
    Write-Host "SUCCESS: SFC found no integrity violations." -ForegroundColor Green
}

Write-Host "`nRepair process finished." -ForegroundColor Cyan
