# Check for Administrative privileges
if (-not ([Security.Principal.WindowsPrincipal][Security.Principal.WindowsIdentity]::GetCurrent()).IsInRole([Security.Principal.WindowsBuiltinRole]::Administrator)) {
    Write-Host "ERROR: Please run this script as an Administrator." -ForegroundColor Red
    break
}

Write-Host "--- Windows System Health Check & Repair ---" -ForegroundColor Cyan

# 1. DISM ScanHealth (The Check)
Write-Host "`n[1/3] Scanning for component store corruption using DISM..." -ForegroundColor Yellow
$scanOutput = dism.exe /Online /Cleanup-Image /ScanHealth 2>&1

if ($scanOutput -match "The component store is repairable" -or $scanOutput -match "corruption was detected") {
    Write-Host "(!) Corruption detected. Starting Repair..." -ForegroundColor Red
    
    # 2. DISM RestoreHealth (The Fix)
    Write-Host "[2/3] Running DISM RestoreHealth..." -ForegroundColor Yellow
    $repairOutput = dism.exe /Online /Cleanup-Image /RestoreHealth 2>&1
    
    if ($LASTEXITCODE -ne 0) {
        Write-Host "ERROR: DISM repair failed. Exit Code: $LASTEXITCODE" -ForegroundColor Red
    } else {
        Write-Host "SUCCESS: DISM repair completed." -ForegroundColor Green
    }
} else {
    Write-Host "SUCCESS: No component store corruption detected. Skipping DISM repair." -ForegroundColor Green
}

# 3. SFC Scannow (System File Verification)
Write-Host "`n[3/3] Running SFC Scannow..." -ForegroundColor Yellow
$sfcOutput = sfc.exe /scannow 2>&1

if ($sfcOutput -match "found corrupt files but was unable to fix some of them") {
    Write-Host "ERROR: SFC found corruption it could NOT repair." -ForegroundColor Red
}
elseif ($sfcOutput -match "Windows Resource Protection could not perform the requested operation") {
    Write-Host "ERROR: SFC failed to execute." -ForegroundColor Red
}
elseif ($sfcOutput -match "found corrupt files and successfully repaired them") {
    Write-Host "SUCCESS: SFC found and repaired system files." -ForegroundColor Green
}
else {
    Write-Host "SUCCESS: No system file integrity violations found." -ForegroundColor Green
}

Write-Host "`n--- Process Completed ---" -ForegroundColor Cyan
