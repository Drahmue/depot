# PowerShell Script to Upgrade Dataserver Share Permission
# Run this script as Administrator on WIN-H7BKO5H0RMC
# Date: 2025-10-20
# Purpose: Change share permission from Read to Change so NTFS Modify can work

#Requires -RunAsAdministrator

$ServiceAccount = "WIN-H7BKO5H0RMC\service"
$ShareName = "Dataserver"

Write-Host "========================================" -ForegroundColor Cyan
Write-Host "Upgrading Share Permission" -ForegroundColor Cyan
Write-Host "========================================" -ForegroundColor Cyan
Write-Host ""
Write-Host "Service Account: $ServiceAccount" -ForegroundColor Yellow
Write-Host "Share: $ShareName" -ForegroundColor Yellow
Write-Host ""

try {
    Write-Host "Step 1: Removing old Read permission..." -ForegroundColor White
    try {
        Revoke-SmbShareAccess -Name $ShareName -AccountName $ServiceAccount -Force -ErrorAction SilentlyContinue
        Write-Host "  Old permission removed" -ForegroundColor Gray
    } catch {
        Write-Host "  No existing permission found (OK)" -ForegroundColor Gray
    }

    Write-Host ""
    Write-Host "Step 2: Granting Change permission..." -ForegroundColor White
    Grant-SmbShareAccess -Name $ShareName -AccountName $ServiceAccount -AccessRight Change -Force -ErrorAction Stop
    Write-Host "  Change permission granted" -ForegroundColor Green

    Write-Host ""
    Write-Host "SUCCESS!" -ForegroundColor Green
    Write-Host ""
    Write-Host "The service account now has:" -ForegroundColor Yellow
    Write-Host "  - Share-level: Change (Read + Write)" -ForegroundColor White
    Write-Host "  - NTFS on Finance_Input: Modify" -ForegroundColor White
    Write-Host "  - NTFS on Finance_Output: Modify" -ForegroundColor White
    Write-Host ""
    Write-Host "This allows the depot script to:" -ForegroundColor Yellow
    Write-Host "  - Open files in r+w mode for file lock checks" -ForegroundColor White
    Write-Host "  - Read input files" -ForegroundColor White
    Write-Host "  - Write output files" -ForegroundColor White
    Write-Host ""

    # Verify
    Write-Host "Verifying share permissions..." -ForegroundColor Cyan
    $shareAccess = Get-SmbShareAccess -Name $ShareName | Where-Object { $_.AccountName -eq $ServiceAccount }
    if ($shareAccess) {
        Write-Host "  Account: $($shareAccess.AccountName)" -ForegroundColor Gray
        Write-Host "  Access: $($shareAccess.AccessRight)" -ForegroundColor Gray
        Write-Host "  Type: $($shareAccess.AccessControlType)" -ForegroundColor Gray
    }

    exit 0
}
catch {
    Write-Host "ERROR: $($_.Exception.Message)" -ForegroundColor Red
    exit 1
}
