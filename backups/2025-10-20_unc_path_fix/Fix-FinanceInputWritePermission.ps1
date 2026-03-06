# PowerShell Script to Add Write Permission to Finance_Input
# Run this script as Administrator on WIN-H7BKO5H0RMC
# Date: 2025-10-20
# Purpose: Grant Write permission to Finance_Input so depot script can check if files are open

#Requires -RunAsAdministrator

$ServiceAccount = "WIN-H7BKO5H0RMC\service"
$FinanceInputPath = "D:\Dataserver\Dummy\Finance_Input"

Write-Host "========================================" -ForegroundColor Cyan
Write-Host "Adding Write Permission to Finance_Input" -ForegroundColor Cyan
Write-Host "========================================" -ForegroundColor Cyan
Write-Host ""
Write-Host "Service Account: $ServiceAccount" -ForegroundColor Yellow
Write-Host "Path: $FinanceInputPath" -ForegroundColor Yellow
Write-Host ""

try {
    Write-Host "Updating NTFS permissions..." -ForegroundColor White

    $acl = Get-Acl -Path $FinanceInputPath

    # Remove old Read-only rule and add ReadAndExecute + Write
    $oldRule = New-Object System.Security.AccessControl.FileSystemAccessRule($ServiceAccount, "Read", "ContainerInherit,ObjectInherit", "None", "Allow")
    $acl.RemoveAccessRule($oldRule) | Out-Null

    # Add new rule with Read, Write, and Execute
    $newRule = New-Object System.Security.AccessControl.FileSystemAccessRule($ServiceAccount, "ReadAndExecute, Write", "ContainerInherit,ObjectInherit", "None", "Allow")
    $acl.AddAccessRule($newRule)

    Set-Acl -Path $FinanceInputPath -AclObject $acl

    Write-Host "Success!" -ForegroundColor Green
    Write-Host ""
    Write-Host "The service account now has:" -ForegroundColor Yellow
    Write-Host "  - Read access to Finance_Input files" -ForegroundColor White
    Write-Host "  - Write access to check if files are open" -ForegroundColor White
    Write-Host "  - Execute access to traverse" -ForegroundColor White
    Write-Host ""
    Write-Host "Note: Service account can now write to Finance_Input." -ForegroundColor Yellow
    Write-Host "This is needed for the file-open check in the depot script." -ForegroundColor Yellow

    exit 0
}
catch {
    Write-Host "Error: $($_.Exception.Message)" -ForegroundColor Red
    exit 1
}
