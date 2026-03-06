# PowerShell Script to Grant Modify Permission to Finance_Input
# Run this script as Administrator on WIN-H7BKO5H0RMC
# Date: 2025-10-20

#Requires -RunAsAdministrator

$ServiceAccount = "WIN-H7BKO5H0RMC\service"
$FinanceInputPath = "D:\Dataserver\Dummy\Finance_Input"

Write-Host "Granting Modify permission to Finance_Input..." -ForegroundColor Cyan
Write-Host "Service Account: $ServiceAccount" -ForegroundColor Yellow
Write-Host "Path: $FinanceInputPath" -ForegroundColor Yellow
Write-Host ""

try {
    $acl = Get-Acl -Path $FinanceInputPath

    # Remove existing service account rules
    $acl.Access | Where-Object { $_.IdentityReference -eq $ServiceAccount } | ForEach-Object {
        $acl.RemoveAccessRule($_) | Out-Null
    }

    # Add Modify permission (includes Read, Write, Execute, Delete)
    $rule = New-Object System.Security.AccessControl.FileSystemAccessRule(
        $ServiceAccount,
        "Modify",
        "ContainerInherit,ObjectInherit",
        "None",
        "Allow"
    )

    $acl.AddAccessRule($rule)
    Set-Acl -Path $FinanceInputPath -AclObject $acl

    Write-Host "SUCCESS: Modify permission granted" -ForegroundColor Green
    Write-Host ""
    Write-Host "Service account can now:" -ForegroundColor Yellow
    Write-Host "  - Read files" -ForegroundColor White
    Write-Host "  - Write to files (needed for file lock check)" -ForegroundColor White
    Write-Host "  - Execute/traverse" -ForegroundColor White
    Write-Host ""

    # Verify
    Write-Host "Verifying permissions..." -ForegroundColor Cyan
    $acl = Get-Acl -Path $FinanceInputPath
    $serviceRules = $acl.Access | Where-Object { $_.IdentityReference -eq $ServiceAccount }

    if ($serviceRules) {
        Write-Host "Current permissions for service account:" -ForegroundColor White
        $serviceRules | ForEach-Object {
            Write-Host "  Rights: $($_.FileSystemRights)" -ForegroundColor Gray
            Write-Host "  Type: $($_.AccessControlType)" -ForegroundColor Gray
        }
    }

    exit 0
}
catch {
    Write-Host "ERROR: $($_.Exception.Message)" -ForegroundColor Red
    exit 1
}
