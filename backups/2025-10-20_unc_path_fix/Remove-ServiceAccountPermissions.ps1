# PowerShell Script to Remove Service Account Access from Finance Folders
# Run this script as Administrator on WIN-H7BKO5H0RMC
# Date: 2025-10-20
# Purpose: Undo script - Remove permissions granted to service account

#Requires -RunAsAdministrator

# Configuration
$ServiceAccount = "WIN-H7BKO5H0RMC\service"
$ShareName = "Dataserver"
$FinanceInputPath = "D:\Dataserver\Dummy\Finance_Input"
$FinanceOutputPath = "D:\Dataserver\Dummy\Finance_Output"

Write-Host "==========================================" -ForegroundColor Cyan
Write-Host "Removing Permissions for Service Account" -ForegroundColor Cyan
Write-Host "==========================================" -ForegroundColor Cyan
Write-Host ""
Write-Host "Service Account: $ServiceAccount" -ForegroundColor Yellow
Write-Host ""

# Function to remove NTFS permissions
function Remove-NTFSPermission {
    param(
        [string]$Path,
        [string]$Account
    )

    try {
        Write-Host "Removing NTFS permissions from: $Path" -ForegroundColor White
        Write-Host "  Account: $Account" -ForegroundColor Gray

        # Get current ACL
        $acl = Get-Acl -Path $Path

        # Find and remove all access rules for the account
        $rules = $acl.Access | Where-Object { $_.IdentityReference -eq $Account }

        if ($rules.Count -eq 0) {
            Write-Host "  ⚠ No permissions found for this account" -ForegroundColor Yellow
            Write-Host ""
            return $true
        }

        foreach ($rule in $rules) {
            $acl.RemoveAccessRule($rule) | Out-Null
        }

        # Apply the modified ACL
        Set-Acl -Path $Path -AclObject $acl

        Write-Host "  ✓ Success - Removed $($rules.Count) permission(s)" -ForegroundColor Green
        Write-Host ""
        return $true
    }
    catch {
        Write-Host "  ✗ Error: $($_.Exception.Message)" -ForegroundColor Red
        Write-Host ""
        return $false
    }
}

# Function to remove share permissions
function Remove-SharePermission {
    param(
        [string]$ShareName,
        [string]$Account
    )

    try {
        Write-Host "Removing share permissions from: $ShareName" -ForegroundColor White
        Write-Host "  Account: $Account" -ForegroundColor Gray

        # Check if permission exists
        $shareAccess = Get-SmbShareAccess -Name $ShareName | Where-Object { $_.AccountName -eq $Account }

        if (-not $shareAccess) {
            Write-Host "  ⚠ No share permissions found for this account" -ForegroundColor Yellow
            Write-Host ""
            return $true
        }

        # Revoke share permission
        Revoke-SmbShareAccess -Name $ShareName -AccountName $Account -Force -ErrorAction Stop

        Write-Host "  ✓ Success" -ForegroundColor Green
        Write-Host ""
        return $true
    }
    catch {
        Write-Host "  ✗ Error: $($_.Exception.Message)" -ForegroundColor Red
        Write-Host ""
        return $false
    }
}

# Main execution
Write-Host "Step 1: Remove Share-Level Access from Dataserver" -ForegroundColor Cyan
Write-Host "--------------------------------------------------" -ForegroundColor Cyan
$shareResult = Remove-SharePermission -ShareName $ShareName -Account $ServiceAccount

Write-Host "Step 2: Remove NTFS Access from Finance_Input" -ForegroundColor Cyan
Write-Host "----------------------------------------------" -ForegroundColor Cyan
$inputResult = Remove-NTFSPermission -Path $FinanceInputPath -Account $ServiceAccount

Write-Host "Step 3: Remove NTFS Access from Finance_Output" -ForegroundColor Cyan
Write-Host "-----------------------------------------------" -ForegroundColor Cyan
$outputResult = Remove-NTFSPermission -Path $FinanceOutputPath -Account $ServiceAccount

# Summary
Write-Host ""
Write-Host "========================================" -ForegroundColor Cyan
Write-Host "Summary" -ForegroundColor Cyan
Write-Host "========================================" -ForegroundColor Cyan
Write-Host ""

if ($shareResult) {
    Write-Host "✓ Share permissions removed successfully" -ForegroundColor Green
} else {
    Write-Host "✗ Share permissions removal failed" -ForegroundColor Red
}

if ($inputResult) {
    Write-Host "✓ Finance_Input NTFS permissions removed successfully" -ForegroundColor Green
} else {
    Write-Host "✗ Finance_Input NTFS permissions removal failed" -ForegroundColor Red
}

if ($outputResult) {
    Write-Host "✓ Finance_Output NTFS permissions removed successfully" -ForegroundColor Green
} else {
    Write-Host "✗ Finance_Output NTFS permissions removal failed" -ForegroundColor Red
}

Write-Host ""

if ($shareResult -and $inputResult -and $outputResult) {
    Write-Host "All permissions removed successfully!" -ForegroundColor Green
    Write-Host ""
    Write-Host "The service account no longer has access to:" -ForegroundColor Yellow
    Write-Host "  - Share: \\WIN-H7BKO5H0RMC\Dataserver" -ForegroundColor White
    Write-Host "  - NTFS: Finance_Input" -ForegroundColor White
    Write-Host "  - NTFS: Finance_Output" -ForegroundColor White
    exit 0
} else {
    Write-Host "Some permissions failed to remove. Please review errors above." -ForegroundColor Red
    exit 1
}
