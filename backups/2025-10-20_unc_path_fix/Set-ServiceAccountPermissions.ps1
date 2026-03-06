# PowerShell Script to Grant Service Account Access to Finance Folders
# Run this script as Administrator on WIN-H7BKO5H0RMC
# Date: 2025-10-20

#Requires -RunAsAdministrator

$ServiceAccount = "WIN-H7BKO5H0RMC\service"
$ShareName = "Dataserver"
$SharePath = "D:\Dataserver"
$FinanceInputPath = "D:\Dataserver\Dummy\Finance_Input"
$FinanceOutputPath = "D:\Dataserver\Dummy\Finance_Output"

Write-Host "========================================" -ForegroundColor Cyan
Write-Host "Setting Permissions for Service Account" -ForegroundColor Cyan
Write-Host "========================================" -ForegroundColor Cyan
Write-Host ""
Write-Host "Service Account: $ServiceAccount" -ForegroundColor Yellow
Write-Host "Share: $ShareName ($SharePath)" -ForegroundColor Yellow
Write-Host ""

function Add-NTFSPermission {
    param(
        [string]$Path,
        [string]$Account,
        [string]$Rights,
        [string]$InheritanceFlags = "ContainerInherit,ObjectInherit",
        [string]$PropagationFlags = "None"
    )

    try {
        Write-Host "Setting NTFS permissions on: $Path" -ForegroundColor White
        Write-Host "  Account: $Account" -ForegroundColor Gray
        Write-Host "  Rights: $Rights" -ForegroundColor Gray

        $acl = Get-Acl -Path $Path
        $accessRule = New-Object System.Security.AccessControl.FileSystemAccessRule($Account, $Rights, $InheritanceFlags, $PropagationFlags, "Allow")
        $acl.SetAccessRule($accessRule)
        Set-Acl -Path $Path -AclObject $acl

        Write-Host "  Success" -ForegroundColor Green
        Write-Host ""
        return $true
    }
    catch {
        Write-Host "  Error: $($_.Exception.Message)" -ForegroundColor Red
        Write-Host ""
        return $false
    }
}

function Add-SharePermission {
    param(
        [string]$ShareName,
        [string]$Account,
        [string]$AccessRight
    )

    try {
        Write-Host "Setting share permissions on: $ShareName" -ForegroundColor White
        Write-Host "  Account: $Account" -ForegroundColor Gray
        Write-Host "  Rights: $AccessRight" -ForegroundColor Gray

        Grant-SmbShareAccess -Name $ShareName -AccountName $Account -AccessRight $AccessRight -Force -ErrorAction Stop

        Write-Host "  Success" -ForegroundColor Green
        Write-Host ""
        return $true
    }
    catch {
        Write-Host "  Error: $($_.Exception.Message)" -ForegroundColor Red
        Write-Host ""
        return $false
    }
}

Write-Host "Step 1: Grant Share-Level Read Access to Dataserver" -ForegroundColor Cyan
Write-Host "------------------------------------------------------" -ForegroundColor Cyan
$shareResult = Add-SharePermission -ShareName $ShareName -Account $ServiceAccount -AccessRight "Read"

Write-Host "Step 2: Grant NTFS Read Access to Finance_Input" -ForegroundColor Cyan
Write-Host "------------------------------------------------" -ForegroundColor Cyan
$inputResult = Add-NTFSPermission -Path $FinanceInputPath -Account $ServiceAccount -Rights "Read"

Write-Host "Step 3: Grant NTFS Modify Access to Finance_Output" -ForegroundColor Cyan
Write-Host "---------------------------------------------------" -ForegroundColor Cyan
$outputResult = Add-NTFSPermission -Path $FinanceOutputPath -Account $ServiceAccount -Rights "Modify"

Write-Host ""
Write-Host "========================================" -ForegroundColor Cyan
Write-Host "Summary" -ForegroundColor Cyan
Write-Host "========================================" -ForegroundColor Cyan
Write-Host ""

if ($shareResult) {
    Write-Host "Share permissions set successfully" -ForegroundColor Green
} else {
    Write-Host "Share permissions failed" -ForegroundColor Red
}

if ($inputResult) {
    Write-Host "Finance_Input NTFS permissions set successfully" -ForegroundColor Green
} else {
    Write-Host "Finance_Input NTFS permissions failed" -ForegroundColor Red
}

if ($outputResult) {
    Write-Host "Finance_Output NTFS permissions set successfully" -ForegroundColor Green
} else {
    Write-Host "Finance_Output NTFS permissions failed" -ForegroundColor Red
}

Write-Host ""

if ($shareResult -and $inputResult -and $outputResult) {
    Write-Host "All permissions set successfully!" -ForegroundColor Green
    Write-Host ""
    Write-Host "The service account now has access to the required folders." -ForegroundColor Yellow
    Write-Host ""
    Write-Host "Test with these commands (as service user):" -ForegroundColor Yellow
    Write-Host "  dir \\WIN-H7BKO5H0RMC\Dataserver\Dummy\Finance_Input" -ForegroundColor White
    Write-Host "  dir \\WIN-H7BKO5H0RMC\Dataserver\Dummy\Finance_Output" -ForegroundColor White
    exit 0
} else {
    Write-Host "Some permissions failed. Review errors above." -ForegroundColor Red
    exit 1
}
