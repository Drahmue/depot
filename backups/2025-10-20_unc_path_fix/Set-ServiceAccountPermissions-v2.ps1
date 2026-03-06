# PowerShell Script to Grant Service Account Access to Finance Folders
# Run this script as Administrator on WIN-H7BKO5H0RMC
# Date: 2025-10-20
# Version 2: Includes traverse permissions on parent folders

#Requires -RunAsAdministrator

$ServiceAccount = "WIN-H7BKO5H0RMC\service"
$ShareName = "Dataserver"
$DataserverPath = "D:\Dataserver"
$DummyPath = "D:\Dataserver\Dummy"
$FinanceInputPath = "D:\Dataserver\Dummy\Finance_Input"
$FinanceOutputPath = "D:\Dataserver\Dummy\Finance_Output"

Write-Host "========================================" -ForegroundColor Cyan
Write-Host "Setting Permissions for Service Account" -ForegroundColor Cyan
Write-Host "========================================" -ForegroundColor Cyan
Write-Host ""
Write-Host "Service Account: $ServiceAccount" -ForegroundColor Yellow
Write-Host ""

function Add-NTFSPermission {
    param(
        [string]$Path,
        [string]$Account,
        [string]$Rights,
        [string]$InheritanceFlags = "ContainerInherit,ObjectInherit",
        [string]$PropagationFlags = "None",
        [string]$Type = "Allow"
    )

    try {
        Write-Host "Setting NTFS permissions on: $Path" -ForegroundColor White
        Write-Host "  Account: $Account" -ForegroundColor Gray
        Write-Host "  Rights: $Rights" -ForegroundColor Gray
        Write-Host "  Inheritance: $InheritanceFlags" -ForegroundColor Gray

        $acl = Get-Acl -Path $Path
        $accessRule = New-Object System.Security.AccessControl.FileSystemAccessRule($Account, $Rights, $InheritanceFlags, $PropagationFlags, $Type)
        $acl.AddAccessRule($accessRule)
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

Write-Host "Step 2: Grant Traverse on Dataserver Root (no inheritance)" -ForegroundColor Cyan
Write-Host "------------------------------------------------------------" -ForegroundColor Cyan
$dataserverResult = Add-NTFSPermission -Path $DataserverPath -Account $ServiceAccount -Rights "Traverse" -InheritanceFlags "None" -PropagationFlags "None"

Write-Host "Step 3: Grant Traverse on Dummy Folder (no inheritance)" -ForegroundColor Cyan
Write-Host "---------------------------------------------------------" -ForegroundColor Cyan
$dummyResult = Add-NTFSPermission -Path $DummyPath -Account $ServiceAccount -Rights "Traverse" -InheritanceFlags "None" -PropagationFlags "None"

Write-Host "Step 4: Grant Read Access to Finance_Input (with inheritance)" -ForegroundColor Cyan
Write-Host "---------------------------------------------------------------" -ForegroundColor Cyan
$inputResult = Add-NTFSPermission -Path $FinanceInputPath -Account $ServiceAccount -Rights "Read"

Write-Host "Step 5: Grant Modify Access to Finance_Output (with inheritance)" -ForegroundColor Cyan
Write-Host "------------------------------------------------------------------" -ForegroundColor Cyan
$outputResult = Add-NTFSPermission -Path $FinanceOutputPath -Account $ServiceAccount -Rights "Modify"

Write-Host ""
Write-Host "========================================" -ForegroundColor Cyan
Write-Host "Summary" -ForegroundColor Cyan
Write-Host "========================================" -ForegroundColor Cyan
Write-Host ""

$allSuccess = $true

if ($shareResult) {
    Write-Host "Share permissions set successfully" -ForegroundColor Green
} else {
    Write-Host "Share permissions failed" -ForegroundColor Red
    $allSuccess = $false
}

if ($dataserverResult) {
    Write-Host "Dataserver traverse permission set successfully" -ForegroundColor Green
} else {
    Write-Host "Dataserver traverse permission failed" -ForegroundColor Red
    $allSuccess = $false
}

if ($dummyResult) {
    Write-Host "Dummy traverse permission set successfully" -ForegroundColor Green
} else {
    Write-Host "Dummy traverse permission failed" -ForegroundColor Red
    $allSuccess = $false
}

if ($inputResult) {
    Write-Host "Finance_Input read permission set successfully" -ForegroundColor Green
} else {
    Write-Host "Finance_Input read permission failed" -ForegroundColor Red
    $allSuccess = $false
}

if ($outputResult) {
    Write-Host "Finance_Output modify permission set successfully" -ForegroundColor Green
} else {
    Write-Host "Finance_Output modify permission failed" -ForegroundColor Red
    $allSuccess = $false
}

Write-Host ""

if ($allSuccess) {
    Write-Host "All permissions set successfully!" -ForegroundColor Green
    Write-Host ""
    Write-Host "The service account can now:" -ForegroundColor Yellow
    Write-Host "  - Traverse through Dataserver and Dummy folders" -ForegroundColor White
    Write-Host "  - Read from Finance_Input" -ForegroundColor White
    Write-Host "  - Write to Finance_Output" -ForegroundColor White
    Write-Host "  - Cannot access other folders in Dataserver" -ForegroundColor White
    Write-Host ""
    Write-Host "Test with (as service user):" -ForegroundColor Yellow
    Write-Host "  dir \\WIN-H7BKO5H0RMC\Dataserver\Dummy\Finance_Input" -ForegroundColor White
    Write-Host "  dir \\WIN-H7BKO5H0RMC\Dataserver\Dummy\Finance_Output" -ForegroundColor White
    exit 0
} else {
    Write-Host "Some permissions failed. Review errors above." -ForegroundColor Red
    exit 1
}
