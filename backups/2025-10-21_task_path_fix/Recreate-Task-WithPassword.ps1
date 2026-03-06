# Script to recreate the Depot Script Daily Fixed task with password
# This script must be run as Administrator

Write-Host "========================================" -ForegroundColor Cyan
Write-Host "Recreating Task: Depot Script Daily Fixed" -ForegroundColor Cyan
Write-Host "========================================" -ForegroundColor Cyan
Write-Host ""

# Prompt for the Service account password
$username = "WIN-H7BKO5H0RMC\Service"
Write-Host "Please enter the password for user: $username" -ForegroundColor Yellow
$password = Read-Host "Password" -AsSecureString

# Convert to plain text for schtasks (unfortunately required)
$BSTR = [System.Runtime.InteropServices.Marshal]::SecureStringToBSTR($password)
$PlainPassword = [System.Runtime.InteropServices.Marshal]::PtrToStringAuto($BSTR)

# Create the task using schtasks with XML import
$xmlPath = "D:\Dataserver\_Batchprozesse\depot\backups\2025-10-21_task_path_fix\task_fixed.xml"

Write-Host ""
Write-Host "Creating task from XML: $xmlPath" -ForegroundColor Green

$result = schtasks /Create /TN "\AHSkripts\Depot Script Daily Fixed" /XML $xmlPath /RU $username /RP $PlainPassword 2>&1

if ($LASTEXITCODE -eq 0) {
    Write-Host ""
    Write-Host "SUCCESS: Task created successfully!" -ForegroundColor Green
    Write-Host ""

    # Verify the task
    Write-Host "Verifying task configuration..." -ForegroundColor Cyan
    $task = Get-ScheduledTask -TaskName "Depot Script Daily Fixed" -TaskPath "\AHSkripts\" -ErrorAction SilentlyContinue

    if ($task) {
        Write-Host "  Task Name: $($task.TaskName)" -ForegroundColor White
        Write-Host "  State: $($task.State)" -ForegroundColor White
        Write-Host "  LogonType: $($task.Principal.LogonType)" -ForegroundColor White
        Write-Host "  RunLevel: $($task.Principal.RunLevel)" -ForegroundColor White
        Write-Host "  Action: $($task.Actions[0].Execute) $($task.Actions[0].Arguments)" -ForegroundColor White

        # Get schedule info
        $info = Get-ScheduledTaskInfo -TaskName "Depot Script Daily Fixed" -TaskPath "\AHSkripts\"
        Write-Host "  Next Run: $($info.NextRunTime)" -ForegroundColor White
        Write-Host ""
        Write-Host "Task is ready to run!" -ForegroundColor Green
    } else {
        Write-Host "WARNING: Could not verify task" -ForegroundColor Yellow
    }
} else {
    Write-Host ""
    Write-Host "ERROR: Failed to create task!" -ForegroundColor Red
    Write-Host "Error output:" -ForegroundColor Red
    Write-Host $result -ForegroundColor Red
    exit 1
}

# Clean up password from memory
$PlainPassword = $null
[System.GC]::Collect()

Write-Host ""
Write-Host "========================================" -ForegroundColor Cyan
Write-Host "Task recreation completed!" -ForegroundColor Cyan
Write-Host "========================================" -ForegroundColor Cyan
