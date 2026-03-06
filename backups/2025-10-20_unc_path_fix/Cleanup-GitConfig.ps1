# PowerShell Script to Clean Up Git Safe Directory Configuration
# Run this script as the service user
# Date: 2025-10-20

Write-Host "========================================" -ForegroundColor Cyan
Write-Host "Cleaning Up Git Safe Directory Config" -ForegroundColor Cyan
Write-Host "========================================" -ForegroundColor Cyan
Write-Host ""

# Show current entries
Write-Host "Current safe.directory entries:" -ForegroundColor Yellow
git config --global --get-all safe.directory | ForEach-Object {
    Write-Host "  $_" -ForegroundColor Gray
}
Write-Host ""

# Remove all safe.directory entries
Write-Host "Removing all safe.directory entries..." -ForegroundColor White
git config --global --unset-all safe.directory

Write-Host "All entries removed." -ForegroundColor Green
Write-Host ""

# Add only the correct one
Write-Host "Adding correct safe.directory entry..." -ForegroundColor White
$correctPath = "//WIN-H7BKO5H0RMC/_Batchprozesse/depot"
git config --global --add safe.directory $correctPath

Write-Host "Added: $correctPath" -ForegroundColor Green
Write-Host ""

# Verify
Write-Host "Verification - Current safe.directory entries:" -ForegroundColor Yellow
$entries = git config --global --get-all safe.directory
if ($entries) {
    $entries | ForEach-Object {
        Write-Host "  $_" -ForegroundColor Gray
    }
} else {
    Write-Host "  (none)" -ForegroundColor Gray
}

Write-Host ""
Write-Host "SUCCESS! Git configuration cleaned up." -ForegroundColor Green
Write-Host ""
Write-Host "Next time the depot script runs, the warnings should be gone." -ForegroundColor Yellow
