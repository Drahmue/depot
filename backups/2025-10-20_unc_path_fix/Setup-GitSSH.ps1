# PowerShell Script to Setup SSH Keys for Git Access
# Run this script as the service user
# Date: 2025-10-20

Write-Host "========================================" -ForegroundColor Cyan
Write-Host "Setting Up SSH Keys for Git/GitHub" -ForegroundColor Cyan
Write-Host "========================================" -ForegroundColor Cyan
Write-Host ""

# Check if running as service user
$currentUser = $env:USERNAME
Write-Host "Current user: $currentUser" -ForegroundColor Yellow

if ($currentUser -ne "service") {
    Write-Host "WARNING: This should be run as the service user!" -ForegroundColor Red
    Write-Host "Current user is: $currentUser" -ForegroundColor Red
    $continue = Read-Host "Continue anyway? (y/n)"
    if ($continue -ne "y") {
        exit 1
    }
}

Write-Host ""

# Define SSH directory
$sshDir = "$env:USERPROFILE\.ssh"
$keyPath = "$sshDir\id_ed25519"
$pubKeyPath = "$keyPath.pub"

# Create .ssh directory if it doesn't exist
if (-not (Test-Path $sshDir)) {
    Write-Host "Creating .ssh directory: $sshDir" -ForegroundColor White
    New-Item -ItemType Directory -Path $sshDir -Force | Out-Null
    Write-Host "  Created" -ForegroundColor Green
} else {
    Write-Host ".ssh directory exists: $sshDir" -ForegroundColor Gray
}

Write-Host ""

# Check if key already exists
if (Test-Path $keyPath) {
    Write-Host "SSH key already exists at: $keyPath" -ForegroundColor Yellow
    Write-Host ""
    Write-Host "Options:" -ForegroundColor Yellow
    Write-Host "  1. Use existing key (show public key to add to GitHub)" -ForegroundColor White
    Write-Host "  2. Generate new key (will overwrite existing)" -ForegroundColor White
    Write-Host "  3. Exit" -ForegroundColor White
    $choice = Read-Host "Choose (1/2/3)"

    if ($choice -eq "3") {
        exit 0
    } elseif ($choice -eq "2") {
        Write-Host ""
        Write-Host "Generating new SSH key..." -ForegroundColor White
        ssh-keygen -t ed25519 -C "service@WIN-H7BKO5H0RMC" -f $keyPath -N '""'
        Write-Host "  New key generated" -ForegroundColor Green
    }
} else {
    Write-Host "Generating SSH key..." -ForegroundColor White
    Write-Host "  Type: ed25519" -ForegroundColor Gray
    Write-Host "  Path: $keyPath" -ForegroundColor Gray
    Write-Host ""

    # Generate key with empty passphrase (for automated scripts)
    ssh-keygen -t ed25519 -C "service@WIN-H7BKO5H0RMC" -f $keyPath -N '""'

    if ($LASTEXITCODE -eq 0) {
        Write-Host "  Key generated successfully" -ForegroundColor Green
    } else {
        Write-Host "  ERROR: Failed to generate key" -ForegroundColor Red
        exit 1
    }
}

Write-Host ""
Write-Host "========================================" -ForegroundColor Cyan
Write-Host "Public Key (Add this to GitHub)" -ForegroundColor Cyan
Write-Host "========================================" -ForegroundColor Cyan
Write-Host ""

if (Test-Path $pubKeyPath) {
    $publicKey = Get-Content $pubKeyPath -Raw
    Write-Host $publicKey -ForegroundColor Yellow

    # Copy to clipboard if possible
    try {
        Set-Clipboard -Value $publicKey
        Write-Host ""
        Write-Host "Public key copied to clipboard!" -ForegroundColor Green
    } catch {
        Write-Host ""
        Write-Host "Could not copy to clipboard. Please copy manually from above." -ForegroundColor Yellow
    }
} else {
    Write-Host "ERROR: Public key file not found at $pubKeyPath" -ForegroundColor Red
    exit 1
}

Write-Host ""
Write-Host "========================================" -ForegroundColor Cyan
Write-Host "Next Steps" -ForegroundColor Cyan
Write-Host "========================================" -ForegroundColor Cyan
Write-Host ""
Write-Host "1. Go to GitHub repository settings" -ForegroundColor White
Write-Host "   URL: https://github.com/YOUR_USERNAME/YOUR_REPO/settings/keys" -ForegroundColor Gray
Write-Host ""
Write-Host "2. Click 'Add deploy key' or go to your GitHub account settings" -ForegroundColor White
Write-Host "   Account SSH keys: https://github.com/settings/keys" -ForegroundColor Gray
Write-Host ""
Write-Host "3. Paste the public key shown above" -ForegroundColor White
Write-Host ""
Write-Host "4. Give it a title like: 'WIN-H7BKO5H0RMC service account'" -ForegroundColor White
Write-Host ""
Write-Host "5. For Deploy Key: Check 'Allow write access' if needed" -ForegroundColor White
Write-Host ""
Write-Host "6. Test the connection by running:" -ForegroundColor White
Write-Host "   ssh -T git@github.com" -ForegroundColor Cyan
Write-Host ""
Write-Host "7. If the repo uses HTTPS, change it to SSH:" -ForegroundColor White
Write-Host "   cd \\WIN-H7BKO5H0RMC\_Batchprozesse\depot" -ForegroundColor Cyan
Write-Host "   git remote set-url origin git@github.com:USERNAME/REPO.git" -ForegroundColor Cyan
Write-Host ""
