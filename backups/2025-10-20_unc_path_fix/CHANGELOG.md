# Changelog - UNC Path Fix for Service Account Access
**Date:** 2025-10-20
**Change ID:** 2025-10-20_unc_path_fix
**Author:** Claude Code

---

## Problem Statement
The scheduled task running as user "service" was failing silently with exit code 0xFFFD0000. Investigation revealed that the service account had no access to the UNC path `\\WIN-H7BKO5H0RMC\Dataserver\_Batchprozesse\depot` because it lacked permissions on the "Dataserver" share.

### Root Cause
- Service account could not access `\\WIN-H7BKO5H0RMC\Dataserver\...`
- All file operations failed with "Zugriff verweigert" (Access Denied)
- No logs were being written since October 10, 2025
- Task Scheduler showed task running but producing no output

### Solution
Created a new network share `_Batchprozesse` with permissions specifically for the service account, avoiding the need to grant access to the entire Dataserver share.

---

## Changes Made

### 1. File: `start_depot.ps1`

#### Line 7 - Script Directory Path
**OLD:**
```powershell
$scriptDir = "\\WIN-H7BKO5H0RMC\Dataserver\_Batchprozesse\depot"
```

**NEW:**
```powershell
$scriptDir = "\\WIN-H7BKO5H0RMC\_Batchprozesse\depot"
```

**Reason:** Changed to use the new `_Batchprozesse` share instead of going through the `Dataserver` share, allowing the service account to access only what it needs.

#### Line 49 - Git Safe Directory Configuration
**OLD:**
```powershell
& git config --global --add safe.directory '//WIN-H7BKO5H0RMC/Dataserver/_Batchprozesse/depot' 2>&1 | Out-Null
```

**NEW:**
```powershell
& git config --global --add safe.directory '//WIN-H7BKO5H0RMC/_Batchprozesse/depot' 2>&1 | Out-Null
```

**Reason:** Updated git safe.directory path to match the new UNC share path.

---

## Backup Location
All original files are backed up at:
```
D:\Dataserver\_Batchprozesse\depot\backups\2025-10-20_unc_path_fix\
```

Files backed up:
- `start_depot.ps1.backup` - Original PowerShell script

---

## Prerequisites Applied
1. Created network share `_Batchprozesse` for folder `D:\Dataserver\_Batchprozesse`
2. Granted `WIN-H7BKO5H0RMC\Service` account the following permissions:
   - Share Permissions: Change (or Full Control)
   - NTFS Permissions: Modify (inherited to subfolders)
3. Verified permissions were inherited by `depot` and `depot\logs` subfolders

---

## How to UNDO These Changes

### Option 1: Restore from Backup (Recommended)
```powershell
# Navigate to the depot directory
cd D:\Dataserver\_Batchprozesse\depot

# Restore the original file
Copy-Item -Path ".\backups\2025-10-20_unc_path_fix\start_depot.ps1.backup" -Destination ".\start_depot.ps1" -Force

# Verify restoration
Get-Content ".\start_depot.ps1" | Select-String "Dataserver"
```

### Option 2: Manual Reversion
Edit `D:\Dataserver\_Batchprozesse\depot\start_depot.ps1`:

**Line 7** - Change:
```powershell
$scriptDir = "\\WIN-H7BKO5H0RMC\_Batchprozesse\depot"
```
Back to:
```powershell
$scriptDir = "\\WIN-H7BKO5H0RMC\Dataserver\_Batchprozesse\depot"
```

**Line 49** - Change:
```powershell
& git config --global --add safe.directory '//WIN-H7BKO5H0RMC/_Batchprozesse/depot' 2>&1 | Out-Null
```
Back to:
```powershell
& git config --global --add safe.directory '//WIN-H7BKO5H0RMC/Dataserver/_Batchprozesse/depot' 2>&1 | Out-Null
```

### Post-Undo Steps
If you undo these changes, you'll also need to:
1. Grant the service account permissions on the `Dataserver` share (if you want it to work)
2. OR remove the `_Batchprozesse` share and revert to whatever access method was working previously

---

## Testing Instructions

### Test 1: Service Account Access
```cmd
# As service user
dir \\WIN-H7BKO5H0RMC\_Batchprozesse\depot
```
Expected: Directory listing should appear without "Zugriff verweigert"

### Test 2: Manual Script Execution
```cmd
# As service user
powershell -ExecutionPolicy Bypass -File "\\WIN-H7BKO5H0RMC\_Batchprozesse\depot\start_depot.ps1"
```
Expected:
- Script runs without permission errors
- Logs are created in `\\WIN-H7BKO5H0RMC\_Batchprozesse\depot\logs\depot_2025-10.log`
- Exit code 0

### Test 3: Scheduled Task
```powershell
# Run task manually from Task Scheduler
# Then check logs
Get-Content "D:\Dataserver\_Batchprozesse\depot\logs\depot_2025-10.log" -Tail 20
```
Expected: New log entries with current timestamp

---

## Notes
- The `_Batchprozesse` share must remain active for this solution to work
- If the share is removed, the script will fail with the same permission errors
- The service account only has access to `_Batchprozesse` and its subfolders, not the entire Dataserver
- All other functionality remains unchanged

---

## Additional Permission Changes for Finance Folders

After the initial fix, we discovered that the depot script also needs access to input/output files under `\\WIN-H7BKO5H0RMC\Dataserver\Dummy\`. To maintain security while enabling access, we've created PowerShell scripts to grant minimal required permissions.

### Required Access (from depot.ini analysis):
1. **`\\WIN-H7BKO5H0RMC\Dataserver\Dummy\Finance_Input\`** - Read access for:
   - Instrumente.xlsx
   - bookings.xlsx
   - provisions.xlsx

2. **`\\WIN-H7BKO5H0RMC\Dataserver\Dummy\Finance_Output\`** - Write access for:
   - 38 Excel export files

### Permission Scripts Created:

#### Set-ServiceAccountPermissions-v2.ps1 (CURRENT VERSION)
Location: `D:\Dataserver\_Batchprozesse\depot\backups\2025-10-20_unc_path_fix\Set-ServiceAccountPermissions-v2.ps1`

**What it does:**
1. Grants Share-level Read access to Dataserver share
2. Grants NTFS Traverse permission on `D:\Dataserver` (no inheritance)
3. Grants NTFS Traverse permission on `D:\Dataserver\Dummy` (no inheritance)
4. Grants NTFS Read access to `D:\Dataserver\Dummy\Finance_Input` (with inheritance)
5. Grants NTFS Modify access to `D:\Dataserver\Dummy\Finance_Output` (with inheritance)

**Run as Administrator:**
```powershell
cd D:\Dataserver\_Batchprozesse\depot\backups\2025-10-20_unc_path_fix
.\Set-ServiceAccountPermissions-v2.ps1
```

**Detailed Permissions Applied:**

| Location | Permission Type | Rights | Inheritance | Purpose |
|----------|----------------|--------|-------------|---------|
| `\\WIN-H7BKO5H0RMC\Dataserver` (Share) | Share Permission | Read | N/A | Network access to share |
| `D:\Dataserver` | NTFS | Traverse | None (this folder only) | Navigate through without listing |
| `D:\Dataserver\Dummy` | NTFS | Traverse | None (this folder only) | Navigate through without listing |
| `D:\Dataserver\Dummy\Finance_Input` | NTFS | Read | Yes (all subfolders/files) | Read input Excel files |
| `D:\Dataserver\Dummy\Finance_Output` | NTFS | Modify | Yes (all subfolders/files) | Write output Excel files |

#### Remove-ServiceAccountPermissions.ps1 (UNDO SCRIPT)
Location: `D:\Dataserver\_Batchprozesse\depot\backups\2025-10-20_unc_path_fix\Remove-ServiceAccountPermissions.ps1`

**What it does:**
Removes ALL permissions granted to `WIN-H7BKO5H0RMC\service` account from:
- Dataserver share (share permissions)
- Finance_Input folder (NTFS permissions)
- Finance_Output folder (NTFS permissions)

**Note:** This script removes permissions but does NOT remove traverse permissions from parent folders. If you used v2, you'll need to manually remove those.

**Run as Administrator:**
```powershell
cd D:\Dataserver\_Batchprozesse\depot\backups\2025-10-20_unc_path_fix
.\Remove-ServiceAccountPermissions.ps1
```

**Manual Undo for v2 Traverse Permissions:**
If you need to remove the traverse permissions from parent folders:

```powershell
# Remove from Dataserver root
$acl = Get-Acl "D:\Dataserver"
$accessRule = New-Object System.Security.AccessControl.FileSystemAccessRule("WIN-H7BKO5H0RMC\service", "Traverse", "None", "None", "Allow")
$acl.RemoveAccessRule($accessRule)
Set-Acl "D:\Dataserver" -AclObject $acl

# Remove from Dummy folder
$acl = Get-Acl "D:\Dataserver\Dummy"
$accessRule = New-Object System.Security.AccessControl.FileSystemAccessRule("WIN-H7BKO5H0RMC\service", "Traverse", "None", "None", "Allow")
$acl.RemoveAccessRule($accessRule)
Set-Acl "D:\Dataserver\Dummy" -AclObject $acl
```

### Security Model:
- Service account gets **Share Change** on Dataserver (network read + write access)
- Service account gets **Traverse** on parent folders (can walk through, cannot list contents)
- Service account gets **Modify** on Finance_Input folder (needed for file lock checks)
- Service account gets **Modify** on Finance_Output folder (read + write output files)
- All other folders under Dataserver remain completely inaccessible
- This provides least-privilege access: only what's needed, nothing more

### Why Additional Fixes Were Needed:

**Issue 1 - Traverse Permissions:**
Initial script (v1) failed because Windows requires traverse permissions on ALL parent folders in a path to access a subfolder over the network. Without traverse on `Dataserver` and `Dummy`, the service account couldn't navigate to `Finance_Input` even with direct permissions on it.

**Issue 2 - Share Permission Limitation:**
After adding traverse permissions, the script still failed with "Permission denied" when trying to open files in read+write mode. The issue was that the **Share-level** permission was set to "Read", which overrode the NTFS "Modify" permission.

Windows applies the **most restrictive** permission between Share-level and NTFS-level:
- Initial: Share=Read (restrictive) + NTFS=Modify (permissive) = **Result: Read only**
- Fixed: Share=Change (permissive) + NTFS=Modify (permissive) = **Result: Modify works**

**Issue 3 - Finance_Input Write Access:**
The depot script uses `ahlib.is_file_open_windows()` which opens files in `r+b` mode to check if they're locked by another process. This requires write access even for "read-only" input files. Finance_Input was upgraded from Read to Modify to allow this safety check.

---

## Final Permission Configuration (APPLIED)

### Network Shares:
| Share Name | Path | Service Account Permission |
|------------|------|---------------------------|
| `_Batchprozesse` | `D:\Dataserver\_Batchprozesse` | Change (Full Control) |
| `Dataserver` | `D:\Dataserver` | Change (Read + Write) |

### NTFS Permissions:
| Path | Permission | Inheritance | Purpose |
|------|-----------|-------------|---------|
| `D:\Dataserver` | Traverse | None (this folder only) | Navigate through |
| `D:\Dataserver\Dummy` | Traverse | None (this folder only) | Navigate through |
| `D:\Dataserver\Dummy\Finance_Input` | Modify | Yes (all files/subfolders) | Read files + file lock check |
| `D:\Dataserver\Dummy\Finance_Output` | Modify | Yes (all files/subfolders) | Write output files |
| `D:\Dataserver\_Batchprozesse\depot` | Modify | Yes (inherited from _Batchprozesse share) | Access depot scripts and logs |

---

## Additional Fix Scripts Created

### Fix-FinanceInputModify.ps1
**Purpose:** Upgrade Finance_Input from Read to Modify permission
**Status:** Executed successfully
**Location:** `D:\Dataserver\_Batchprozesse\depot\backups\2025-10-20_unc_path_fix\Fix-FinanceInputModify.ps1`

### Fix-DataserverSharePermission.ps1
**Purpose:** Upgrade Dataserver share permission from Read to Change
**Status:** Executed successfully
**Location:** `D:\Dataserver\_Batchprozesse\depot\backups\2025-10-20_unc_path_fix\Fix-DataserverSharePermission.ps1`

---

## Complete Undo Instructions

To completely remove all permissions granted to the service account:

### Step 1: Run Automated Undo Script
```powershell
cd D:\Dataserver\_Batchprozesse\depot\backups\2025-10-20_unc_path_fix
.\Remove-ServiceAccountPermissions.ps1
```

### Step 2: Remove Traverse Permissions (Manual)
```powershell
# Remove from Dataserver root
$acl = Get-Acl "D:\Dataserver"
$acl.Access | Where-Object { $_.IdentityReference -eq "WIN-H7BKO5H0RMC\service" } | ForEach-Object { $acl.RemoveAccessRule($_) }
Set-Acl "D:\Dataserver" -AclObject $acl

# Remove from Dummy folder
$acl = Get-Acl "D:\Dataserver\Dummy"
$acl.Access | Where-Object { $_.IdentityReference -eq "WIN-H7BKO5H0RMC\service" } | ForEach-Object { $acl.RemoveAccessRule($_) }
Set-Acl "D:\Dataserver\Dummy" -AclObject $acl

# Remove from Finance_Input
$acl = Get-Acl "D:\Dataserver\Dummy\Finance_Input"
$acl.Access | Where-Object { $_.IdentityReference -eq "WIN-H7BKO5H0RMC\service" } | ForEach-Object { $acl.RemoveAccessRule($_) }
Set-Acl "D:\Dataserver\Dummy\Finance_Input" -AclObject $acl

# Remove from Finance_Output
$acl = Get-Acl "D:\Dataserver\Dummy\Finance_Output"
$acl.Access | Where-Object { $_.IdentityReference -eq "WIN-H7BKO5H0RMC\service" } | ForEach-Object { $acl.RemoveAccessRule($_) }
Set-Acl "D:\Dataserver\Dummy\Finance_Output" -AclObject $acl
```

### Step 3: Remove Share Permissions
```powershell
# Remove from _Batchprozesse share
Revoke-SmbShareAccess -Name "_Batchprozesse" -AccountName "WIN-H7BKO5H0RMC\service" -Force

# Remove from Dataserver share
Revoke-SmbShareAccess -Name "Dataserver" -AccountName "WIN-H7BKO5H0RMC\service" -Force
```

### Step 4: Revert start_depot.ps1
```powershell
cd D:\Dataserver\_Batchprozesse\depot
Copy-Item -Path ".\backups\2025-10-20_unc_path_fix\start_depot.ps1.backup" -Destination ".\start_depot.ps1" -Force
```

---

## Verification Checklist
- [x] Original file backed up
- [x] Changes documented
- [x] Undo instructions provided
- [x] Testing instructions included
- [x] Changes applied to production file (start_depot.ps1)
- [x] Tested as service user (depot script runs, logs created)
- [x] Permission scripts created
- [x] Set-ServiceAccountPermissions-v2.ps1 executed successfully
- [x] Fix-FinanceInputModify.ps1 executed successfully
- [x] Fix-DataserverSharePermission.ps1 executed successfully
- [x] Finance folder access tested (dir commands work)
- [x] Full depot script tested as service user (ExitCode=0)
- [ ] Scheduled task verified working at 5 AM (to be tested on 2025-10-21)

---

## Test Results

### Manual Test Run (2025-10-20 18:44:41)
**Status:** ✅ SUCCESS
**User:** Service
**Exit Code:** 0
**Duration:** ~65 seconds
**Results:**
- Successfully read 3 input files from Finance_Input
- Successfully wrote 38+ output files to Finance_Output
- All file operations completed without permission errors
- Log file written successfully

### Key Log Entries:
```
[2025-10-20 18:44:41] Script started by user: Service
INFO: Verfügbarkeitscheck abgeschlossen: 5/5 Dateien verfügbar.
INFO: Alle Dateien verfügbar und erfolgreich geladen.
[Multiple successful exports to Finance_Output...]
[2025-10-20 18:45:46] ENDE Depot Script (ExitCode=0)
```

---

## Git Configuration Issues Fixed

After the initial implementation, two git-related issues were discovered and resolved:

### Issue 1: Duplicate Git Safe Directory Entries

**Problem:**
The service account had multiple duplicate `safe.directory` entries in git global config:
```
'\\WIN-H7BKO5H0RMC\Dataserver\_Batchprozesse\depot'  (old path)
//WIN-H7BKO5H0RMC/Dataserver/_Batchprozesse/depot  (old path)
//WIN-H7BKO5H0RMC/_Batchprozesse/depot  (duplicated 6 times)
```

This caused warning messages in the logs:
```
warning: safe.directory ''\\WIN-H7BKO5H0RMC\Dataserver\_Batchprozesse\depot'' not absolute
```

**Solution:**
Created `Cleanup-GitConfig.ps1` script that:
1. Removed all existing safe.directory entries using `git config --global --unset-all safe.directory`
2. Added only the correct paths:
   - `//WIN-H7BKO5H0RMC/_Batchprozesse/depot` (for UNC access)
   - `D:/Dataserver/_Batchprozesse/depot` (for local access ownership check)

**Script Location:** `D:\Dataserver\_Batchprozesse\depot\backups\2025-10-20_unc_path_fix\Cleanup-GitConfig.ps1`

### Issue 2: SSH Authentication for Git Pull

**Problem:**
Git pull was failing with:
```
Host key verification failed.
fatal: Could not read from remote repository.
```

The service account had no SSH keys configured for GitHub authentication.

**Solution:**
Created `Setup-GitSSH.ps1` script that:
1. Generated SSH key pair (ed25519 type) for service account
2. Stored keys in `C:\Users\Service\.ssh\`
3. Displayed public key for adding to GitHub

**Steps Performed:**
1. Ran `Setup-GitSSH.ps1` as service user
2. Added public key to GitHub account SSH keys
3. Tested connection: `ssh -T git@github.com` - **SUCCESS**
4. Verified git remote was already using SSH: `git@github.com:Drahmue/depot.git`
5. Tested git pull: `git pull origin main` - **SUCCESS**

**Script Location:** `D:\Dataserver\_Batchprozesse\depot\backups\2025-10-20_unc_path_fix\Setup-GitSSH.ps1`

**Public Key Added to GitHub:**
```
ssh-ed25519 AAAAC3NzaC1lZDI1NTE5AAAAIKqP5CjfY+hSTMBm4vuZV/XCsbxoSkhv+oR3IsmxeEwe service@WIN-H7BKO5H0RMC
```

---

## Test Results

### Manual Test Run (2025-10-20 18:44:41)
**Status:** ✅ SUCCESS
**User:** Service
**Exit Code:** 0
**Duration:** ~65 seconds
**Results:**
- Successfully read 3 input files from Finance_Input
- Successfully wrote 38+ output files to Finance_Output
- All file operations completed without permission errors
- Log file written successfully

**Key Log Entries:**
```
[2025-10-20 18:44:41] Script started by user: Service
INFO: Verfügbarkeitscheck abgeschlossen: 5/5 Dateien verfügbar.
INFO: Alle Dateien verfügbar und erfolgreich geladen.
[Multiple successful exports to Finance_Output...]
[2025-10-20 18:45:46] ENDE Depot Script (ExitCode=0)
```

### Final Test Run with Git Fixes (2025-10-20 22:21:40)
**Status:** ✅ SUCCESS
**User:** Service
**Exit Code:** 0
**Git Pull:** ✅ Working (no warnings or errors)

**Clean Log Output:**
```
[2025-10-20 22:21:40] Script started by user: Service
[2025-10-20 22:21:40] Changed directory to: \\WIN-H7BKO5H0RMC\_Batchprozesse\depot
[2025-10-20 22:21:41] Pulling updates from GitHub
From github.com:Drahmue/depot
 * branch            main       -> FETCH_HEAD
Already up to date.
Import der ahlib Bibliothek erfolgreich
[Script continues successfully...]
```

**Comparison:**

| Issue | Before | After |
|-------|--------|-------|
| Git warnings | ⚠️ 2 warnings | ✅ None |
| Git pull | ❌ Failed | ✅ Success |
| File access | ❌ Permission denied | ✅ All files accessible |
| Script execution | ❌ ExitCode=1 | ✅ ExitCode=0 |

---

## Complete Solution Summary

### Files Modified:
1. `start_depot.ps1` - Updated UNC path from Dataserver share to _Batchprozesse share
2. Git global config (service user) - Cleaned up safe.directory entries
3. SSH keys (service user) - Generated and added to GitHub

### Scripts Created:
1. `Set-ServiceAccountPermissions-v2.ps1` - Applied all necessary permissions
2. `Fix-FinanceInputModify.ps1` - Upgraded Finance_Input to Modify
3. `Fix-DataserverSharePermission.ps1` - Upgraded Dataserver share to Change
4. `Cleanup-GitConfig.ps1` - Cleaned git configuration
5. `Setup-GitSSH.ps1` - Setup SSH keys for GitHub
6. `Remove-ServiceAccountPermissions.ps1` - Undo script

### Permissions Applied:
- **Share:** `_Batchprozesse` - Change permission
- **Share:** `Dataserver` - Change permission
- **NTFS:** `D:\Dataserver` - Traverse (no inheritance)
- **NTFS:** `D:\Dataserver\Dummy` - Traverse (no inheritance)
- **NTFS:** `D:\Dataserver\Dummy\Finance_Input` - Modify (with inheritance)
- **NTFS:** `D:\Dataserver\Dummy\Finance_Output` - Modify (with inheritance)

### Security Configuration:
- SSH keys generated for automated git operations
- Service account has minimal required permissions
- All other Dataserver folders remain inaccessible
- Least-privilege access model maintained

---

## Verification Checklist
- [x] Original file backed up
- [x] Changes documented
- [x] Undo instructions provided
- [x] Testing instructions included
- [x] Changes applied to production file (start_depot.ps1)
- [x] Tested as service user (depot script runs, logs created)
- [x] Permission scripts created and executed
- [x] Set-ServiceAccountPermissions-v2.ps1 executed successfully
- [x] Fix-FinanceInputModify.ps1 executed successfully
- [x] Fix-DataserverSharePermission.ps1 executed successfully
- [x] Finance folder access tested (dir commands work)
- [x] Full depot script tested as service user (ExitCode=0)
- [x] Git configuration cleaned (Cleanup-GitConfig.ps1)
- [x] SSH keys setup (Setup-GitSSH.ps1)
- [x] Git pull tested and working
- [x] Final test run with clean logs (no warnings)
- [ ] Scheduled task verified working at 5 AM (to be tested on 2025-10-21)

---

## Next Steps

1. **Monitor scheduled task** on 2025-10-21 at 05:00 AM
2. **Verify log entry** in `depot_2025-10.log` shows:
   - Successful git pull
   - No warnings or errors
   - ExitCode=0
3. **Check output files** were generated in Finance_Output folder
4. If successful, the issue is fully resolved

---

## Important Notes

### Security Considerations:
- SSH private key for service account is stored unencrypted (required for automation)
- Location: `C:\Users\Service\.ssh\id_ed25519`
- This is standard practice for automated git operations
- Only the service account can access this file (Windows file permissions)

### Maintenance:
- If SSH key needs to be rotated, re-run `Setup-GitSSH.ps1` and update GitHub
- If permissions need to be revoked, follow the "Complete Undo Instructions" section
- All scripts are saved in `backups/2025-10-20_unc_path_fix/` for future reference

### Known Limitations:
- Git pull requires network access to github.com (port 22)
- SSH key has no passphrase (required for unattended automation)
- Service account has write access to Finance_Input (needed for file lock checks)
