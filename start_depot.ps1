# Set error action preference and encoding
$ErrorActionPreference = "Continue"
[Console]::OutputEncoding = [System.Text.Encoding]::UTF8
$env:PYTHONIOENCODING = "utf-8"

# Early error logging setup (before main script execution)
$scriptDir = "\\WIN-H7BKO5H0RMC\_Batchprozesse\depot"
$LOGDIR = "$scriptDir\logs"
$LOGSTAMP = (Get-Date).ToString("yyyy-MM")
$LOGFILE = "$LOGDIR\depot_$LOGSTAMP.log"
$ERRORLOG = "$LOGDIR\depot_errors_$LOGSTAMP.log"

# Create logs directory if it doesn't exist
try {
    if (-not (Test-Path -Path $LOGDIR)) {
        New-Item -ItemType Directory -Path $LOGDIR -Force | Out-Null
    }
} catch {
    # If we can't even create the log directory, write to a fallback location
    $LOGFILE = "C:\Temp\depot_emergency_$LOGSTAMP.log"
    $ERRORLOG = "C:\Temp\depot_emergency_errors_$LOGSTAMP.log"
    $timestamp = Get-Date -Format "[yyyy-MM-dd HH:mm:ss]"
    "$timestamp CRITICAL: Could not create log directory at $LOGDIR. Error: $($_.Exception.Message)" | Out-File -FilePath $ERRORLOG -Append
}

# Log script start
try {
    $timestamp = Get-Date -Format "[yyyy-MM-dd HH:mm:ss]"
    Add-Content -Path $LOGFILE -Value "$timestamp Script started by user: $env:USERNAME"
} catch {
    $timestamp = Get-Date -Format "[yyyy-MM-dd HH:mm:ss]"
    "$timestamp ERROR: Could not write to log file: $($_.Exception.Message)" | Out-File -FilePath $ERRORLOG -Append
}

# Navigate to script directory
try {
    Push-Location $scriptDir
    $timestamp = Get-Date -Format "[yyyy-MM-dd HH:mm:ss]"
    Add-Content -Path $LOGFILE -Value "$timestamp Changed directory to: $scriptDir"
} catch {
    $timestamp = Get-Date -Format "[yyyy-MM-dd HH:mm:ss]"
    "$timestamp ERROR: Could not change to directory $scriptDir. Error: $($_.Exception.Message)" | Out-File -FilePath $ERRORLOG -Append
    Add-Content -Path $LOGFILE -Value "$timestamp ERROR: Could not change to directory $scriptDir"
    exit 1
}

try {
    # Configure git safe.directory for UNC path
    & git config --global --add safe.directory '//WIN-H7BKO5H0RMC/_Batchprozesse/depot' 2>&1 | Out-Null

    # Pull latest updates from GitHub
    $timestamp = Get-Date -Format "[yyyy-MM-dd HH:mm:ss]"
    Add-Content -Path $LOGFILE -Value "$timestamp Pulling updates from GitHub"

    $gitResult = & git pull origin main 2>&1
    Add-Content -Path $LOGFILE -Value $gitResult
    
    if ($LASTEXITCODE -ne 0) {
        $timestamp = Get-Date -Format "[yyyy-MM-dd HH:mm:ss]"
        Add-Content -Path $LOGFILE -Value "$timestamp Git pull failed, continuing with existing code"
    }
    
    # Main script execution
    $pythonPath = "$scriptDir\.venv\Scripts\python.exe"
    $scriptPath = "$scriptDir\depot.py"
    
    $pythonResult = & $pythonPath -u $scriptPath 2>&1
    Add-Content -Path $LOGFILE -Value $pythonResult
    $RC = $LASTEXITCODE
    
    # Log completion
    $timestamp = Get-Date -Format "[yyyy-MM-dd HH:mm:ss]"
    Add-Content -Path $LOGFILE -Value "$timestamp ENDE Depot Script (ExitCode=$RC)"
    
    # Clean up old log files (older than 120 days)
    $cutoffDate = (Get-Date).AddDays(-120)
    Get-ChildItem -Path $LOGDIR -Filter "depot_*.log" | 
        Where-Object { $_.LastWriteTime -lt $cutoffDate } | 
        Remove-Item -Force -ErrorAction SilentlyContinue
    
} catch {
    $timestamp = Get-Date -Format "[yyyy-MM-dd HH:mm:ss]"
    $errorMsg = "$timestamp ERROR: $($_.Exception.Message)"
    Add-Content -Path $LOGFILE -Value $errorMsg -ErrorAction SilentlyContinue
    "$errorMsg`nStack Trace: $($_.ScriptStackTrace)" | Out-File -FilePath $ERRORLOG -Append
    $RC = 1
}

# Exit with the return code
Pop-Location
exit $RC