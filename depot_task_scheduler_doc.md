# PowerShell Task Scheduler Automation - Dokumentation

## Überblick

Dieses Dokument beschreibt die Implementierung einer automatisierten Lösung für die Ausführung von Python-Skripten über den Windows Task Scheduler mit PowerShell. Die Lösung wurde entwickelt, um ein Depot-Management-Skript regelmäßig und automatisiert auszuführen, wobei verschiedene technische Herausforderungen gelöst wurden.

## Problem und Herausforderungen

### Ursprüngliche Situation
- Ein Python-Skript (`depot.py`) sollte regelmäßig über den Windows Task Scheduler ausgeführt werden
- Das Skript befindet sich auf einem Netzlaufwerk (UNC-Pfad: `\\WIN-H7BKO5H0RMC\Dataserver\_Batchprozesse\depot`)
- Ursprünglich wurde eine Batch-Datei verwendet, die jedoch Probleme mit UNC-Pfaden hatte

### Technische Herausforderungen
1. **UNC-Pfad Limitierungen**: CMD.EXE unterstützt keine UNC-Pfade als Arbeitsverzeichnis
2. **PowerShell FileSystem Provider**: PowerShell fügt bei UNC-Pfaden automatisch Provider-Prefixe hinzu (`Microsoft.PowerShell.Core\FileSystem::`)
3. **Task Scheduler Execution Policy**: Skripte von Netzlaufwerken werden durch PowerShell-Ausführungsrichtlinien blockiert
4. **Benutzerkontext**: Task Scheduler läuft in anderem Sicherheitskontext als interaktive Sitzungen

## Lösungsansatz

### 1. Migration von Batch zu PowerShell

**Warum PowerShell?**
- Bessere Unterstützung für UNC-Pfade
- Flexiblere Fehlerbehandlung
- Native UTF-8 Unterstützung
- Erweiterte Logging-Funktionen

### 2. Pfad-Management Strategie

**Problem**: PowerShell's `$PWD` und `Join-Path` erzeugten FileSystem Provider Prefixe:
```
Microsoft.PowerShell.Core\FileSystem::\\WIN-H7BKO5H0RMC\Dataserver\_Batchprozesse\depot
```

**Lösung**: Verwendung von expliziten String-Pfaden:
```powershell
$scriptDir = "\\WIN-H7BKO5H0RMC\Dataserver\_Batchprozesse\depot"
$pythonPath = "$scriptDir\.venv\Scripts\python.exe"
$scriptPath = "$scriptDir\depot.py"
```

### 3. Task Scheduler Integration

**Herausforderung**: Direkte Ausführung von PowerShell-Skripten von UNC-Pfaden wird blockiert.

**Lösung**: Verwendung von `pushd`/`popd` Befehlen:
```cmd
powershell.exe -ExecutionPolicy Bypass -Command "pushd '\\WIN-H7BKO5H0RMC\Dataserver\_Batchprozesse\depot'; & '.\start_depot.ps1'; popd"
```

**Funktionsweise**:
- `pushd` mappt temporär das UNC-Verzeichnis und wechselt dorthin
- `& '.\start_depot.ps1'` führt das Skript aus
- `popd` kehrt zum ursprünglichen Verzeichnis zurück und bereinigt die temporäre Zuordnung

## Finale Lösung

### PowerShell-Skript (start_depot.ps1)

```powershell
# Set error action preference and encoding
$ErrorActionPreference = "Continue"
[Console]::OutputEncoding = [System.Text.Encoding]::UTF8
$env:PYTHONIOENCODING = "utf-8"

# Ensure we're in the correct directory using UNC path
$scriptDir = "\\WIN-H7BKO5H0RMC\Dataserver\_Batchprozesse\depot"
Push-Location $scriptDir

# Setup logging directory and file
$LOGDIR = "$scriptDir\logs"
if (-not (Test-Path -Path $LOGDIR)) {
    New-Item -ItemType Directory -Path $LOGDIR -Force | Out-Null
}

$LOGSTAMP = (Get-Date).ToString("yyyy-MM")
$LOGFILE = "$LOGDIR\depot_$LOGSTAMP.log"

try {
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
    Add-Content -Path $LOGFILE -Value "$timestamp ERROR: $($_.Exception.Message)"
    $RC = 1
}

# Exit with the return code
Pop-Location
exit $RC
```

### Task Scheduler Konfiguration

**Programm/Skript**:
```
powershell.exe
```

**Argumente**:
```
-ExecutionPolicy Bypass -Command "pushd '\\WIN-H7BKO5H0RMC\Dataserver\_Batchprozesse\depot'; & '.\start_depot.ps1'; popd"
```

**Wichtige Einstellungen**:
- **Allgemein**: "Unabhängig von der Benutzeranmeldung ausführen"
- **Allgemein**: "Mit höchsten Privilegien ausführen" (falls erforderlich)
- **Bedingungen**: "Nur starten, wenn Computer im Netzbetrieb läuft" deaktivieren
- **Einstellungen**: "Bedarfsgesteuerte Ausführung der Aufgabe zulassen" aktivieren

## Funktionsweise im Detail

### 1. Encoding und Umgebung
```powershell
$ErrorActionPreference = "Continue"
[Console]::OutputEncoding = [System.Text.Encoding]::UTF8
$env:PYTHONIOENCODING = "utf-8"
```
- Setzt UTF-8 Encoding für Python und PowerShell
- Sorgt für korrekte Behandlung deutscher Umlaute
- Fortsetzung bei nicht-kritischen Fehlern

### 2. Verzeichniswechsel
```powershell
$scriptDir = "\\WIN-H7BKO5H0RMC\Dataserver\_Batchprozesse\depot"
Push-Location $scriptDir
```
- Explizite Definition des Arbeitsverzeichnisses
- `Push-Location` für saubere Verzeichnisverwaltung
- Vermeidung von FileSystem Provider Prefixen

### 3. Logging Setup
```powershell
$LOGSTAMP = (Get-Date).ToString("yyyy-MM")
$LOGFILE = "$LOGDIR\depot_$LOGSTAMP.log"
```
- Monatliche Log-Dateien (Format: depot_2025-09.log)
- Automatische Erstellung des logs-Verzeichnisses
- Strukturiertes Logging mit Zeitstempel

### 4. Git Integration
```powershell
$gitResult = & git pull origin main 2>&1
Add-Content -Path $LOGFILE -Value $gitResult
```
- Automatisches Update des Codes vor Ausführung
- Fehlerbehandlung falls Git Pull fehlschlägt
- Umleitung von STDOUT und STDERR in Log-Datei

### 5. Python Ausführung
```powershell
$pythonPath = "$scriptDir\.venv\Scripts\python.exe"
$scriptPath = "$scriptDir\depot.py"
$pythonResult = & $pythonPath -u $scriptPath 2>&1
```
- Verwendung der lokalen Virtual Environment
- Parameter `-u` für ungepufferte Ausgabe
- Umleitung aller Ausgaben ins Log

### 6. Cleanup und Exit
```powershell
$cutoffDate = (Get-Date).AddDays(-120)
Get-ChildItem -Path $LOGDIR -Filter "depot_*.log" | 
    Where-Object { $_.LastWriteTime -lt $cutoffDate } | 
    Remove-Item -Force -ErrorAction SilentlyContinue

Pop-Location
exit $RC
```
- Automatische Bereinigung alter Log-Dateien (>120 Tage)
- Rückkehr zum ursprünglichen Verzeichnis
- Weitergabe des Python-Exit-Codes an Task Scheduler

## Vorteile der Lösung

### 1. Robustheit
- Umfassende Fehlerbehandlung mit Try-Catch
- Graceful Degradation (Git Pull Fehler stoppt nicht die Hauptausführung)
- Saubere Bereinigung bei Fehlern

### 2. Wartbarkeit
- Strukturierte, monatliche Log-Dateien
- Automatische Log-Bereinigung
- Detaillierte Zeitstempel für Debugging

### 3. Automatisierung
- Automatische Code-Updates via Git
- Keine manuelle Intervention erforderlich
- Zuverlässige Task Scheduler Integration

### 4. Sicherheit
- Execution Policy Bypass nur für diese Aufgabe
- Keine systemweiten Sicherheitsänderungen
- Verwendung von Benutzerkonten mit minimalen Rechten möglich

## Troubleshooting

### Häufige Probleme

**1. "Datei kann nicht geladen werden" Fehler**
- **Ursache**: PowerShell Execution Policy
- **Lösung**: Verwendung von `-ExecutionPolicy Bypass` Parameter

**2. FileSystem Provider Prefix in Pfaden**
- **Ursache**: PowerShell's automatische Pfad-Konvertierung
- **Lösung**: Verwendung expliziter String-Pfade statt PowerShell-Cmdlets

**3. Task Scheduler Exit Code 2**
- **Ursache**: Datei nicht gefunden oder Pfad-Probleme
- **Lösung**: Überprüfung der UNC-Pfad Zugänglichkeit und Berechtigungen

**4. Leere Log-Dateien**
- **Ursache**: Skript startet nicht oder keine Netzwerk-Zugriff
- **Lösung**: Manuelle Ausführung zum Testen, Berechtigungen überprüfen

### Debugging-Ansätze

**1. Manuelle Ausführung**
```powershell
# Test in PowerShell ISE
Set-Location D:\Dataserver\_Batchprozesse\depot
.\start_depot.ps1
```

**2. Task Scheduler Simulation**
```powershell
# Test der exakten Task Scheduler Befehlszeile
powershell -ExecutionPolicy Bypass -Command "pushd '\\WIN-H7BKO5H0RMC\Dataserver\_Batchprozesse\depot'; & '.\start_depot.ps1'; popd"
```

**3. Berechtigungen testen**
```powershell
# Test des Netzwerk-Zugriffs
Test-Path "\\WIN-H7BKO5H0RMC\Dataserver\_Batchprozesse\depot"
```

## Best Practices

### 1. Entwicklung
- Immer zuerst manuell testen bevor Task Scheduler konfiguriert wird
- Verwendung von relativen Pfaden wo möglich
- Explizite Pfad-Definition für kritische Komponenten

### 2. Deployment
- Task Scheduler mit "Run only when user is logged on" für initiale Tests
- Erst nach erfolgreichen Tests auf "Run whether user is logged on or not" umstellen
- Verwendung dedizierter Service-Accounts für Produktionsumgebungen

### 3. Monitoring
- Regelmäßige Kontrolle der Log-Dateien
- Überwachung der Task Scheduler History
- Einrichtung von Alerts bei wiederkehrenden Fehlern

### 4. Wartung
- Monatliche Überprüfung der Log-Größen
- Backup der Skript-Dateien bei Änderungen
- Dokumentation von Konfigurationsänderungen

## Fazit

Die entwickelte Lösung bietet eine robuste, wartbare und automatisierte Möglichkeit, Python-Skripte von Netzlaufwerken über den Windows Task Scheduler auszuführen. Durch die Kombination verschiedener Techniken (PowerShell, pushd/popd, explizite Pfad-Verwaltung) werden die inhärenten Limitierungen von Windows UNC-Pfaden und Task Scheduler erfolgreich umgangen.

Die Lösung ist produktionstauglich und kann als Template für ähnliche Automatisierungsaufgaben verwendet werden.