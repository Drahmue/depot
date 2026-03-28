# Fehler-Benachrichtigung im Projekt Depot – Dokumentation

## Kontext

Dieses Dokument beschreibt die Integration des PowerShell-Benachrichtigungssystems
in das Projekt Depot. Die Implementierung erfolgte am 2026-03-28 basierend auf den
Erfahrungen aus dem Projekt MyFitnessPal_Sync.

## Architektur (Option B – autonom pro Projekt)

Jedes Projektverzeichnis enthaelt seine eigene Kopie der Benachrichtigungsdateien:
- `Send-ErrorNotification.ps1`  – Bibliothek mit Email- und Telegram-Funktionen
- `notify_config.json`          – Credentials (NICHT im Git-Repo, nur lokal)

## Projektstruktur Depot

| Datei | Beschreibung |
|---|---|
| `depot.py` | Hauptskript (Kursabruf, Portfolioberechnung) |
| `start_depot.ps1` | PowerShell-Wrapper, wird taeglich per Task Scheduler ausgefuehrt |
| `Send-ErrorNotification.ps1` | Notification-Bibliothek (Email + Telegram) |
| `notify_config.json` | Lokale Konfiguration mit Credentials (nicht committet) |
| `logs/depot_YYYY-MM.log` | Monatlich rotierendes Logfile |
| `logs/depot_errors_YYYY-MM.log` | Fehler-Logfile |

**Projektverzeichnis:** `D:\Dataserver\_Batchprozesse\depot`
**UNC-Pfad (Task Scheduler):** `\\WIN-H7BKO5H0RMC\_Batchprozesse\depot`
**Python:** `.venv\Scripts\python.exe`

## Durchgefuehrte Aenderungen (2026-03-28)

### 1. Send-ErrorNotification.ps1 hinzugefuegt
Direkte Kopie aus MyFitnessPal_Sync. Generische Bibliothek, keine Anpassungen noetig.
Unterstuetzt Email (Gmail SMTP) und Telegram Bot.

### 2. notify_config.json erstellt (lokal, nicht committet)
Gleiche Konfiguration wie MyFitnessPal_Sync:
- Telegram aktiv (`notify_telegram: true`)
- Email deaktiviert (`notify_email: false`)
- Empfaenger: `leo@haunschild-family.de`, `ah@haunschild-family.de`
- Computername: `WIN-H7BKO5H0RMC`

### 3. start_depot.ps1 angepasst

**Dot-Sourcing nach den Pfadvariablen (Zeilen 13-22):**
```powershell
$notifyAvailable = $false
$notifyLib = Join-Path $scriptDir "Send-ErrorNotification.ps1"
if (Test-Path $notifyLib) {
    try {
        . $notifyLib
        $notifyAvailable = $true
    }
    catch { Write-Warning "FEHLER beim Laden der Notification-Bibliothek: $_" }
}
```

**Notification-Aufruf im catch-Block (Zeilen 98-101):**
```powershell
if ($notifyAvailable) {
    Send-ErrorNotification -ScriptName "depot" -ExitCode $RC `
        -ErrorMessage $($_.Exception.Message) -LogFile $LOGFILE
}
```

### 4. .gitignore aktualisiert
`notify_config.json` hinzugefuegt (Abschnitt "Notification credentials").

**Git-Commit:** `b1ba2cf` – gepusht nach `origin/main`.

## Bekannte Einschraenkung: UNC-Pfad bei manueller Ausfuehrung

`start_depot.ps1` verwendet `$scriptDir = "\\WIN-H7BKO5H0RMC\_Batchprozesse\depot"`.
Bei manueller Ausfuehrung als normaler Benutzer (nicht SYSTEM) entstehen
"Zugriff verweigert"-Fehler auf den UNC-Pfad — das ist ein **pre-existierendes Problem**,
nicht durch die Notification-Aenderungen verursacht.

Das Skript laeuft korrekt, wenn es vom Task Scheduler unter dem SYSTEM-Konto ausgefuehrt wird.

### Notification manuell testen (Workaround)

Statt das Hauptskript auszufuehren, die Bibliothek direkt aufrufen:

```powershell
cd "D:\Dataserver\_Batchprozesse\depot"
. .\Send-ErrorNotification.ps1
Send-ErrorNotification -ScriptName "depot" -ExitCode 1 -ErrorMessage "Das ist ein Test"
```

## Offene Punkte / Moegliche Verbesserungen

- **UNC-Pfad-Problem loesen:** `$scriptDir` koennte auf `$PSScriptRoot` umgestellt werden,
  damit das Skript sowohl manuell (lokaler Pfad) als auch per Task Scheduler (UNC-Pfad)
  korrekt laeuft. Aenderung benoetigt Pruefung ob Task Scheduler dadurch beeinflusst wird.

- **End-to-End-Test:** Erster echter Test erfolgt beim naechsten Task-Scheduler-Lauf
  (taeglich 05:00 Uhr). Bei Fehler sollte eine Telegram-Nachricht ankommen.

## notify_config.json Struktur (Referenz)

```json
{
  "smtp_server":        "smtp.gmail.com",
  "smtp_port":          587,
  "smtp_user":          "GMAIL_ADRESSE",
  "smtp_password":      "GMAIL_APP_PASSWORT",
  "from_address":       "GMAIL_ADRESSE",
  "to_addresses":       ["EMPFAENGER_1", "EMPFAENGER_2"],
  "telegram_bot_token": "BOT_TOKEN",
  "telegram_chat_id":   "CHAT_ID",
  "notify_email":       false,
  "notify_telegram":    true,
  "computername":       "WIN-H7BKO5H0RMC"
}
```
