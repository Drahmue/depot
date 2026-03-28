# =============================================================================
# Send-ErrorNotification.ps1
# Benachrichtigungsfunktion fuer geplante Python-Skripte.
# Kanaele: Email (Gmail SMTP) und/oder Telegram Bot.
# Konfiguration: notify_config.json im selben Verzeichnis wie dieses Skript.
#
# Einbindung per Dot-Sourcing im Wrapper:
#   . "$PSScriptRoot\Send-ErrorNotification.ps1"
# =============================================================================
# Version : 1.1
# Autor   : AH
# Datum   : 2026-03
# Aenderungen:
#   1.1 - Telegram-Kanal hinzugefuegt, notify_email/notify_telegram konfigurierbar
#   1.0 - Erstversion mit Gmail SMTP
# =============================================================================

$NOTIFICATION_LIB_VERSION = "1.1"

function Send-ErrorNotification {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $true)]
        [string]$ScriptName,

        [Parameter(Mandatory = $false)]
        [int]$ExitCode = -1,

        [Parameter(Mandatory = $false)]
        [string]$LogFile = "",

        [Parameter(Mandatory = $false)]
        [string]$ErrorMessage = "",

        [Parameter(Mandatory = $false)]
        [string]$ConfigFile = ""
    )

    if ([string]::IsNullOrEmpty($ConfigFile)) {
        $ConfigFile = Join-Path $PSScriptRoot "notify_config.json"
    }

    if (-not (Test-Path $ConfigFile)) {
        Write-Warning "[Notification] Konfigurationsdatei nicht gefunden: $ConfigFile"
        return
    }

    try {
        $cfg = Get-Content -Path $ConfigFile -Encoding UTF8 -Raw | ConvertFrom-Json
    }
    catch {
        Write-Warning "[Notification] Fehler beim Lesen der Konfiguration: $_"
        return
    }

    $timestamp   = Get-Date -Format "yyyy-MM-dd HH:mm:ss"
    $hostname    = if ($cfg.computername) { $cfg.computername } else { $env:COMPUTERNAME }
    $notifyEmail = if ($null -ne $cfg.notify_email) { [bool]$cfg.notify_email } else { $true }
    $notifyTG    = if ($null -ne $cfg.notify_telegram) { [bool]$cfg.notify_telegram } else { $false }

    $logText = ""
    if (-not [string]::IsNullOrEmpty($LogFile) -and (Test-Path $LogFile)) {
        $logText = (Get-Content -Path $LogFile -Encoding UTF8 -Tail 50) -join "`n"
    }

    if ($notifyEmail) {
        Send-EmailNotification -cfg $cfg -ScriptName $ScriptName -ExitCode $ExitCode `
            -ErrorMessage $ErrorMessage -logText $logText -timestamp $timestamp -hostname $hostname
    }

    if ($notifyTG) {
        Send-TelegramNotification -cfg $cfg -ScriptName $ScriptName -ExitCode $ExitCode `
            -ErrorMessage $ErrorMessage -logText $logText -timestamp $timestamp -hostname $hostname
    }
}

function Send-EmailNotification {
    param($cfg, $ScriptName, $ExitCode, $ErrorMessage, $logText, $timestamp, $hostname)

    $subject = "[FEHLER] $ScriptName auf $hostname - ExitCode $ExitCode"

    $htmlBody = @"
<!DOCTYPE html>
<html>
<head>
  <meta charset="UTF-8">
  <style>
    body      { font-family: Calibri, Arial, sans-serif; font-size: 14px; color: #222; }
    h2        { color: #c0392b; }
    table     { border-collapse: collapse; width: 100%; margin-bottom: 16px; }
    td, th    { border: 1px solid #ccc; padding: 6px 10px; text-align: left; }
    th        { background-color: #f0f0f0; font-weight: bold; width: 140px; }
    .errmsg   { background-color: #fdecea; color: #c0392b; padding: 8px;
                border-left: 4px solid #c0392b; margin-bottom: 12px; }
    .logblock { background-color: #f8f8f8; font-family: Consolas, monospace; font-size: 12px;
                padding: 10px; border: 1px solid #ddd; white-space: pre-wrap; }
    .footer   { color: #888; font-size: 11px; margin-top: 20px; }
  </style>
</head>
<body>
  <h2>&#9888; Fehler in geplantem Skript</h2>
  <table>
    <tr><th>Skript</th>    <td><b>$ScriptName</b></td></tr>
    <tr><th>Server</th>    <td>$hostname</td></tr>
    <tr><th>Zeitpunkt</th> <td>$timestamp</td></tr>
    <tr><th>Exit-Code</th> <td><b>$ExitCode</b></td></tr>
  </table>
"@

    if (-not [string]::IsNullOrEmpty($ErrorMessage)) {
        $escaped = [System.Web.HttpUtility]::HtmlEncode($ErrorMessage)
        $htmlBody += "  <div class=`"errmsg`"><b>Fehlermeldung:</b><br>$escaped</div>`n"
    }

    if (-not [string]::IsNullOrEmpty($logText)) {
        $escapedLog = [System.Web.HttpUtility]::HtmlEncode($logText)
        $htmlBody += "  <h3>Log-Ausgabe (letzte 50 Zeilen)</h3>`n"
        $htmlBody += "  <div class=`"logblock`">$escapedLog</div>`n"
    }

    $htmlBody += @"
  <p class="footer">Automatische Benachrichtigung | Aufgabenplanung $hostname</p>
</body>
</html>
"@

    try {
        $smtpClient             = New-Object System.Net.Mail.SmtpClient($cfg.smtp_server, $cfg.smtp_port)
        $smtpClient.EnableSsl   = $true
        $smtpClient.Credentials = New-Object System.Net.NetworkCredential($cfg.smtp_user, $cfg.smtp_password)

        $mailMsg                 = New-Object System.Net.Mail.MailMessage
        $mailMsg.From            = $cfg.from_address
        $mailMsg.Subject         = $subject
        $mailMsg.Body            = $htmlBody
        $mailMsg.IsBodyHtml      = $true
        $mailMsg.BodyEncoding    = [System.Text.Encoding]::UTF8
        $mailMsg.SubjectEncoding = [System.Text.Encoding]::UTF8

        if ($cfg.to_addresses) {
            foreach ($addr in $cfg.to_addresses) { $mailMsg.To.Add($addr.Trim()) }
        }
        elseif ($cfg.to_address) {
            $mailMsg.To.Add($cfg.to_address)
        }

        $smtpClient.Send($mailMsg)
        $recipients = if ($cfg.to_addresses) { $cfg.to_addresses -join ", " } else { $cfg.to_address }
        Write-Host "[Notification] Email gesendet an: $recipients"
        $mailMsg.Dispose()
        $smtpClient.Dispose()
    }
    catch {
        Write-Warning "[Notification] Email fehlgeschlagen: $_"
        try {
            Write-EventLog -LogName Application -Source "TaskScheduler" -EventId 9998 `
                -EntryType Warning `
                -Message "Email-Benachrichtigung fuer '$ScriptName' fehlgeschlagen: $_" `
                -ErrorAction SilentlyContinue
        }
        catch { }
    }
}

function Send-TelegramNotification {
    param($cfg, $ScriptName, $ExitCode, $ErrorMessage, $logText, $timestamp, $hostname)

    if ([string]::IsNullOrEmpty($cfg.telegram_bot_token) -or
        $cfg.telegram_bot_token -eq "TELEGRAM_BOT_TOKEN") {
        Write-Warning "[Notification] Telegram Bot-Token nicht konfiguriert."
        return
    }
    if ([string]::IsNullOrEmpty($cfg.telegram_chat_id) -or
        $cfg.telegram_chat_id -eq "TELEGRAM_CHAT_ID") {
        Write-Warning "[Notification] Telegram Chat-ID nicht konfiguriert."
        return
    }

    $msg  = "FEHLER in geplantem Skript`n"
    $msg += "========================`n"
    $msg += "Skript   : $ScriptName`n"
    $msg += "Server   : $hostname`n"
    $msg += "Zeit     : $timestamp`n"
    $msg += "ExitCode : $ExitCode`n"

    if (-not [string]::IsNullOrEmpty($ErrorMessage)) {
        $shortErr = if ($ErrorMessage.Length -gt 300) {
            $ErrorMessage.Substring(0, 300) + "..."
        } else { $ErrorMessage }
        $msg += "Fehler   : $shortErr`n"
    }

    if (-not [string]::IsNullOrEmpty($logText)) {
        $lastLines = ($logText -split "`n" | Select-Object -Last 5) -join "`n"
        $msg += "--- Letzte Log-Zeilen ---`n$lastLines"
    }

    $apiUrl = "https://api.telegram.org/bot$($cfg.telegram_bot_token)/sendMessage"
    $body   = @{
        chat_id    = $cfg.telegram_chat_id
        text       = $msg
        parse_mode = ""
    } | ConvertTo-Json -Compress

    try {
        $response = Invoke-RestMethod -Uri $apiUrl -Method Post -Body $body `
            -ContentType "application/json; charset=utf-8"
        if ($response.ok) {
            Write-Host "[Notification] Telegram-Nachricht gesendet (Chat-ID: $($cfg.telegram_chat_id))"
        }
        else {
            Write-Warning "[Notification] Telegram API Fehler: $($response.description)"
        }
    }
    catch {
        Write-Warning "[Notification] Telegram fehlgeschlagen: $_"
    }
}
