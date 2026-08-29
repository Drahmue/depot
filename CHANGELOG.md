# Changelog

## 2026-08-29 — Fehlende Kurse für 28.8. durch abgelaufenes Service-Account-Passwort + stille NaN-Übernahme

**Beobachtung:**
Nutzer bemerkte, dass die Reports für den 28.8.2026 dieselben Kurse wie für den 27.8.2026 zeigten ("keine neuen Kurse geliefert").

**Ursache (Kette):**
1. Das Passwort des Service-Accounts, unter dem der Windows Task Scheduler-Job läuft, war abgelaufen. Dadurch ist der reguläre Tageslauf (sonst täglich 05:00:02 Uhr) am **27.8. und 28.8. gar nicht gestartet** — 2-Tage-Lücke in `logs/depot_2026-08.log` (letzter regulärer Lauf: 26.8. 05:01, nächster erst 29.8. 09:39).
2. Der Nachhol-Lauf am 29.8. hat die fehlenden Tage (27./28.8.) über `prices_update()` (depot.py:423–533) nachgeholt. Für die meisten WKNs kam dabei am 28.8. von yfinance zwar eine Zeile zurück, aber mit leerem/NaN `Close`-Wert — das wird von `prices_update()` aktuell **ohne Warnung** übernommen (nur der explizite `KeyError`-Fall, wenn für das Datum gar keine Zeile existiert, wird geloggt und gewarnt, s. depot.py:511–515).
3. `prices_import_and_process()` füllt fehlende/NaN-Kurse anschließend per Forward-Fill pro WKN auf (depot.py:1719–1723), bevor die Daten für alle Excel-Reports verwendet werden — dadurch erschien in den Reports für den 28.8. überall der Kurs vom 27.8., obwohl die zugrunde liegende `prices.parquet` für den 28.8. eigentlich `NaN` enthielt.

**Behebung:**
- Service-Account-Passwort erneuert und auf "läuft nie ab" gesetzt (durch Nutzer).
- 32 Zeilen für 2026-08-28 aus `prices.parquet` gelöscht (Backup: `backups/prices_backup_before_2026-08-28_fix_20260829_104122.parquet`), damit `prices_update()` sie beim nächsten Lauf sauber neu abfragt.
- `depot.py` manuell erneut ausgeführt (Exit Code 0) — für 20 von 21 Marktinstrumenten liegen jetzt echte, von 27.8. abweichende Kurse für den 28.8. vor.

**Offen / bekannte Lücke:**
- WKN `a2yzk6` (BTC-EUR) blieb nach dem Nachholen weiterhin ohne Kurs für den 28.8. (`NaN`) — yfinance liefert für diesen Ticker an dem Tag offenbar keine Daten. Wird beim nächsten regulären Lauf automatisch erneut versucht.
- **Strukturelles Risiko, unabhängig vom aktuellen Fall:** Sobald für ein (Datum, WKN) einmal eine Zeile mit `NaN`-Preis existiert, hält `prices_update()` dieses Datum für "bereits vorhanden" (Prüfung nur auf Index-Existenz, nicht auf validen Wert, s. depot.py:446–449 sowie den bereits dokumentierten Fix vom 2026-03-06 zu genau diesem Muster bei BTC-EUR) — der Tag wird dann **nie automatisch neu abgefragt** und muss wie hier manuell bereinigt werden. Ebenso fehlt aktuell eine Warnung, wenn yfinance zwar eine Zeile, aber einen NaN-`Close`-Wert liefert (nur der `KeyError`-Fall bei komplett fehlender Zeile wird geloggt).
- **Empfehlung für einen Code-Fix:** In `prices_update()` nach dem `Close`-Lookup zusätzlich `pd.isna(close_value)` prüfen und in diesem Fall ebenfalls warnen; zusätzlich `wkn_last_dates` nur aus Zeilen mit nicht-`NaN`-Preis bilden, damit Lücken automatisch erneut abgefragt werden.

**Betroffen:** `prices.parquet` (2026-08-28, 20 von 21 Marktinstrumenten korrigiert, BTC-EUR offen), Task-Scheduler Service-Account.

## 2026-07-30 — CM-Positionsabweichung durch negativen Bankbestand (Buchungskorrektur)

**Beobachtung:**
Manuelle Addition aller Buchungen (`delta`) für WKN CM ergab ein anderes Ergebnis als die von `depot.py` ausgewiesene Position in `shares_day.xlsx` (124.686,74 € vs. 124.692,96 €, Differenz 6,22 €).

**Ursache:**
`shares_from_bookings()` (depot.py:698–733) berechnet die Position pro Bank als laufende Summe und setzt einen kumulierten Bestand, der unter 0,0001 fällt, auf 0 (depot.py:726) — negative Bestände sind fachlich nicht zulässig. Bank `vw` hatte für CM einen kumulierten Bestand von −6,22 € (mehr abgebucht als je eingebucht), der dadurch auf 0 statt gegen andere Banken verrechnet wurde. Das ist kein Fehler in der Berechnungslogik von `depot.py`, sondern eine fehlerhafte/fehlende Buchung in `bookings.xlsx` bei Bank `vw`.

**Behebung:**
Korrekturbuchung in `bookings.xlsx` durch den Nutzer ergänzt. Nach der Korrektur liegt Bank `vw` bei 0,00 €, und manuelle Summe sowie `depot.py`-Ergebnis stimmen exakt überein (124.692,96 €, Restdifferenz nur Gleitkomma-Rauschen ~1e-11).

**Hinweis für künftige Fälle:**
Bei Abweichungen zwischen manueller Summe und `depot.py`-Ausgabe für ein Instrument zuerst die Bestände pro Bank prüfen (nicht nur die Gesamtsumme) — ein negativer Bestand bei einer einzelnen Bank wird von `shares_from_bookings()` stillschweigend auf 0 geflooört und erklärt eine positive Abweichung von `depot.py` gegenüber der reinen Buchungssumme.

**Betroffen:** WKN CM, Bank `vw`, siehe auch bereits vorhandene `cash_cm_ftd_corrections.csv` / `fix_cash_cm_ftd_transactions.py` für ähnliche frühere CM/FTD-Korrekturen.

## 2026-03-06 — Per-WKN last_date fix in prices_update()

**Bug fixed:**
`prices_update()` computed a single global `last_date` (maximum across all WKNs).
If one WKN missed an update for any reason (transient yfinance failure, empty response),
subsequent runs would skip its gap permanently because the global cutoff had already moved forward.

**Root cause identified via:**
WKN A2YZK6 (BTC-EUR) was stuck at 2026-02-27 while all other WKNs reached 2026-03-05.
yfinance returned data correctly — the issue was purely the global `last_date` logic.
The gap in `prices.parquet` was patched manually before deploying the fix.

**Change in `prices_update()` (depot.py):**
- Removed global `last_date` and `missing_dates` computed once before the loop
- Removed global early-exit `if not missing_dates: return prices`
- Added pre-computation of `wkn_last_dates` (per-WKN last known date via `groupby`)
- Inside the loop: each WKN now computes its own `last_date` and `missing_dates`
- Per-WKN `if not missing_dates: continue` replaces the global early-exit

**Effect:**
Each WKN independently catches up from its own last known date, regardless of
how current other WKNs are. A transient download failure for one WKN no longer
causes a permanent gap.

**Files changed:**
- `depot.py` — `prices_update()` function (lines 423–540)
- `CHANGELOG.md` — created
- `README.md` — updated

**Backup:** `backups/depot_2026-03-06_prices_update_per_wkn_fix.py`
