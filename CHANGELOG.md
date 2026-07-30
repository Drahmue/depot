# Changelog

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
