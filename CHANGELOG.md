# Changelog

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
