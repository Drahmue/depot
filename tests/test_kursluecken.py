"""Offline-Tests DH-054 Schritt 4 (Kurslücken) – ohne Netz, ohne echte Dateien."""
import sys
from pathlib import Path
from datetime import datetime as _dt, date
import numpy as np
import pandas as pd

sys.path.insert(0, str(Path(__file__).resolve().parent.parent))
import depot  # noqa: E402


class Log:
    def __init__(self):
        self.lines = []

    def info(self, m):
        self.lines.append(("INFO", m))

    def warning(self, m):
        self.lines.append(("WARNING", m))

    def error(self, m):
        self.lines.append(("ERROR", m))


def set_today(d):
    class FakeDT(_dt):
        @classmethod
        def today(cls):
            return _dt(d.year, d.month, d.day, 5, 0)
    depot.datetime = FakeDT


# Kursquelle: BTC-Kurs für jeden Tag bis 'available_until'
AVAILABLE = {"BTC-EUR": None, "EUNL.DE": None}
CALLS = []


def fake_download(ticker, start, end, **_k):
    CALLS.append((ticker, pd.Timestamp(start).date(), pd.Timestamp(end).date()))
    until = AVAILABLE[ticker]
    idx = [d for d in pd.date_range(start, pd.Timestamp(end) - pd.Timedelta(days=1)) if until and d <= until]
    if ticker != "BTC-EUR":
        idx = [d for d in idx if d.weekday() < 5]
    return pd.DataFrame({"Close": [100.0 + d.day for d in idx]}, index=pd.DatetimeIndex(idx))


depot.yf.download = fake_download

instruments = pd.DataFrame({"ticker": ["BTC-EUR", "EUNL.DE", None],
                            "default_value": [None, None, 1.0]},
                           index=pd.Index(["a2yzk6", "a0rpwh", "cash"], name="wkn"))


def base_prices():
    rows = []
    for d in pd.bdate_range("2026-11-02", "2026-11-06"):  # Mo–Fr
        for w in ["a2yzk6", "a0rpwh", "cash"]:
            rows.append((d, w, 1.0))
    rows.append((pd.Timestamp("2026-11-04"), "a2yzk6", np.nan))  # alte NaN-Lücke (Mi)
    df = pd.DataFrame(rows, columns=["date", "wkn", "price"])
    df = df.drop_duplicates(["date", "wkn"], keep="last")
    return df.set_index(["date", "wkn"])


ok = fail = 0


def check(name, cond):
    global ok, fail
    if cond:
        ok += 1
        print("OK  ", name)
    else:
        fail += 1
        print("FEHL", name)


def price(p, d, w):
    try:
        return p.loc[(pd.Timestamp(d), w), "price"]
    except KeyError:
        return "fehlt"


# 1) Montag 09.11. 05:00: Vortag Fr 06.11. ist vorhanden -> nur NaN vom 04.11. erneut anfragen
set_today(date(2026, 11, 9))
log = Log()
AVAILABLE["BTC-EUR"] = pd.Timestamp("2026-11-08")
AVAILABLE["EUNL.DE"] = pd.Timestamp("2026-11-06")
p = depot.prices_update(base_prices(), instruments, log)
check("1 alte NaN 04.11. nachgeholt", price(p, "2026-11-04", "a2yzk6") == 104.0)

# 2) Mittwoch 11.11. 05:00: BTC nur bis 09.11. verfügbar -> 10.11. NICHT als NaN schreiben
set_today(date(2026, 11, 11))
log = Log()
AVAILABLE["BTC-EUR"] = pd.Timestamp("2026-11-09")
AVAILABLE["EUNL.DE"] = pd.Timestamp("2026-11-10")
p = depot.prices_update(p, instruments, log)
check("2a BTC 09.11. geschrieben", price(p, "2026-11-09", "a2yzk6") == 109.0)
check("2b BTC 10.11. kein NaN-Eintrag", price(p, "2026-11-10", "a2yzk6") == "fehlt")
check("2c ETF 10.11. geschrieben", price(p, "2026-11-10", "a0rpwh") == 110.0)
check("2d Vortag fehlt nur als INFO", any(l == "INFO" and "10.11" not in m and "2026-11-10" in m for l, m in log.lines)
      and not any(l == "WARNING" for l, m in log.lines))
check("2e cash Defaultwert", price(p, "2026-11-10", "cash") == 1.0)

# 3) Donnerstag 12.11.: BTC jetzt bis 11.11. -> 10.11. wird nachgeholt, 11.11. auch
set_today(date(2026, 11, 12))
log = Log()
AVAILABLE["BTC-EUR"] = pd.Timestamp("2026-11-11")
p = depot.prices_update(p, instruments, log)
check("3a BTC 10.11. nachgeholt", price(p, "2026-11-10", "a2yzk6") == 110.0)
check("3b BTC 11.11. geschrieben", price(p, "2026-11-11", "a2yzk6") == 111.0)
check("3c keine NaN in prices", not p["price"].isna().any())

# 4) Ausfall: BTC liefert ab 12.11. nichts mehr; am 18.11. -> 4 Handelstage ohne Kurs -> Warnung
for d in [date(2026, 11, 13), date(2026, 11, 16), date(2026, 11, 17), date(2026, 11, 18)]:
    set_today(d)
    log = Log()
    AVAILABLE["EUNL.DE"] = pd.Timestamp(d) - pd.Timedelta(days=1)
    p = depot.prices_update(p, instruments, log)
stale = depot.prices_stale_check(p, instruments, Log(), today=date(2026, 11, 18))
check("4a Warnung bei 4 Handelstagen (12.–17.11.)", len(stale) == 1 and "a2yzk6" in stale[0] and "4 Handelstage" in stale[0])
stale17 = depot.prices_stale_check(p, instruments, Log(), today=date(2026, 11, 17))
check("4b keine Warnung bei 3 Handelstagen", stale17 == [])
check("4c ältere Tage als WARNING geloggt", any(l == "WARNING" and "BTC-EUR" in m for l, m in log.lines))
check("4d ETF ohne Lücke nicht gemeldet", not any("a0rpwh" in s for s in stale))

# 5) Rückwirkendes Fenster: Lücke älter als 10 Tage wird nicht mehr angefragt
#    (letztes Datum 27.11., Kurse fehlen seit 12.11. -> nur Handelstage ab 21.11. erneut)
dates = depot.price_dates_to_fetch(p, "a2yzk6", pd.Timestamp("2026-11-27"), pd.Timestamp("2026-11-30"),
                                   depot.holidays.Germany())
check("5a Fenster 10 Tage (ab 21.11. -> erster Handelstag 23.11.)", min(dates) == pd.Timestamp("2026-11-23"))
check("5b Wochenende nicht angefragt", all(d.weekday() < 5 for d in dates))
dates = depot.price_dates_to_fetch(p, "a0rpwh", pd.Timestamp("2026-11-17"), pd.Timestamp("2026-11-18"),
                                   depot.holidays.Germany())
check("5c lückenlose WKN: nur der neue Tag", dates == [pd.Timestamp("2026-11-18")])

print(f"\n{ok}/{ok + fail} Tests bestanden")
sys.exit(1 if fail else 0)
