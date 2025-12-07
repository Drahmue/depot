# -*- coding: utf-8 -*-
"""
Vergleich gespeicherter Kurse (Parquet, MultiIndex: (date, wkn), Spalte 'price')
mit yfinance-Schlusskursen (immer 'Close', Fallback 'Adj Close').

Neu/Änderungen:
- Parquet-Auswahl öffnet standardmäßig im **Ordner des Skripts** mit vorgeschlagenem Dateinamen 'prices.parquet'.
- Instruments-Excel hat Default: \\WIN-H7BKO5H0RMC\Dataserver\Dummy\Finance_Input\Instrumente.xlsx
- Export der Abweichungen nach "new prices.xlsx" (diff & pct_diff, inkl. backfill-Flag).
- Erstellung eines **korrigierten DataFrames** auf Basis des Original-DFs:
  * Wenn yfinance **neue Werte** liefert und im Original-DF bereits ein Wert steht **und** dieser abweicht → aktualisieren.
  * Wenn im Original-DF **kein Wert** steht (NaN) und yfinance einen Wert liefert → **auffüllen**.
  * **BACKFILL-Modus** (ENABLE_BACKFILL=True): Fügt historische Daten für WKNs mit limitierter Historie hinzu.
    - Erkennt automatisch Instrumente mit unvollständiger Historie (Start > globaler Start)
    - Ermittelt echtes Issue-Datum des Instruments über yfinance (period="max")
    - Füllt fehlende historische Daten nur ab Issue-Datum (verhindert Fehler bei neu emittierten Instrumenten)
    - Erstellt neue Datums-Zeilen im Output (markiert mit backfill=True)

Voraussetzungen:
- Python 3.12+
- pandas, yfinance, openpyxl, pyarrow (für Parquet)
- tkinter (für Dateiauswahl; fällt sonst auf CLI-Eingabe zurück)
"""

from __future__ import annotations

import os
import sys
import math
import logging
from typing import Dict, List, Optional, Tuple
from datetime import datetime
from pathlib import Path

import pandas as pd
import yfinance as yf

# ---------------------------- Konfiguration ---------------------------------

# Absoluter Toleranzwert für Preisvergleich (z. B. Rundungsdifferenzen)
ABS_TOL = 1e-4

# Spaltenname im Input-DF für den Preis
PRICE_COL = "price"

# yfinance-Spaltenpräferenz
YF_PREFERRED_COL = "Close"       # immer Close verwenden
YF_FALLBACK_COL = "Adj Close"    # Fallback

# WKNs (case-insensitive), die NICHT geprüft werden sollen
EXCLUDE_WKNS = {"cash", "cm", "ftd", "lb1kwr"}

# Defaults
DEFAULT_PARQUET_NAME = "prices.parquet"
DEFAULT_INSTRUMENTS_PATH = r"\\WIN-H7BKO5H0RMC\Dataserver\Dummy\Finance_Input\Instrumente.xlsx"

# Backfill configuration
# Set to True to enable historical data backfilling for WKNs with limited history
ENABLE_BACKFILL = True
# Global start date for backfilling (will use earliest date from other WKNs if None)
BACKFILL_START_DATE = None  # e.g., "2021-01-02" or None for auto-detect
# Minimum number of successful trading days to consider an instrument valid
MIN_TRADING_DAYS_THRESHOLD = 5

# ----------------------------------------------------------------------------

def setup_logging() -> None:
    logging.basicConfig(
        level=logging.INFO,
        format="%(asctime)s [%(levelname)s] %(message)s",
        datefmt="%H:%M:%S",
    )


def _script_dir() -> Path:
    """Ordner, aus dem das Skript gestartet wurde (robust)."""
    try:
        # Wenn als Skript gestartet:
        p = Path(sys.argv[0]).resolve()
        if p.is_file():
            return p.parent
        if p.exists():
            return p
    except Exception:
        pass
    # Fallback: Verzeichnis dieser Datei oder CWD
    try:
        return Path(__file__).resolve().parent
    except Exception:
        return Path.cwd()


# ------------------------------ File-Dialoge --------------------------------

def _select_path(title: str, filetypes: list[tuple[str, str]], initialdir: Optional[Path], initialfile: Optional[str], default_fallback: Optional[Path]) -> Optional[Path]:
    """Dateiauswahldialog mit initialdir/initialfile. Bei Abbruch kann auf default_fallback zurückgegriffen werden (falls existent)."""
    try:
        import tkinter as tk
        from tkinter import filedialog

        root = tk.Tk()
        root.withdraw()
        root.update()
        path_str = filedialog.askopenfilename(
            title=title,
            filetypes=filetypes,
            initialdir=str(initialdir) if initialdir else None,
            initialfile=str(initialfile) if initialfile else None,
        )
        root.destroy()

        if path_str:
            return Path(path_str)

        # Abbruch → ggf. Fallback verwenden
        if default_fallback is not None and Path(default_fallback).exists():
            logging.info("Keine Auswahl getroffen. Verwende Default: %s", default_fallback)
            return Path(default_fallback)
        return None
    except Exception as e:
        logging.warning("Dateidialog nicht verfügbar (%s). Fallback: CLI.", e)
        suggestion = ""
        if initialdir or initialfile:
            suggestion = f" [{(Path(initialdir) / initialfile) if (initialdir and initialfile) else (initialdir or initialfile)}]"
        p = input(f"{title}{suggestion}: ").strip().strip('"')
        if not p and default_fallback is not None and Path(default_fallback).exists():
            logging.info("Keine Eingabe. Verwende Default: %s", default_fallback)
            return Path(default_fallback)
        return Path(p) if p else None


def select_parquet_path() -> Path:
    """Parquet-Datei mit Kursdaten auswählen (erforderlich)."""
    base = _script_dir()
    initialdir = base
    initialfile = DEFAULT_PARQUET_NAME
    fallback = base / DEFAULT_PARQUET_NAME  # nur verwenden, wenn vorhanden
    p = _select_path(
        "Parquet-Datei mit Kursdaten auswählen",
        [("Parquet", "*.parquet"), ("Alle Dateien", "*.*")],
        initialdir=initialdir,
        initialfile=initialfile,
        default_fallback=fallback,
    )
    if p is None:
        raise SystemExit("Abbruch: keine Parquet-Datei ausgewählt.")
    return p


def select_excel_mapping_path() -> Optional[Path]:
    """Excel-Datei für WKN->Ticker Mapping auswählen (optional, mit Default-UNC)."""
    fallback = Path(DEFAULT_INSTRUMENTS_PATH)
    return _select_path(
        "OPTIONAL: Excel mit WKN-Ticker-Mapping auswählen",
        [("Excel", "*.xlsx *.xls"), ("Alle Dateien", "*.*")],
        initialdir=fallback.parent if fallback.parent.exists() else None,
        initialfile=fallback.name,
        default_fallback=fallback,
    )


# ------------------------------ Daten laden ---------------------------------

def read_input_parquet(parquet_path: Path) -> pd.DataFrame:
    """Liest das Parquet und stellt sicher, dass Struktur passt:
    MultiIndex (date, wkn) und Spalte PRICE_COL.
    """
    if not parquet_path.exists():
        raise FileNotFoundError(f"Datei nicht gefunden: {parquet_path}")
    df = pd.read_parquet(parquet_path)
    if PRICE_COL not in df.columns:
        raise ValueError(f"Erwartete Spalte '{PRICE_COL}' fehlt im DataFrame. Spalten: {list(df.columns)}")

    # MultiIndex sicherstellen
    if not isinstance(df.index, pd.MultiIndex) or df.index.nlevels != 2:
        raise ValueError("Erwartet: MultiIndex mit Leveln ('date','wkn').")

    # Benenne Indexlevel robust
    level_names = list(df.index.names)
    if len(level_names) != 2:
        raise ValueError("MultiIndex hat nicht genau 2 Ebenen.")
    # Mappe mögliche Namen auf 'date' und 'wkn'
    name_map = {}
    for i, name in enumerate(level_names):
        nm = (name or "").lower()
        if nm in {"date", "datum", "dt"}:
            name_map[i] = "date"
        elif nm in {"wkn", "isin", "ticker", "symbol"}:
            name_map[i] = "wkn"
    if set(name_map.values()) != {"date", "wkn"}:
        # Wenn unbenannt, nehme an: [0]=date, [1]=wkn
        df.index = df.index.set_names(["date", "wkn"])
    else:
        # Reihung anhand name_map
        ordered = ["date", "wkn"]
        current = [name_map.get(i) for i in range(2)]
        if current != ordered:
            df = df.reorder_levels(ordered).sort_index()
        df.index = df.index.set_names(ordered)

    # Datumsnormalisierung (nur Datum, ohne Zeit)
    idx_dates = pd.to_datetime(df.index.get_level_values("date")).date
    df.index = pd.MultiIndex.from_arrays(
        [pd.to_datetime(idx_dates), df.index.get_level_values("wkn")],
        names=["date", "wkn"],
    )

    # Sortierung für stabile Iteration
    df = df.sort_index()
    return df


def instruments_import(filename: str) -> Optional[pd.DataFrame]:
    """
    Liest die Excel-Datei und importiert die ersten vier Spalten (wkn, ticker, instrument_name, default_value)
    in einen DataFrame.
    - wkn (Index) und ticker werden in Kleinbuchstaben umgewandelt.
    - Spalten werden auf ['ticker','instrument_name','default_value'] gesetzt.
    Gibt den DataFrame zurück oder None bei Fehler.
    """
    try:
        if not filename.lower().endswith((".xlsx", ".xls")):
            raise ValueError(f"Die Datei '{filename}' ist keine Excel-Datei.")
        df = pd.read_excel(filename, usecols=[0, 1, 2, 3], index_col=0)
        # Normalize
        if df.index.dtype == "object":
            df.index = df.index.str.strip().str.lower()
        else:
            df.index = df.index.astype(str).str.strip().str.lower()
        if "ticker" not in df.columns:
            # Versuche, die 2. Spalte als 'ticker' zu interpretieren
            df.columns = ["ticker", "instrument_name", "default_value"][: len(df.columns)]
        else:
            # Setze saubere Namen
            df.columns = ["ticker", "instrument_name", "default_value"][: len(df.columns)]
        # ticker lower
        df["ticker"] = df["ticker"].astype(str).str.strip().str.lower()
        return df
    except Exception as e:
        logging.error("Mapping-Import fehlgeschlagen: %s", e)
        return None


def build_wkn_map(df_instr: Optional[pd.DataFrame]) -> Dict[str, str]:
    """Erzeugt ein Dict {wkn_lower: ticker} aus dem Instruments-DataFrame."""
    if df_instr is None or df_instr.empty:
        return {}
    m = df_instr["ticker"].to_dict()
    # Keys sind bereits lower; stelle sicher, dass Werte Strings sind
    return {str(k).strip().lower(): str(v).strip() for k, v in m.items() if pd.notna(v) and str(v).strip()}


# ------------------------------ Hilfsfunktionen -----------------------------

def is_weekend(dt: datetime | pd.Timestamp) -> bool:
    """True, wenn Samstag(5) oder Sonntag(6)."""
    wd = pd.Timestamp(dt).weekday()
    return wd >= 5


def normalize_wkn(wkn: str) -> str:
    """Sanitisiert/vereinheitlicht die WKN als String."""
    return str(wkn).strip()


def wkn_to_yf_symbol(wkn: str, wkn_map: Dict[str, str]) -> Optional[str]:
    """Mappt WKN -> yfinance Symbol über Mapping (falls vorhanden).
    Gibt None zurück, wenn ausgeschlossen.
    """
    if wkn is None:
        return None
    wkn_norm = normalize_wkn(wkn)
    if wkn_norm.lower() in EXCLUDE_WKNS:
        return None
    # Mapping beachten (Case-insensitive)
    sym = wkn_map.get(wkn_norm.lower())
    return sym if sym else wkn_norm


def detect_instrument_issue_date(symbol: str, global_start: pd.Timestamp) -> Optional[pd.Timestamp]:
    """Detects the actual issue/listing date of an instrument by fetching max available history.
    Returns the first date where data is available, or None if no data found.
    Tries to fetch from global_start, but uses yfinance's max period if that fails.
    """
    try:
        t = yf.Ticker(symbol)
        # Try fetching max available history
        hist = t.history(period="max", interval="1d", actions=False, auto_adjust=False)

        if hist.empty:
            # Try with explicit start date
            hist = t.history(start=global_start, interval="1d", actions=False, auto_adjust=False)

        if not hist.empty:
            first_date = pd.to_datetime(hist.index[0]).tz_localize(None).normalize()
            logging.info("Instrument %s: Earliest available data from %s", symbol, first_date.date())
            return first_date
        else:
            logging.warning("Instrument %s: No historical data available", symbol)
            return None
    except Exception as e:
        logging.error("Error detecting issue date for %s: %s", symbol, e)
        return None


def fetch_yf_series(symbol: str, start_date: pd.Timestamp, end_date: pd.Timestamp) -> Tuple[pd.Series, str]:
    """Lädt Tageskurse ('Close' bevorzugt, Fallback 'Adj Close') für Symbol im Intervall [start_date, end_date] (inkl.).
    Gibt (Series, verwendete_Spalte) zurück. Series-Index sind Handelstage als Datum (Timestamp 00:00).
    """
    # yfinance erwartet end EXKLUSIV -> +1 Tag
    yf_start = pd.Timestamp(start_date).tz_localize(None)
    yf_end = pd.Timestamp(end_date + pd.Timedelta(days=1)).tz_localize(None)

    logging.info("Hole yfinance-Daten: %s von %s bis %s", symbol, yf_start.date(), (yf_end - pd.Timedelta(days=1)).date())

    try:
        t = yf.Ticker(symbol)
        hist = t.history(start=yf_start, end=yf_end, interval="1d", actions=False, auto_adjust=False)
    except Exception as e:
        logging.error("yfinance Fehler für %s: %s", symbol, e)
        return pd.Series(dtype="float64"), YF_PREFERRED_COL

    if hist.empty:
        logging.warning("Keine Daten von yfinance für %s im angefragten Intervall.", symbol)
        return pd.Series(dtype="float64"), YF_PREFERRED_COL

    col = YF_PREFERRED_COL if YF_PREFERRED_COL in hist.columns else (YF_FALLBACK_COL if YF_FALLBACK_COL in hist.columns else None)
    if not col:
        logging.warning("Weder '%s' noch '%s' in yfinance-Daten für %s. Verfügbar: %s",
                        YF_PREFERRED_COL, YF_FALLBACK_COL, symbol, list(hist.columns))
        return pd.Series(dtype="float64"), YF_PREFERRED_COL

    # Index auf Datum (ohne Zeit) normalisieren
    s = hist[col].copy()
    s.index = pd.to_datetime(s.index).tz_localize(None).date
    s.index = pd.to_datetime(s.index)  # wieder Timestamp (00:00)
    s.name = "yf_price"
    return s, col


# ------------------------ Vergleich & Korrektur -----------------------------

def compare_and_correct_prices(df: pd.DataFrame, wkn_map: Dict[str, str]) -> Tuple[pd.DataFrame, pd.DataFrame]:
    """Vergleicht Preise im Input-DF mit yfinance pro WKN über deren Datumsbereich.
    Mit ENABLE_BACKFILL=True werden auch historische Daten vor dem ersten vorhandenen Datum nachgeladen.

    Rückgabe:
      - diffs: DataFrame (date, wkn, old_price, yf_price, diff, pct_diff, yf_symbol, yf_col, backfill_flag)
      - corrected: DataFrame gleicher Struktur wie Input; Werte werden
                   * aktualisiert, wenn beide vorhanden und abweichend
                   * aufgefüllt, wenn old_price NaN und yfinance vorhanden
                   * BACKFILLED: neue Datums-Zeilen für historische Daten (wenn ENABLE_BACKFILL=True)
    """
    deviations: List[dict] = []

    # sichere Kopien
    df_local = df.copy()
    df_local[PRICE_COL] = pd.to_numeric(df_local[PRICE_COL], errors="coerce")

    corrected = df_local.copy()

    # Determine global date range for backfilling
    global_start_date = None
    global_end_date = None

    if ENABLE_BACKFILL:
        all_dates = pd.to_datetime(df_local.index.get_level_values("date"))
        global_end_date = all_dates.max().normalize()

        if BACKFILL_START_DATE:
            global_start_date = pd.to_datetime(BACKFILL_START_DATE).normalize()
            logging.info("Backfill aktiviert: Start-Datum aus Konfiguration: %s", global_start_date.date())
        else:
            global_start_date = all_dates.min().normalize()
            logging.info("Backfill aktiviert: Start-Datum automatisch ermittelt: %s", global_start_date.date())

    # List to collect new rows for backfilling
    new_rows: List[dict] = []

    # Iteration über WKNs
    for wkn, df_wkn in df_local.groupby(level="wkn"):
        symbol = wkn_to_yf_symbol(wkn, wkn_map)
        if symbol is None:
            logging.info("Überspringe WKN '%s' (ausgeschlossen).", wkn)
            continue

        # Datumsbereich für diese WKN
        dates = pd.to_datetime(df_wkn.index.get_level_values("date"))
        # Wochenenden vorab filtern
        dates_bd = dates[~dates.map(is_weekend)]
        if dates_bd.empty:
            continue

        wkn_start_date = dates_bd.min().normalize()
        wkn_end_date = dates_bd.max().normalize()

        # Decide on fetch range
        fetch_start = wkn_start_date
        fetch_end = wkn_end_date
        needs_backfill = False

        if ENABLE_BACKFILL and global_start_date and wkn_start_date > global_start_date:
            # This WKN has limited history - attempt backfill
            needs_backfill = True
            logging.info("WKN '%s': Limitierte Historie erkannt (Start: %s vs Global: %s)",
                        wkn, wkn_start_date.date(), global_start_date.date())

            # Detect actual instrument issue date
            issue_date = detect_instrument_issue_date(symbol, global_start_date)
            if issue_date and issue_date < wkn_start_date:
                fetch_start = max(issue_date, global_start_date)
                logging.info("WKN '%s': Backfill von %s bis %s", wkn, fetch_start.date(), wkn_start_date.date())
            elif issue_date:
                logging.info("WKN '%s': Instrument-Start (%s) entspricht oder liegt nach vorhandenem Start",
                            wkn, issue_date.date())
                needs_backfill = False
            else:
                logging.warning("WKN '%s': Konnte Instrument-Start nicht ermitteln, kein Backfill", wkn)
                needs_backfill = False

        # yfinance-Daten holen
        yf_series, used_col = fetch_yf_series(symbol, fetch_start, fetch_end)
        if yf_series.empty:
            continue

        # Check if we got sufficient data (validates that ticker mapping is correct)
        if len(yf_series) < MIN_TRADING_DAYS_THRESHOLD:
            logging.warning("WKN '%s': Nur %d Handelstage gefunden (< %d), möglicherweise falsches Ticker-Mapping",
                           wkn, len(yf_series), MIN_TRADING_DAYS_THRESHOLD)
            continue

        # Separate backfill dates from existing dates
        existing_dates = set(df_wkn.index.get_level_values("date").date)
        yf_dates = set(yf_series.index.date)

        backfill_dates = set()
        if needs_backfill:
            backfill_dates = yf_dates - existing_dates
            if backfill_dates:
                logging.info("WKN '%s': %d neue historische Datums-Zeilen zum Backfill",
                            wkn, len(backfill_dates))

        # Process existing dates (update/correct)
        df_wkn_local = df_wkn.reset_index()
        df_wkn_local["date"] = pd.to_datetime(df_wkn_local["date"]).dt.normalize()

        # Filter to dates that exist in yfinance
        df_wkn_local = df_wkn_local[df_wkn_local["date"].isin(yf_series.index)]

        if not df_wkn_local.empty:
            merged = df_wkn_local.merge(
                yf_series.rename("yf_price"),
                left_on="date",
                right_index=True,
                how="left",
            )

            # Vergleichen & ggf. korrigieren/auffüllen
            for _, row in merged.iterrows():
                old_price = row[PRICE_COL]
                new_price = row["yf_price"]
                date_key = pd.to_datetime(row["date"]).normalize()

                if pd.isna(old_price) and pd.notna(new_price):
                    # Auffüllen
                    deviations.append({
                        "date": date_key.date(),
                        "wkn": wkn,
                        "old_price": None,
                        "yf_price": float(new_price),
                        "diff": None,
                        "pct_diff": None,
                        "yf_symbol": symbol,
                        "yf_col": used_col,
                        "backfill": False,
                    })
                    corrected.loc[(date_key, wkn), PRICE_COL] = float(new_price)
                    continue

                if pd.notna(old_price) and pd.notna(new_price):
                    diff = float(new_price) - float(old_price)
                    if math.isfinite(diff) and abs(diff) > ABS_TOL:
                        pct = None
                        if float(old_price) != 0.0 and math.isfinite(float(old_price)):
                            pct = diff / float(old_price)
                        deviations.append({
                            "date": date_key.date(),
                            "wkn": wkn,
                            "old_price": float(old_price),
                            "yf_price": float(new_price),
                            "diff": float(diff),
                            "pct_diff": pct,
                            "yf_symbol": symbol,
                            "yf_col": used_col,
                            "backfill": False,
                        })
                        # Aktualisieren
                        corrected.loc[(date_key, wkn), PRICE_COL] = float(new_price)

                # Fall: yfinance fehlt oder beide fehlen → keine Aktion

        # Handle backfill dates (new rows)
        if backfill_dates:
            for date_val in backfill_dates:
                date_ts = pd.to_datetime(date_val).normalize()
                if date_ts in yf_series.index:
                    new_price = yf_series.loc[date_ts]
                    if pd.notna(new_price):
                        # Add to deviations for reporting
                        deviations.append({
                            "date": date_val,
                            "wkn": wkn,
                            "old_price": None,
                            "yf_price": float(new_price),
                            "diff": None,
                            "pct_diff": None,
                            "yf_symbol": symbol,
                            "yf_col": used_col,
                            "backfill": True,
                        })
                        # Add to new_rows for later concat
                        new_rows.append({
                            "date": date_ts,
                            "wkn": wkn,
                            PRICE_COL: float(new_price),
                        })

    # diffs DataFrame
    if deviations:
        diffs = pd.DataFrame(deviations).sort_values(["date", "wkn"]).reset_index(drop=True)
    else:
        cols = ["date", "wkn", "old_price", "yf_price", "diff", "pct_diff", "yf_symbol", "yf_col", "backfill"]
        diffs = pd.DataFrame(columns=cols)

    # Add new backfilled rows to corrected DataFrame
    if new_rows:
        logging.info("Füge %d neue historische Zeilen zum korrigierten DataFrame hinzu", len(new_rows))
        new_df = pd.DataFrame(new_rows)
        new_df = new_df.set_index(["date", "wkn"])
        # Combine with existing corrected data
        corrected = pd.concat([corrected, new_df])
        # Remove any duplicates (prefer existing data)
        corrected = corrected[~corrected.index.duplicated(keep='first')]

    # corrected DataFrame Struktur und Sortierung wahren
    corrected = corrected.sort_index()
    return diffs, corrected


# ------------------------------ Exporte -------------------------------------

def export_diffs_to_excel(diffs: pd.DataFrame, out_path: Path) -> None:
    """Exportiert die Abweichungen nach Excel. Formatiert pct_diff als Prozent falls möglich."""
    with pd.ExcelWriter(out_path, engine="openpyxl") as writer:
        diffs.to_excel(writer, index=False, sheet_name="differences")
        try:
            ws = writer.sheets["differences"]
            header = [cell.value for cell in next(ws.iter_rows(min_row=1, max_row=1))]
            if "pct_diff" in header:
                col_idx = header.index("pct_diff") + 1  # 1-basiert
                for col in ws.iter_cols(min_col=col_idx, max_col=col_idx, min_row=2, max_row=ws.max_row):
                    for c in col:
                        c.number_format = "0.00%"
        except Exception:
            pass


def save_corrected_parquet(corrected: pd.DataFrame, original_path: Path) -> Path:
    """Speichert den korrigierten DataFrame als Parquet neben das Original mit Suffix _new."""
    new_path = original_path.with_name(f"{original_path.stem}_new{original_path.suffix}")
    corrected.to_parquet(new_path)
    return new_path


# --------------------------------- main -------------------------------------

def main(argv: Optional[List[str]] = None) -> int:
    setup_logging()
    try:
        # 1) Parquet auswählen (Default: Skript-Ordner + prices.parquet)
        parquet_path = select_parquet_path()
        logging.info("Eingabedatei: %s", parquet_path)

        # 2) OPTIONAL: Excel-Mapping auswählen (Default: UNC-Pfad)
        excel_path = select_excel_mapping_path()
        if excel_path and Path(excel_path).exists():
            logging.info("Mapping-Datei: %s", excel_path)
            instr_df = instruments_import(str(excel_path))
        else:
            if excel_path:
                logging.warning("Gewählte Mapping-Datei nicht gefunden: %s", excel_path)
            else:
                logging.info("Kein Mapping gewählt/gefunden. WKN wird als yfinance-Symbol verwendet.")
            instr_df = None

        wkn_map = build_wkn_map(instr_df)

        # 3) Parquet laden
        df = read_input_parquet(parquet_path)

        # 4) Vergleichen & korrigieren/auffüllen
        diffs, corrected = compare_and_correct_prices(df, wkn_map)

        # 5) Exporte
        excel_path_out = parquet_path.parent / "new prices.xlsx"
        export_diffs_to_excel(diffs, excel_path_out)
        parquet_new_path = save_corrected_parquet(corrected, parquet_path)

        logging.info("Abweichungen: %d Zeilen -> %s", len(diffs), excel_path_out)
        logging.info("Korrigierter Parquet geschrieben: %s", parquet_new_path)

        print(f"Fertig.\nExcel:   {excel_path_out}\nParquet: {parquet_new_path}")
        return 0
    except Exception as e:
        logging.exception("Fehler: %s", e)
        return 1


if __name__ == "__main__":
    raise SystemExit(main())
