import os
import pandas as pd
from tkinter import Tk, filedialog
from datetime import datetime

# Funktion zur Auswahl einer Parquet-Datei
def select_parquet_file():
    current_dir = os.path.dirname(os.path.abspath(__file__))
    Tk().withdraw()  # Tkinter-Fenster verbergen
    file_path = filedialog.askopenfilename(
        initialdir=current_dir, 
        title="Wähle eine Parquet-Datei",
        filetypes=[("Parquet files", "*.parquet")]
    )
    return file_path

# Funktion zum Einlesen und Analysieren des Datumsbereichs in der Parquet-Datei
def load_and_analyze_dates(file_path):
    df = pd.read_parquet(file_path)
    if 'date' in df.index.names or 'date' in df.columns:
        date_column = df.index.get_level_values('date') if 'date' in df.index.names else df['date']
        min_date, max_date = date_column.min(), date_column.max()
        print(f"\nÄltestes Datum im Datensatz: {min_date}")
        print(f"Jüngstes Datum im Datensatz: {max_date}")
        return df, min_date, max_date
    else:
        print("Fehler: Keine Datumsspalte im Datensatz gefunden.")
        return None, None, None

# Funktion zur Durchführung der Löschoperation
def delete_records(df, file_path, cutoff_date):
    # Filtere Datensätze ab dem angegebenen Datum
    records_to_delete = df[df.index.get_level_values('date') >= cutoff_date]
    records_count = len(records_to_delete)
    
    if records_count > 0:
        # Sicherheitsabfrage
        confirm = input(f"{records_count} Datensätze ab dem {cutoff_date} werden gelöscht. Möchten Sie fortfahren? (ja/nein): ")
        
        if confirm.lower() == 'ja':
            # Lösche Datensätze und speichere unter dem gleichen Dateinamen
            df = df[df.index.get_level_values('date') < cutoff_date]
            df.to_parquet(file_path)
            print(f"Datensatz wurde erfolgreich gekürzt und unter '{file_path}' gespeichert.")
        else:
            print("Vorgang abgebrochen.")
    else:
        print("Keine Datensätze zum Löschen gefunden.")

# Hauptprogramm
if __name__ == "__main__":
    # Parquet-Datei auswählen
    file_path = select_parquet_file()
    if file_path:
        # Parquet-Datei laden und Datumsbereich analysieren
        df, min_date, max_date = load_and_analyze_dates(file_path)
        
        if df is not None:
            # Benutzereingabe für das Löschdatum
            try:
                cutoff_date = input("Ab welchem Datum sollen Datensätze gelöscht werden? (YYYY-MM-DD): ")
                cutoff_date = pd.Timestamp(datetime.strptime(cutoff_date, "%Y-%m-%d"))
                
                # Prüfen, ob das Datum im Datumsbereich liegt
                if min_date <= cutoff_date <= max_date:
                    delete_records(df, file_path, cutoff_date)
                else:
                    print(f"Das eingegebene Datum {cutoff_date} liegt nicht im Datumsbereich des Datensatzes.")
            except ValueError:
                print("Ungültiges Datum eingegeben. Bitte im Format YYYY-MM-DD eingeben.")
    else:
        print("Keine Datei ausgewählt.")
