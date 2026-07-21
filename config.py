from __future__ import annotations
# ==========================================
# CONFIG - v2.0 (auto-detect semestre OMI)
# Rileva automaticamente QI_*_VALORI.csv / QI_*_ZONE.csv
# dal contenuto degli archivi Omi_*.zip: per aggiornare i dati
# basta sostituire gli zip, nessuna modifica al codice.
# ==========================================
import os
import re
import zipfile
import glob

# ==========================================
# MODALITÀ DEBUG
# ==========================================

DEBUG_MODE = True  # True per log più verbosi


# ==========================================
# PERCORSI BASE PROGETTO
# ==========================================

# Cartella del progetto (es. Planet AI)
BASE_DIR = os.path.abspath(os.path.dirname(__file__))

# Cartella per report / export
REPORTS_DIR = os.path.join(BASE_DIR, "reports")

# Cartella dove devono stare i dati OMI estratti
OMI_DIR = os.path.join(BASE_DIR, "Omi")

# Pattern per gli archivi compressi Omi_*.zip (Omi_1.zip, Omi_2.zip, ...)
OMI_ZIP_GLOB = os.path.join(BASE_DIR, "Omi_*.zip")

# Cartella che contiene tutti i KML OMI (A001.kml ... M437.kml)
OMI_KML_PATH = OMI_DIR


# ==========================================
# UTILITY FILESYSTEM
# ==========================================

def _ensure_dir(path: str) -> None:
    """Crea la directory se non esiste."""
    if not os.path.exists(path):
        os.makedirs(path)


# ==========================================
# RILEVAMENTO AUTOMATICO SEMESTRE OMI
# ==========================================

def _csv_expected_from_zips() -> str | None:
    """
    Ispeziona gli archivi Omi_*.zip e restituisce il nome del file
    QI_*_VALORI.csv contenuto (es. 'QI_20252_VALORI.csv').
    Restituisce None se non trovato.
    """
    for zpath in sorted(glob.glob(OMI_ZIP_GLOB)):
        try:
            with zipfile.ZipFile(zpath, "r") as zf:
                for name in zf.namelist():
                    base = os.path.basename(name)
                    if re.fullmatch(r"QI_\d+_VALORI\.csv", base, re.IGNORECASE):
                        return base
        except Exception as e:
            print(f"[OMI][WARN] Impossibile ispezionare {os.path.basename(zpath)}: {e}")
    return None


def _find_extracted_csv(suffix: str) -> str | None:
    """
    Cerca in OMI_DIR un file QI_*_{suffix}.csv già estratto.
    Se ce n'è più di uno (semestri diversi), sceglie il più recente
    (l'ordinamento lessicografico QI_YYYYS funziona: 20251 < 20252 < 20261).
    """
    matches = sorted(glob.glob(os.path.join(OMI_DIR, f"QI_*_{suffix}.csv")))
    return matches[-1] if matches else None


# ==========================================
# ESTRAZIONE AUTOMATICA DEI DATI OMI
# ==========================================

def ensure_omi_unzipped() -> None:
    """
    Estrae gli archivi Omi_*.zip se i dati del semestre che contengono
    non sono ancora presenti in Omi/.
    A differenza della versione precedente, il controllo è fatto sul
    NOME EFFETTIVO del CSV dentro gli zip: se sostituisci gli zip con
    un nuovo semestre, l'estrazione riparte automaticamente.
    """
    _ensure_dir(OMI_DIR)

    expected_csv = _csv_expected_from_zips()

    if expected_csv is None:
        # Nessuno zip (o zip senza CSV): ok solo se un CSV è già estratto
        if _find_extracted_csv("VALORI"):
            if DEBUG_MODE:
                print("[OMI] Nessun archivio zip, uso i CSV già estratti.")
        else:
            print("[OMI][WARN] Nessun archivio Omi_*.zip trovato. Dati OMI non disponibili.")
        return

    # Il CSV del semestre contenuto negli zip è già estratto? → niente da fare
    if os.path.exists(os.path.join(OMI_DIR, expected_csv)):
        if DEBUG_MODE:
            print(f"[OMI] Dati {expected_csv} già estratti. Nessuna estrazione necessaria.")
        return

    zip_files = sorted(glob.glob(OMI_ZIP_GLOB))

    if DEBUG_MODE:
        print(f"[OMI] Nuovo semestre rilevato ({expected_csv}). Estraggo {len(zip_files)} archivi...")

    for zpath in zip_files:
        zname = os.path.basename(zpath)
        try:
            if DEBUG_MODE:
                print(f"[OMI] Estraggo: {zname}")
            with zipfile.ZipFile(zpath, "r") as zip_ref:
                zip_ref.extractall(OMI_DIR)
        except Exception as e:
            print(f"[OMI][ERROR] Errore durante l'estrazione di {zname}: {e}")

    if DEBUG_MODE:
        print("[OMI] Estrazione completata.")


# ==========================================
# RISOLUZIONE PATH CSV (dopo eventuale estrazione)
# ==========================================

# L'estrazione viene garantita già all'import del modulo, così le
# costanti OMI_CSV_PATH / OMI_ZONE_CSV_PATH puntano sempre al file giusto
# anche su deploy pulito (Streamlit Cloud).
ensure_omi_unzipped()

_valori = _find_extracted_csv("VALORI")
_zone = _find_extracted_csv("ZONE")

OMI_CSV_PATH = _valori or os.path.join(OMI_DIR, "QI_MISSING_VALORI.csv")
OMI_ZONE_CSV_PATH = _zone or os.path.join(OMI_DIR, "QI_MISSING_ZONE.csv")


def _semestre_label(csv_path: str) -> str:
    """'QI_20252_VALORI.csv' → '2025/2'; fallback 'N/D'."""
    m = re.search(r"QI_(\d{4})(\d)_VALORI", os.path.basename(csv_path))
    if m:
        return f"{m.group(1)}/{m.group(2)}"
    return "N/D"


# Stringa informativa sui dati OMI usati (auto-generata dal nome file)
OMI_DATA_INFO = f"Dati OMI – Semestre {_semestre_label(OMI_CSV_PATH)}"

if DEBUG_MODE:
    print(f"[OMI] {OMI_DATA_INFO} | CSV: {os.path.basename(OMI_CSV_PATH)}")


# ==========================================
# PARAMETRI DI MODELLO USATI DA agent_core
# ==========================================

# Numero minimo di comparabili usati per stimare il "nuovo"
MIN_COMPARABLE_NUOVO = 5

# Range di spread sul valore OMI per stimare il nuovo (es. 15%–35%)
SPREAD_NUOVO_MIN = 0.15
SPREAD_NUOVO_MAX = 0.35

# Coordinate di fallback (es. centro di Como) se geocoding fallisce
FALLBACK_COORDINATE = (45.8081, 9.0852)

# Fattore di default da applicare ai valori OMI per stimare il nuovo
FATTORE_DEFAULT_SU_OMI = 1.25

# Stringa informativa sui dati OMI usati: vedi OMI_DATA_INFO sopra (auto-generata)