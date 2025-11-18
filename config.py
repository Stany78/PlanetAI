from __future__ import annotations
import os
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

# CSV OMI (dopo estrazione)
OMI_CSV_PATH = os.path.join(OMI_DIR, "QI_20251_VALORI.csv")
OMI_ZONE_CSV_PATH = os.path.join(OMI_DIR, "QI_20251_ZONE.csv")

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
# ESTRAZIONE AUTOMATICA DEI DATI OMI
# ==========================================

def ensure_omi_unzipped() -> None:
    """
    Se la cartella Omi/ non è inizializzata, estrae automaticamente
    tutti gli archivi Omi_*.zip presenti nella root del progetto.
    Funziona sia in locale sia su Streamlit Cloud.
    """
    _ensure_dir(OMI_DIR)

    # Se il CSV principale esiste già, assumiamo che sia tutto estratto
    if os.path.exists(OMI_CSV_PATH):
        if DEBUG_MODE:
            print("[OMI] Cartella Omi già inizializzata. Nessuna estrazione necessaria.")
        return

    zip_files = sorted(glob.glob(OMI_ZIP_GLOB))

    if not zip_files:
        print("[OMI][WARN] Nessun archivio Omi_*.zip trovato. Dati OMI non disponibili.")
        return

    if DEBUG_MODE:
        print(f"[OMI] Estraggo {len(zip_files)} archivi...")

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

# Stringa informativa sui dati OMI usati
OMI_DATA_INFO = "Dati OMI QI_20251 – Semestre 2025/1"
