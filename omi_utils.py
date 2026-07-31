from __future__ import annotations

import os
import glob
import re
import xml.etree.ElementTree as ET
from dataclasses import dataclass
from typing import List, Optional, Dict, Tuple

import pandas as pd

from config import (
    ensure_omi_unzipped,
    OMI_CSV_PATH,
    OMI_ZONE_CSV_PATH,
    OMI_KML_PATH,
    DEBUG_MODE,
)


# =====================================================
# Dataclass risultato
# =====================================================


@dataclass
class OMIQuotazione:
    comune: str
    provincia: str
    zona_codice: str
    zona_descrizione: str
    val_min_mq: Optional[float]
    val_med_mq: Optional[float]
    val_max_mq: Optional[float]


# =====================================================
# Cache in memoria
# =====================================================

# ogni elemento della lista:
# {"zona": "B2", "comune": "COMO", "provincia": "CO", "polygon": [(lon, lat), ...]}
_omi_polygons: List[Dict] = []
_omi_valori_df: Optional[pd.DataFrame] = None
_omi_zone_df: Optional[pd.DataFrame] = None
_omi_cache_ready: bool = False


# =====================================================
# Utility varie
# =====================================================


def _safe_float(x) -> Optional[float]:
    """
    Converte stringhe tipo "1,2" / "1.2" in float.
    Restituisce None se non convertibile.
    """
    if x is None:
        return None
    if isinstance(x, (int, float)):
        return float(x)
    s = str(x).strip()
    if not s or s.upper() == "NA":
        return None

    # Gestione migliaia + virgola decimale (es: "1.234,56")
    if s.count(",") == 1 and s.count(".") > 1:
        s = s.replace(".", "").replace(",", ".")
    else:
        s = s.replace(",", ".")

    try:
        return float(s)
    except ValueError:
        return None


def _parse_kml_file(path: str) -> List[Dict]:
    """
    Parsing di un singolo file KML.
    Restituisce una lista di poligoni:
    [
        {"zona": "B1", "comune": "COMO", "provincia": "CO", "polygon": [(lon, lat), ...]},
        ...
    ]
    """
    ns = {"k": "http://www.opengis.net/kml/2.2"}

    try:
        tree = ET.parse(path)
        root = tree.getroot()
    except Exception as e:
        if DEBUG_MODE:
            print(f"[OMI][KML] Errore parsing {os.path.basename(path)}: {e}")
        return []

    # Esempio Document name:
    # "COMO (CO) Anno/Semestre 2025/1 generato il ..."
    comune = ""
    prov = ""

    doc_name_el = root.find(".//k:Document/k:name", ns)
    if doc_name_el is not None and doc_name_el.text:
        txt = doc_name_el.text.strip()
        m = re.match(r"^(.*?)\s+\((..)\)", txt)
        if m:
            comune = m.group(1).strip().upper()
            prov = m.group(2).strip().upper()

    polygons: List[Dict] = []

    for pm in root.findall(".//k:Placemark", ns):
        # 1) Codice zona
        zona_code: Optional[str] = None

        # Preferisco ExtendedData/CODZONA
        for data in pm.findall(".//k:ExtendedData/k:Data", ns):
            if data.attrib.get("name") == "CODZONA":
                v = data.find("k:value", ns)
                if v is not None and v.text:
                    zona_code = v.text.strip().upper()
                    break

        # Fallback: dal <name> tipo "...Zona OMI B1"
        if not zona_code:
            name_el = pm.find("k:name", ns)
            if name_el is not None and name_el.text:
                m = re.search(r"ZONA\s+OMI\s+([A-Z0-9]+)", name_el.text.upper())
                if m:
                    zona_code = m.group(1)

        if not zona_code:
            continue

        # 2) Tutti i poligoni (Polygon o MultiGeometry/Polygon)
        for poly_el in pm.findall(".//k:Polygon", ns):
            coords_el = poly_el.find(".//k:outerBoundaryIs/k:LinearRing/k:coordinates", ns)
            if coords_el is None or not coords_el.text:
                continue

            coords_text = coords_el.text.strip()
            ring: List[Tuple[float, float]] = []
            for part in re.split(r"\s+", coords_text):
                if not part:
                    continue
                bits = part.split(",")
                if len(bits) < 2:
                    continue
                try:
                    lon = float(bits[0])
                    lat = float(bits[1])
                except ValueError:
                    continue
                ring.append((lon, lat))

            if len(ring) >= 3:
                polygons.append(
                    {
                        "zona": zona_code,
                        "comune": comune,
                        "provincia": prov,
                        "polygon": ring,
                    }
                )

    if DEBUG_MODE:
        print(f"[OMI] {os.path.basename(path)}: trovate {len(polygons)} zone")

    return polygons


def _point_in_polygon(x: float, y: float, poly: List[Tuple[float, float]]) -> bool:
    """
    Test standard "ray casting" per punto in poligono.
    poly è una lista di (x=lon, y=lat).
    """
    inside = False
    n = len(poly)
    if n < 3:
        return False

    for i in range(n):
        x1, y1 = poly[i]
        x2, y2 = poly[(i + 1) % n]

        intersects = ((y1 > y) != (y2 > y)) and (
            x < (x2 - x1) * (y - y1) / (y2 - y1 + 1e-12) + x1
        )
        if intersects:
            inside = not inside

    return inside


# =====================================================
# Caricamento CSV OMI (valori + zone)
# =====================================================


def _load_omi_csvs() -> None:
    """
    Carica in cache:
    - QI_20251_VALORI.csv
    - QI_20251_ZONE.csv (se esiste)
    """
    global _omi_valori_df, _omi_zone_df

    if _omi_valori_df is None and os.path.exists(OMI_CSV_PATH):
        _omi_valori_df = pd.read_csv(
            OMI_CSV_PATH,
            sep=";",
            skiprows=1,  # salta la riga di intestazione "umana"
            dtype=str,
        )
        if DEBUG_MODE:
            print("[OMI] Valori CSV caricato:", _omi_valori_df.shape)

    if _omi_zone_df is None and os.path.exists(OMI_ZONE_CSV_PATH):
        _omi_zone_df = pd.read_csv(
            OMI_ZONE_CSV_PATH,
            sep=";",
            skiprows=1,
            dtype=str,
        )
        if DEBUG_MODE:
            print("[OMI] Zone CSV caricato:", _omi_zone_df.shape)


# =====================================================
# Caricamento poligoni KML
# =====================================================


def _load_omi_polygons() -> None:
    """
    Carica tutti i poligoni OMI dai KML in OMI_KML_PATH.
    Funziona anche se i KML sono in sottocartelle (estrazione da Omi_*.zip).
    """
    global _omi_polygons

    if _omi_polygons:
        return

    ensure_omi_unzipped()

    pattern = os.path.join(OMI_KML_PATH, "**", "*.kml")
    kml_files = sorted(glob.glob(pattern, recursive=True))

    if DEBUG_MODE:
        print(f"[OMI] Carico KML da {pattern} - n file: {len(kml_files)}")

    polygons: List[Dict] = []
    for path in kml_files:
        polygons.extend(_parse_kml_file(path))

    _omi_polygons = polygons

    if DEBUG_MODE:
        print(f"[OMI] Totale poligoni caricati: {len(_omi_polygons)}")


# =====================================================
# Da zona OMI → valori €/mq
# =====================================================


def _get_valori_for_zona(
    zona: str,
    comune_kml: Optional[str] = None,
    provincia_kml: Optional[str] = None,
) -> Optional[OMIQuotazione]:
    """
    Dato il codice di zona (es. 'B2') e Comune/Provincia dal KML,
    estrae i valori dal CSV OMI.

    REGOLA FONDAMENTALE (fix v2): la ricerca NON attraversa MAI i confini
    del comune. Il vecchio fallback "Solo Zona" restituiva la prima zona
    omonima d'Italia (es. R4 di Castel Volturno -> Montiglio Monferrato AT).

    Logica:
      1. Righe con Zona + Comune (+ Provincia se disponibile)
      2. Tipologia: preferisci Abitazioni civili/signorili;
         se assenti, altre tipologie residenziali dello STESSO comune
         (economiche, ville e villini)
      3. Se nessuna quotazione residenziale esiste per quella zona/comune,
         restituisce la zona con valori None (onesto: "quotazione non
         disponibile"), MAI dati di un altro comune.
    """
    _load_omi_csvs()

    if _omi_valori_df is None or _omi_valori_df.empty:
        return None

    df = _omi_valori_df
    dfz = _omi_zone_df

    zona_up = zona.strip().upper()

    def _norm(s: Optional[str]) -> str:
        return "" if s is None else s.strip().upper()

    comune_up = _norm(comune_kml)
    prov_up = _norm(provincia_kml)

    # ---- 1) Candidati: SEMPRE vincolati al comune del KML ----
    mask = df["Zona"].str.upper().eq(zona_up)
    if comune_up:
        mask = mask & df["Comune_descrizione"].str.upper().eq(comune_up)
    if prov_up:
        mask = mask & df["Prov"].str.upper().eq(prov_up)
    cand = df.loc[mask]

    # Riprova senza provincia (possibili discrepanze di sigla), comune sempre obbligatorio
    if cand.empty and comune_up and prov_up:
        mask = df["Zona"].str.upper().eq(zona_up) & df[
            "Comune_descrizione"
        ].str.upper().eq(comune_up)
        cand = df.loc[mask]

    # ---- 2) Selezione tipologia (preferenza, non filtro rigido) ----
    TIPOLOGIE_PREFERITE = ["ABITAZIONI CIVILI", "ABITAZIONI SIGNORILI"]
    TIPOLOGIE_RESIDENZIALI = TIPOLOGIE_PREFERITE + [
        "ABITAZIONI DI TIPO ECONOMICO",
        "VILLE E VILLINI",
    ]

    rows = cand
    if not cand.empty and "Descr_Tipologia" in cand.columns:
        tip = cand["Descr_Tipologia"].str.upper()
        pref = cand[tip.isin(TIPOLOGIE_PREFERITE)]
        if pref.empty:
            pref = cand[tip.isin(TIPOLOGIE_RESIDENZIALI)]
            if DEBUG_MODE and not pref.empty:
                print(
                    f"[OMI] Zona {zona_up} ({comune_up}): nessuna abitazione "
                    f"civile/signorile, uso altre tipologie residenziali."
                )
        rows = pref  # se vuoto: zona senza quotazioni residenziali

    # ---- Descrizione zona (dal CSV zone, vincolata al comune) ----
    zona_descr = f"Zona OMI {zona_up}"
    if dfz is not None and not dfz.empty:
        mask_z = dfz["Zona"].str.upper().eq(zona_up)
        if comune_up:
            mask_z = mask_z & dfz["Comune_descrizione"].str.upper().eq(comune_up)
        rz = dfz.loc[mask_z]
        if not rz.empty:
            raw_descr = str(rz["Zona_Descr"].iloc[0] or "")
            clean_descr = raw_descr.strip().strip("'").strip().title()
            clean_descr = clean_descr.replace(" :", " \u2013 ").replace(": ", " \u2013 ")
            zona_descr = clean_descr

    # ---- 3) Nessuna quotazione residenziale: zona onesta, valori None ----
    if rows.empty:
        if DEBUG_MODE:
            print(
                f"[OMI] Zona {zona_up} di {comune_up or '?'} ({prov_up or '?'}): "
                f"nessuna quotazione residenziale disponibile nel CSV. "
                f"NESSUN fallback su altri comuni."
            )
        return OMIQuotazione(
            comune=(comune_kml or "").title() or "N/D",
            provincia=prov_up or "N/D",
            zona_codice=zona_up,
            zona_descrizione=zona_descr,
            val_min_mq=None,
            val_med_mq=None,
            val_max_mq=None,
        )

    comune = rows["Comune_descrizione"].iloc[0].title()
    provincia = rows["Prov"].iloc[0].upper()

    vals_min = rows["Compr_min"].apply(_safe_float).dropna().tolist()
    vals_max = rows["Compr_max"].apply(_safe_float).dropna().tolist()

    if not vals_min or not vals_max:
        val_min = val_med = val_max = None
    else:
        val_min = min(vals_min)
        val_max = max(vals_max)
        val_med = (val_min + val_max) / 2.0

    return OMIQuotazione(
        comune=comune,
        provincia=provincia,
        zona_codice=zona_up,
        zona_descrizione=zona_descr,
        val_min_mq=val_min,
        val_med_mq=val_med,
        val_max_mq=val_max,
    )


# =====================================================
# API pubbliche
# =====================================================


def warmup_omi_cache() -> None:
    """
    Pre-carica KML + CSV in memoria. Da chiamare una sola volta all'avvio.
    """
    global _omi_cache_ready
    if _omi_cache_ready:
        return

    ensure_omi_unzipped()
    _load_omi_csvs()
    _load_omi_polygons()

    _omi_cache_ready = True

    if DEBUG_MODE:
        print("[OMI] Cache OMI inizializzata.")


def get_quotazione_omi_da_coordinate(
    lat: float,
    lon: float,
) -> Optional[OMIQuotazione]:
    """
    Trova la zona OMI che contiene il punto (lat, lon) e
    restituisce una OMIQuotazione con min/med/max €/mq.
    """
    warmup_omi_cache()

    # ATTENZIONE: KML usa (lon, lat)
    for poly in _omi_polygons:
        if _point_in_polygon(lon, lat, poly["polygon"]):
            if DEBUG_MODE:
                print(
                    f"[OMI] Coordinate ({lat:.6f}, {lon:.6f}) "
                    f"→ zona {poly['zona']} (comune KML: {poly['comune']} {poly['provincia']})"
                )
            return _get_valori_for_zona(
                zona=poly["zona"],
                comune_kml=poly["comune"],
                provincia_kml=poly["provincia"],
            )

    if DEBUG_MODE:
        print(f"[OMI] Nessuna zona trovata per coordinate ({lat}, {lon})")

    return None