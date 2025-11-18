from __future__ import annotations

import io
from datetime import datetime
from typing import Optional

import pandas as pd
import streamlit as st

from config import ensure_omi_unzipped, OMI_DATA_INFO
from agent_core import geocode_indirizzo
from omi_utils import warmup_omi_cache, get_quotazione_omi_da_coordinate

# ---------------------------------------------------------
# Opzionale: report Word, se python-docx è installato
# ---------------------------------------------------------
try:  # pragma: no cover - opzionale
    from docx import Document  # type: ignore

    HAVE_DOCX = True
except Exception:  # pragma: no cover - se la libreria non c'è (es. su Streamlit Cloud)
    HAVE_DOCX = False


# ---------------------------------------------------------
# Utility
# ---------------------------------------------------------
def fmt_euro(value: Optional[float]) -> str:
    if value is None:
        return "-"
    try:
        return f"{value:,.0f} €".replace(",", ".")
    except Exception:
        return f"{value} €"


@st.cache_resource(show_spinner="📦 Caricamento dati OMI (solo il primo avvio è più lento)...")
def init_omi_data() -> bool:
    """
    Estrae gli archivi OMI se necessario e inizializza la cache in memoria.
    """
    ensure_omi_unzipped()
    warmup_omi_cache()
    return True


def build_word_report(
    *,
    comune: str,
    indirizzo: str,
    superficie: float,
    lat: float,
    lon: float,
    omi,
    val_min_mq: float,
    val_med_mq: float,
    val_max_mq: float,
) -> io.BytesIO:
    """
    Crea un semplice report Word basato SOLO sui dati OMI.
    Ritorna un buffer BytesIO pronto per essere scaricato.
    """
    document = Document()

    document.add_heading("Report di valutazione OMI", level=1)

    document.add_paragraph(
        f"Data generazione report: {datetime.now().strftime('%d/%m/%Y %H:%M')}"
    )

    document.add_heading("1. Dati immobile", level=2)
    p = document.add_paragraph()
    p.add_run("Comune: ").bold = True
    p.add_run(comune)
    p = document.add_paragraph()
    p.add_run("Indirizzo: ").bold = True
    p.add_run(indirizzo)
    p = document.add_paragraph()
    p.add_run("Superficie commerciale: ").bold = True
    p.add_run(f"{superficie:.0f} m²")
    p = document.add_paragraph()
    p.add_run("Coordinate: ").bold = True
    p.add_run(f"{lat:.6f}, {lon:.6f}")

    document.add_heading("2. Dati OMI", level=2)
    p = document.add_paragraph()
    p.add_run("Comune (OMI): ").bold = True
    p.add_run(omi.comune_kml)
    p = document.add_paragraph()
    p.add_run("Provincia: ").bold = True
    p.add_run(omi.provincia_kml)
    p = document.add_paragraph()
    p.add_run("Zona OMI: ").bold = True
    p.add_run(omi.zona_codice)
    p = document.add_paragraph()
    p.add_run("Descrizione zona: ").bold = True
    p.add_run(omi.zona_descrizione)

    # Tabella valori €/m²
    document.add_paragraph("")
    document.add_paragraph("Valori di riferimento €/m² (compravendita):")
    table = document.add_table(rows=2, cols=4)
    hdr_cells = table.rows[0].cells
    hdr_cells[0].text = "Tipologia"
    hdr_cells[1].text = "Min"
    hdr_cells[2].text = "Med"
    hdr_cells[3].text = "Max"

    row = table.rows[1].cells
    row[0].text = "Residenziale"
    row[1].text = f"{val_min_mq:,.0f}".replace(",", ".")
    row[2].text = f"{val_med_mq:,.0f}".replace(",", ".")
    row[3].text = f"{val_max_mq:,.0f}".replace(",", ".")

    # Tabella stima totale
    val_tot_min = val_min_mq * superficie
    val_tot_med = val_med_mq * superficie
    val_tot_max = val_max_mq * superficie

    document.add_paragraph("")
    document.add_paragraph("Stima di massima del valore complessivo:")

    table2 = document.add_table(rows=2, cols=4)
    hdr2 = table2.rows[0].cells
    hdr2[0].text = "Tipologia"
    hdr2[1].text = "Min"
    hdr2[2].text = "Med"
    hdr2[3].text = "Max"

    r2 = table2.rows[1].cells
    r2[0].text = "Valore immobile"
    r2[1].text = f"{val_tot_min:,.0f} €".replace(",", ".")
    r2[2].text = f"{val_tot_med:,.0f} €".replace(",", ".")
    r2[3].text = f"{val_tot_max:,.0f} €".replace(",", ".")

    document.add_paragraph("")
    document.add_paragraph(
        "Nota: i valori riportati sono puramente indicativi e derivano unicamente "
        "dalle quotazioni OMI ufficiali. Non costituiscono una perizia."
    )

    buf = io.BytesIO()
    document.save(buf)
    buf.seek(0)
    return buf


# ---------------------------------------------------------
# UI principale
# ---------------------------------------------------------
def main():
    st.set_page_config(
        page_title="PlanetAI – Valutazione OMI Italia",
        page_icon="🏙️",
        layout="centered",
    )

    st.title("🏙️ PlanetAI – Valutazione OMI Italia")
    st.caption(f"📊 {OMI_DATA_INFO}")

    # Inizializzazione OMI (cache globale)
    init_omi_data()

    with st.sidebar:
        st.header("⚙️ Impostazioni")
        st.markdown(
            """
            Questa versione dell'app utilizza **solo i dati OMI ufficiali**:

            - Geocodifica l'indirizzo  
            - Trova la **zona OMI** corrispondente  
            - Mostra il range di valori **€/m²**  
            - Calcola una stima indicativa per la superficie inserita
            """
        )
        st.markdown("---")
        st.markdown(
            "💡 Suggerimento: inserisci anche il **numero civico** per una geocodifica più precisa."
        )

    # -----------------------------
    # Input utente
    # -----------------------------
    col1, col2 = st.columns(2)

    with col1:
        comune = st.text_input("🏘️ Comune", value="Como").strip()
        indirizzo = st.text_input("📍 Indirizzo (via e numero)", value="Via Mentana 1").strip()

    with col2:
        superficie = st.number_input(
            "📐 Superficie commerciale (m²)",
            min_value=10.0,
            max_value=1000.0,
            value=80.0,
            step=5.0,
        )
        st.markdown("")
        st.markdown("")
        calcola = st.button("🚀 Calcola con valori OMI")

    if not calcola:
        st.info("⬅️ Inserisci i dati e premi **“🚀 Calcola con valori OMI”** per iniziare.")
        return

    if not comune or not indirizzo:
        st.error("Per favore inserisci **Comune** e **Indirizzo**.")
        return

    # -----------------------------
    # Geocoding
    # -----------------------------
    with st.spinner("🛰️ Geocodifica dell'indirizzo..."):
        coord = geocode_indirizzo(comune, indirizzo)

    if coord is None:
        st.error("Impossibile geocodificare l'indirizzo. Prova a specificare meglio via e civico.")
        return

    lat, lon = coord
    st.success(f"📌 Coordinate trovate: **{lat:.6f}, {lon:.6f}**")

    # -----------------------------
    # Ricerca zona OMI
    # -----------------------------
    with st.spinner("🏙️ Ricerca della zona OMI..."):
        omi = get_quotazione_omi_da_coordinate(lat, lon)

    if omi is None:
        st.error(
            "Nessuna zona OMI trovata per queste coordinate. "
            "Verifica che il comune sia corretto."
        )
        return

    st.success("✅ Zona OMI trovata!")

    # -----------------------------
    # Dati zona OMI
    # -----------------------------
    st.subheader("🗺️ Zona OMI")

    col_z1, col_z2 = st.columns(2)
    with col_z1:
        st.markdown(f"**Comune (OMI):** {omi.comune_kml}")
        st.markdown(f"**Provincia:** {omi.provincia_kml}")
        st.markdown(f"**Zona OMI:** `{omi.zona_codice}`")
    with col_z2:
        st.markdown("**Descrizione zona:**")
        st.markdown(f"_{omi.zona_descrizione}_")

    # -----------------------------
    # Valori €/m²
    # -----------------------------
    val_min_mq = omi.val_min_mq or 0.0
    val_med_mq = omi.val_med_mq or val_min_mq
    val_max_mq = omi.val_max_mq or val_med_mq

    st.subheader("💶 Valori di riferimento €/m² (compravendita)")

    c1, c2, c3 = st.columns(3)
    c1.metric("Min", fmt_euro(val_min_mq))
    c2.metric("Medio", fmt_euro(val_med_mq))
    c3.metric("Max", fmt_euro(val_max_mq))

    # Grafico a barre
    df_vals = pd.DataFrame(
        {"€/m²": [val_min_mq, val_med_mq, val_max_mq]},
        index=["Min", "Medio", "Max"],
    )
    st.bar_chart(df_vals)

    # -----------------------------
    # Stima valore complessivo
    # -----------------------------
    st.subheader("🏡 Stima indicativa del valore dell'immobile")

    val_tot_min = val_min_mq * superficie
    val_tot_med = val_med_mq * superficie
    val_tot_max = val_max_mq * superficie

    c1, c2, c3 = st.columns(3)
    c1.metric("Valore minimo", fmt_euro(val_tot_min))
    c2.metric("Valore medio", fmt_euro(val_tot_med))
    c3.metric("Valore massimo", fmt_euro(val_tot_max))

    st.caption(
        "⚖️ La stima è puramente indicativa e si basa esclusivamente sui valori OMI "
        "della zona, senza considerare lo stato interno dell'immobile o altre "
        "caratteristiche specifiche."
    )

    # -----------------------------
    # Report Word
    # -----------------------------
    st.markdown("---")
    st.subheader("📄 Esporta report")

    if not HAVE_DOCX:
        st.warning(
            "Per generare il report Word è necessario installare la libreria "
            "`python-docx` nel file `requirements.txt` del progetto."
        )
        return

    buf = build_word_report(
        comune=comune,
        indirizzo=indirizzo,
        superficie=superficie,
        lat=lat,
        lon=lon,
        omi=omi,
        val_min_mq=val_min_mq,
        val_med_mq=val_med_mq,
        val_max_mq=val_max_mq,
    )

    default_filename = (
        f"Report_OMI_{comune.replace(' ', '_')}_"
        f"{indirizzo.replace(' ', '_')}_"
        f"{datetime.now().strftime('%Y%m%d_%H%M')}.docx"
    )

    st.download_button(
        "📥 Scarica report Word",
        data=buf,
        file_name=default_filename,
        mime=(
            "application/vnd.openxmlformats-officedocument."
            "wordprocessingml.document"
        ),
    )


if __name__ == "__main__":
    main()
