import io
from datetime import datetime

import pandas as pd
import streamlit as st

from agent_core import geocode_indirizzo
from omi_utils import get_quotazione_omi_da_coordinate, warmup_omi_cache

# ---------------------------------------------------------
# Word report (opzionale, se python-docx è installato)
# ---------------------------------------------------------
try:
    from docx import Document  # type: ignore

    HAVE_DOCX = True
except Exception:
    HAVE_DOCX = False


# ---------------------------------------------------------
# Utility
# ---------------------------------------------------------
def build_word_report(
    *,
    comune_input: str,
    indirizzo_input: str,
    lat: float,
    lon: float,
    zona_omi,
) -> io.BytesIO:
    """
    Crea un report Word basato solo sui dati OMI (€/m²),
    SENZA superficie e senza valore totale.
    """
    document = Document()

    document.add_heading("Report di valutazione OMI", level=1)
    document.add_paragraph(
        f"Data generazione report: {datetime.now().strftime('%d/%m/%Y %H:%M')}"
    )

    # 1. Dati immobile
    document.add_heading("1. Dati immobile", level=2)
    p = document.add_paragraph()
    p.add_run("Comune inserito: ").bold = True
    p.add_run(comune_input)

    p = document.add_paragraph()
    p.add_run("Indirizzo inserito: ").bold = True
    p.add_run(indirizzo_input)

    p = document.add_paragraph()
    p.add_run("Coordinate (lat, lon): ").bold = True
    p.add_run(f"{lat:.6f}, {lon:.6f}")

    # 2. Dati zona OMI
    document.add_heading("2. Zona OMI", level=2)

    p = document.add_paragraph()
    p.add_run("Comune (OMI): ").bold = True
    p.add_run(str(zona_omi.comune))

    p = document.add_paragraph()
    p.add_run("Provincia: ").bold = True
    p.add_run(str(zona_omi.provincia))

    p = document.add_paragraph()
    p.add_run("Zona OMI: ").bold = True
    p.add_run(str(zona_omi.zona_codice))

    p = document.add_paragraph()
    p.add_run("Descrizione zona: ").bold = True
    p.add_run(str(zona_omi.zona_descrizione))

    # 3. Valori €/m²
    document.add_heading("3. Valori OMI €/m² (compravendita)", level=2)
    table = document.add_table(rows=2, cols=4)
    hdr = table.rows[0].cells
    hdr[0].text = "Tipologia"
    hdr[1].text = "Min"
    hdr[2].text = "Med"
    hdr[3].text = "Max"

    row = table.rows[1].cells
    row[0].text = "Residenziale"
    row[1].text = f"{zona_omi.val_min_mq:,.0f}".replace(",", ".")
    row[2].text = f"{zona_omi.val_med_mq:,.0f}".replace(",", ".")
    row[3].text = f"{zona_omi.val_max_mq:,.0f}".replace(",", ".")

    # Nota finale
    document.add_paragraph(
        "Il presente report riporta esclusivamente le quotazioni OMI espresse in €/m², "
        "senza considerare caratteristiche specifiche dell'immobile (stato, piano, vista, ecc.)."
    )

    buffer = io.BytesIO()
    document.save(buffer)
    buffer.seek(0)
    return buffer


# ---------------------------------------------------------
# Cache OMI
# ---------------------------------------------------------
@st.cache_resource(show_spinner="📦 Caricamento dati OMI (solo al primo avvio)...")
def init_omi():
    warmup_omi_cache()
    return True


init_omi()

# ---------------------------------------------------------
# UI Streamlit
# ---------------------------------------------------------
st.set_page_config(
    page_title="PlanetAI – Valutazione OMI",
    page_icon="🏙️",
    layout="centered",
)

st.title("🏙️ PlanetAI – Valutazione OMI")

st.write(
    """
Inserisci **Comune** e **Indirizzo** (via, civico).

L'app:
1. geocoda l'indirizzo,
2. trova la **zona OMI**,
3. mostra i valori OMI **min / med / max €/m²**,
4. permette di scaricare un **report Word** solo con quotazioni €/m².
"""
)

with st.form("omi_form"):
    col_a, col_b = st.columns(2)
    with col_a:
        comune = st.text_input("Comune", value="Como", placeholder="Es: Como")
    with col_b:
        indirizzo = st.text_input(
            "Indirizzo",
            value="Via Borgovico 150",
            placeholder="Es: Via Borgovico 150",
        )

    submit = st.form_submit_button("Calcola valutazione OMI 🧮")

if not submit:
    st.stop()

if not comune.strip() or not indirizzo.strip():
    st.error("Inserisci sia il **Comune** che l'**Indirizzo**.")
    st.stop()

with st.spinner("📍 Geocoding indirizzo e ricerca zona OMI..."):
    # 1) Geocoding
    lat, lon = geocode_indirizzo(comune, indirizzo)

    # 2) Quotazione OMI da coordinate
    zona_omi = get_quotazione_omi_da_coordinate(lat, lon)

if zona_omi is None or zona_omi.val_med_mq is None:
    st.error(
        "Non è stato possibile trovare una zona OMI per queste coordinate. "
        "Verifica che l'indirizzo sia corretto oppure che i dati OMI/KML coprano questa zona."
    )
    st.stop()

st.success("✅ Zona OMI trovata!")

# ---------------------------------------------------------
# Dettagli zona
# ---------------------------------------------------------
st.subheader("📌 Zona OMI")

zona_col1, zona_col2 = st.columns(2)
with zona_col1:
    st.write(f"**Comune (OMI):** {zona_omi.comune}")
    st.write(f"**Provincia:** {zona_omi.provincia}")
with zona_col2:
    st.write(f"**Zona OMI:** {zona_omi.zona_codice}")
    st.write(f"**Descrizione zona:** {zona_omi.zona_descrizione}")

st.markdown("---")

# ---------------------------------------------------------
# Valori €/m²
# ---------------------------------------------------------
st.subheader("💶 Valori OMI €/m²")

col_min, col_med, col_max = st.columns(3)
col_min.metric("Minimo", f"{zona_omi.val_min_mq:,.0f} €/m²".replace(",", "."))
col_med.metric("Mediano", f"{zona_omi.val_med_mq:,.0f} €/m²".replace(",", "."))
col_max.metric("Massimo", f"{zona_omi.val_max_mq:,.0f} €/m²".replace(",", "."))

# 🔁 Istogramma invertito: Massimo → Mediano → Minimo
df_valori = pd.DataFrame(
    {
        "Tipologia": ["Massimo", "Mediano", "Minimo"],
        "Valore €/m²": [
            zona_omi.val_max_mq,
            zona_omi.val_med_mq,
            zona_omi.val_min_mq,
        ],
    }
)
st.bar_chart(
    data=df_valori.set_index("Tipologia"),
    height=260,
)

st.caption("Fonte: dati OMI caricati dai file CSV e KML (Agenzia delle Entrate).")

st.markdown("---")

# ---------------------------------------------------------
# Download report Word
# ---------------------------------------------------------
if HAVE_DOCX:
    report_buffer = build_word_report(
        comune_input=comune,
        indirizzo_input=indirizzo,
        lat=lat,
        lon=lon,
        zona_omi=zona_omi,
    )

    file_name = (
        f"Report_OMI_{comune.replace(' ', '_')}_"
        f"{indirizzo.replace(' ', '_')}_"
        f"{datetime.now().strftime('%Y%m%d_%H%M')}.docx"
    )

    st.download_button(
        label="📄 Scarica report Word (quotazioni €/m²)",
        data=report_buffer,
        file_name=file_name,
        mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document",
    )
else:
    st.info(
        "Per abilitare il download del report Word, aggiungi `python-docx` al file `requirements.txt`."
    )
