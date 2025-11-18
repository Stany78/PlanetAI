import io
from datetime import datetime

import pandas as pd
import streamlit as st

from agent_core import geocode_indirizzo
from omi_utils import get_quotazione_omi_da_coordinate, warmup_omi_cache

# ---------------------------------------------------------
# DOCX
# ---------------------------------------------------------
try:
    from docx import Document
    HAVE_DOCX = True
except:
    HAVE_DOCX = False


def build_word_report(comune_input, indirizzo_input, lat, lon, zona):
    """Crea il report Word premium basato SOLO su dati OMI."""
    doc = Document()

    doc.add_heading("Report di Valutazione OMI", level=1)
    doc.add_paragraph(f"Generato il: {datetime.now().strftime('%d/%m/%Y %H:%M')}")

    doc.add_heading("1. Dati di Input", level=2)
    doc.add_paragraph(f"Comune inserito: {comune_input}")
    doc.add_paragraph(f"Indirizzo inserito: {indirizzo_input}")
    doc.add_paragraph(f"Coordinate geografiche: {lat:.6f}, {lon:.6f}")

    doc.add_heading("2. Zona OMI trovata", level=2)
    doc.add_paragraph(f"Comune (OMI): {zona.comune}")
    doc.add_paragraph(f"Provincia: {zona.provincia}")
    doc.add_paragraph(f"Codice zona: {zona.zona_codice}")
    doc.add_paragraph(f"Descrizione zona: {zona.zona_descrizione}")

    doc.add_heading("3. Quotazioni OMI €/m²", level=2)
    table = doc.add_table(rows=2, cols=4)
    hdr = table.rows[0].cells
    hdr[0].text = "Parametro"
    hdr[1].text = "Minimo"
    hdr[2].text = "Mediano"
    hdr[3].text = "Massimo"

    row = table.rows[1].cells
    row[0].text = "Valori €/m²"
    row[1].text = f"{zona.val_min_mq:,.0f}".replace(",", ".")
    row[2].text = f"{zona.val_med_mq:,.0f}".replace(",", ".")
    row[3].text = f"{zona.val_max_mq:,.0f}".replace(",", ".")

    doc.add_heading("4. Interpretazione sintetica", level=2)
    doc.add_paragraph(
        f"Il valore mediano di {zona.val_med_mq:,.0f} €/m² indica che "
        f"la zona '{zona.zona_codice}' è una fascia di mercato "
        f"generalmente { 'alta' if zona.val_med_mq > zona.val_max_mq*0.7 else 'media' }."
    )

    doc.add_heading("5. Note metodologiche", level=2)
    doc.add_paragraph(
        "Le quotazioni OMI rappresentano valori statistici ufficiali dell’Agenzia delle Entrate "
        "espressi in €/m². Non considerano caratteristiche specifiche dell'immobile come "
        "piano, stato, esposizione, vista, anno di costruzione, qualità del condominio."
    )

    buffer = io.BytesIO()
    doc.save(buffer)
    buffer.seek(0)
    return buffer


# ---------------------------------------------------------
# INIT
# ---------------------------------------------------------
st.set_page_config(
    page_title="PlanetAI – Valutazione OMI",
    page_icon="🏙️",
    layout="centered",
)


@st.cache_resource(show_spinner="📦 Caricamento dati OMI...")
def init_omi():
    warmup_omi_cache()
    return True


init_omi()

# ---------------------------------------------------------
# UI
# ---------------------------------------------------------
st.title("🏙️ PlanetAI – Valutazione OMI")

st.markdown(
    """
Inserisci **Comune** e **Indirizzo** per ottenere:

- 📍 Zona OMI
- 💶 Quotazioni minimo / mediano / massimo in €/m²
- 📊 Grafico immediato
- 📄 Report Word scaricabile
"""
)

with st.form("form_omi"):
    colA, colB = st.columns(2)
    with colA:
        comune = st.text_input("Comune", value="Como")
    with colB:
        indirizzo = st.text_input("Indirizzo", value="Via Borgovico 150")

    submit = st.form_submit_button("Calcola valutazione OMI 🧮")

if not submit:
    st.stop()

if not comune.strip() or not indirizzo.strip():
    st.error("Inserisci sia il **Comune** che l'**Indirizzo**.")
    st.stop()

with st.spinner("📍 Geocoding e ricerca poligono OMI..."):
    lat, lon = geocode_indirizzo(comune, indirizzo)
    zona = get_quotazione_omi_da_coordinate(lat, lon)

if zona is None or zona.val_med_mq is None:
    st.error("⚠️ Nessuna zona OMI trovata per questo indirizzo.")
    st.stop()

st.success("✅ Zona OMI identificata correttamente!")

# ---------------------------------------------------------
# INFO ZONA
# ---------------------------------------------------------
st.subheader("📌 Zona OMI")

card1, card2 = st.columns(2)

with card1:
    st.markdown(
        f"""
        **Comune (OMI):** {zona.comune}  
        **Provincia:** {zona.provincia}  
        """
    )

with card2:
    st.markdown(
        f"""
        **Zona:** {zona.zona_codice}  
        **Descrizione:** {zona.zona_descrizione}  
        """
    )

st.markdown("---")

# ---------------------------------------------------------
# VALORI €/m²
# ---------------------------------------------------------
st.subheader("💶 Valori OMI €/m²")

c1, c2, c3 = st.columns(3)
c3.metric("Massimo", f"{zona.val_max_mq:,.0f} €/m²".replace(",", "."))
c2.metric("Mediano", f"{zona.val_med_mq:,.0f} €/m²".replace(",", "."))
c1.metric("Minimo", f"{zona.val_min_mq:,.0f} €/m²".replace(",", "."))

# ---------------------------------------------------------
# GRAFICO
# ---------------------------------------------------------
df = pd.DataFrame(
    {
        "Parametro": ["Minimo", "Mediano", "Massimo"],
        "Valore €/m²": [zona.val_min_mq, zona.val_med_mq, zona.val_max_mq],
    }
).set_index("Parametro")

st.bar_chart(df, height=260)

st.caption("Fonte: Agenzia delle Entrate – OMI")

st.markdown("---")

# ---------------------------------------------------------
# REPORT WORD
# ---------------------------------------------------------
if HAVE_DOCX:
    report = build_word_report(comune, indirizzo, lat, lon, zona)
    name = (
        f"Report_OMI_{comune.replace(' ', '_')}_"
        f"{indirizzo.replace(' ', '_')}_"
        f"{datetime.now().strftime('%Y%m%d_%H%M')}.docx"
    )
    st.download_button("📄 Scarica report Word", report, name)
else:
    st.info("Per il report Word installare: `python-docx` nel requirements.txt.")
