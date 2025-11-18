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
    con grafica migliorata e titolo Planet AI.
    """
    document = Document()

    # ------------------------------
    #  TITOLO PRINCIPALE
    # ------------------------------
    title = document.add_heading("", level=0)
    run_title = title.add_run("PLANET AI – Report OMI")
    run_title.bold = True
    run_title.font.size = Pt(22)

    document.add_paragraph(
        f"Data generazione report: {datetime.now().strftime('%d/%m/%Y %H:%M')}"
    )

    document.add_paragraph(
        "Il presente report riporta una stima basata esclusivamente sulle "
        "quotazioni OMI (€/m²) dell'Agenzia delle Entrate."
    )

    # ------------------------------
    # 1. DATI DI INPUT
    # ------------------------------
    document.add_heading("1. Dati di input", level=1)

    p = document.add_paragraph()
    p.add_run("Comune inserito: ").bold = True
    p.add_run(comune_input)

    p = document.add_paragraph()
    p.add_run("Indirizzo inserito: ").bold = True
    p.add_run(indirizzo_input)

    p = document.add_paragraph()
    p.add_run("Coordinate (lat, lon): ").bold = True
    p.add_run(f"{lat:.6f}, {lon:.6f}")

    # ------------------------------
    # 2. ZONA OMI
    # ------------------------------
    document.add_heading("2. Zona OMI di riferimento", level=1)

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

    # ------------------------------
    # 3. QUOTAZIONI €/m²
    # ------------------------------
    document.add_heading("3. Quotazioni OMI €/m² (compravendita)", level=1)

    table = document.add_table(rows=2, cols=4)
    hdr = table.rows[0].cells
    hdr[0].text = "Parametro"
    hdr[1].text = "Min"
    hdr[2].text = "Med"
    hdr[3].text = "Max"

    row = table.rows[1].cells
    row[0].text = "Valori €/m²"
    row[1].text = f"{zona_omi.val_min_mq:,.0f} €/m²".replace(",", ".")
    row[2].text = f"{zona_omi.val_med_mq:,.0f} €/m²".replace(",", ".")
    row[3].text = f"{zona_omi.val_max_mq:,.0f} €/m²".replace(",", ".")

    # ------------------------------
    # 4. INTERPRETAZIONE
    # ------------------------------
    document.add_heading("4. Interpretazione sintetica", level=1)

    med = zona_omi.val_med_mq
    if med < 2500:
        fascia = "Fascia bassa"
    elif med < 4500:
        fascia = "Fascia media"
    else:
        fascia = "Fascia alta"

    document.add_paragraph(
        f"La zona OMI '{zona_omi.zona_codice}' può essere classificata come **{fascia}** "
        f"in base al valore mediano di {med:,.0f} €/m²."
    )

    document.add_paragraph(
        "Il valore minimo rappresenta immobili da ristrutturare o in condizioni inferiori alla media, "
        "mentre il valore massimo riflette immobili di alta qualità, piani alti, o contesti premium."
    )

    # ------------------------------
    # 5. NOTE FINALI
    # ------------------------------
    document.add_heading("5. Limiti e note metodologiche", level=1)

    document.add_paragraph(
        "Le quotazioni OMI sono valori statistici della zona e non tengono conto delle caratteristiche "
        "specifiche dell'immobile (stato manutentivo, piano, vista, esposizione, anno di ristrutturazione, "
        "qualità del condominio, presenza di ascensore, ecc.)."
    )

    document.add_paragraph(
        "Il presente report non costituisce una perizia asseverata, ma una valutazione statistica "
        "basata su dati ufficiali dell'Agenzia delle Entrate."
    )

    # ------------------------------
    # ESPORTAZIONE
    # ------------------------------
    buffer = io.BytesIO()
    document.save(buffer)
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
st.title("🏙️ Valutazione OMI – PlanetAI")

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
