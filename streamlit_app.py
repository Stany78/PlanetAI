import os
from datetime import datetime

import streamlit as st

from agent_core import (
    stima_prezzo_zona_nuova_costruzione,
    crea_grafico_omi_vs_stima,
    genera_report_word,
)
from omi_utils import warmup_omi_cache
@st.cache_resource
def init_omi_cache():
    warmup_omi_cache()
    return True
from config import BASE_DIR


# ===================================================
# CONFIGURAZIONE BASE APP
# ===================================================

st.set_page_config(
    page_title="Planet AI – Stima nuova costruzione",
    page_icon="🏗️",
    layout="centered",
)

@st.cache_resource
def init_omi_cache():
    """Warmup OMI all'avvio (una sola volta)."""
    try:
        warmup_omi_cache()
    except Exception as e:
        st.warning(f"Impossibile pre-caricare i dati OMI: {e}")

init_omi_cache()

st.title("🏙️ Planet AI – Stima nuova costruzione")
st.write(
    """
Questa app stima il **valore €/mq di una nuova costruzione** a partire da:
- Geocoding dell'indirizzo
- Valori **OMI** (Agenzia delle Entrate)
- Annunci reali da **Immobiliare.it** sul comune

La stima finale combina OMI e portali (quando disponibili).
"""
)

# ===================================================
# INPUT UTENTE
# ===================================================

with st.form("input_stima"):
    st.subheader("Dati immobile")
    comune = st.text_input("Comune", value="Como", placeholder="Es: Como")
    indirizzo = st.text_input("Indirizzo", value="Via Borgovico 150", placeholder="Es: Via Borgovico 150")
    genera_report = st.checkbox("Genera anche report Word (.docx)", value=True)

    submitted = st.form_submit_button("Calcola stima")

if not submitted:
    st.stop()

if not comune.strip() or not indirizzo.strip():
    st.error("Inserisci **comune** e **indirizzo** per procedere.")
    st.stop()

# ===================================================
# ESECUZIONE STIMA
# ===================================================

with st.spinner("Calcolo stima, recupero dati OMI e annunci di mercato..."):
    try:
        risultato = stima_prezzo_zona_nuova_costruzione(comune, indirizzo)
    except Exception as e:
        st.error(f"Errore nel calcolo della stima: {e}")
        st.stop()

st.success("Stima completata!")

# ===================================================
# MOSTRA RISULTATO PRINCIPALE
# ===================================================

st.subheader("Risultato – Stima nuova costruzione (€ / mq)")

stima = risultato["stima_nuova_costruzione"]

col_p, col_c, col_a = st.columns(3)
col_p.metric("Prudente", f"{stima['prudente']:,.0f} €/mq".replace(",", "."))
col_c.metric("Centrale", f"{stima['centrale']:,.0f} €/mq".replace(",", "."))
col_a.metric("Aggressivo", f"{stima['aggressivo']:,.0f} €/mq".replace(",", "."))

# ===================================================
# DETTAGLIO OMI
# ===================================================

st.markdown("---")
st.subheader("Valori OMI")

zona = risultato.get("zona_omi")
if zona and zona.get("val_med") is not None:
    st.write(f"**Comune:** {zona['comune']} – **Zona OMI:** {zona['codice']}")

    descr = zona.get("descrizione")

    # Se la descrizione è una tabella HTML, la rendiamo come HTML
    if descr and "<table" in descr.lower():
        st.markdown("**Descrizione zona:**", unsafe_allow_html=True)
        st.markdown(descr, unsafe_allow_html=True)
    else:
        st.write(f"**Descrizione zona:** {descr or 'N/D'}")

    st.write(
        f"**Valori OMI €/mq** – "
        f"min: **{zona['val_min']:.0f}**, "
        f"med: **{zona['val_med']:.0f}**, "
        f"max: **{zona['val_max']:.0f}**"
    )
else:
    st.info("Valori OMI non disponibili per questa zona.")


# ===================================================
# STATISTICHE ANNUNCI
# ===================================================

st.markdown("---")
st.subheader("Statistiche annunci (Immobiliare.it)")

usato = risultato.get("usato") or {}
nuovo = risultato.get("nuovo") or {}

def render_stats(label, stats, col):
    if not stats or "mediana" not in stats:
        col.write(f"**{label}:** nessun dato utilizzabile")
        return
    col.markdown(f"### {label}")
    col.write(f"Mediana: **{stats['mediana']:.0f} €/mq**")
    col.write(f"P25: **{stats['p25']:.0f} €/mq**")
    col.write(f"P75: **{stats['p75']:.0f} €/mq**")
    col.write(f"N annunci: **{stats['n']}**")

col_u, col_n = st.columns(2)
render_stats("Usato", usato, col_u)
render_stats("Nuovo", nuovo, col_n)

# ===================================================
# NOTE
# ===================================================

st.markdown("---")
st.subheader("Note sulla stima")

for n in risultato.get("note", []):
    st.write(f"- {n}")

# ===================================================
# GRAFICO OMI vs STIMA
# ===================================================

st.markdown("---")
st.subheader("Confronto grafico OMI vs stima")

try:
    reports_dir = os.path.join(BASE_DIR, "reports_streamlit")
    os.makedirs(reports_dir, exist_ok=True)

    grafico_path = crea_grafico_omi_vs_stima(
        zona_omi=risultato.get("zona_omi"),
        stima=stima,
        output_dir=reports_dir,
    )

    if grafico_path and os.path.exists(grafico_path):
        st.image(grafico_path, caption="Valori OMI vs stima nuova costruzione", use_column_width=True)
    else:
        st.info("Grafico non disponibile (errore in generazione).")
except Exception as e:
    st.warning(f"Errore nella generazione/visualizzazione del grafico: {e}")

# ===================================================
# REPORT WORD – DOWNLOAD
# ===================================================

if genera_report:
    st.markdown("---")
    st.subheader("Report Word")

    try:
        report_path = genera_report_word(risultato)
        if report_path and os.path.exists(report_path):
            with open(report_path, "rb") as f:
                data = f.read()
            filename = os.path.basename(report_path)
            st.download_button(
                label="📄 Scarica report Word",
                data=data,
                file_name=filename,
                mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document",
            )
        else:
            st.info("Impossibile generare il report Word.")
    except Exception as e:
        st.warning(f"Errore nella generazione del report Word: {e}")

st.caption(f"Ultimo aggiornamento: {datetime.now().strftime('%d/%m/%Y %H:%M')}")
