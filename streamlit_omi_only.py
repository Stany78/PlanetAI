import streamlit as st

from agent_core import geocode_indirizzo
from omi_utils import get_quotazione_omi_da_coordinate, warmup_omi_cache

st.set_page_config(
    page_title="Planet AI – Valutazione OMI per via",
    page_icon="🏙️",
    layout="centered",
)

# Inizializza cache OMI una sola volta
@st.cache_resource
def init_omi():
    warmup_omi_cache()
    return True

init_omi()

st.title("🏙️ Planet AI – Valutazione OMI per via")

st.write(
    """
Inserisci **Comune** e **Indirizzo** (via, civico): l'app:
1. geocoda l'indirizzo,
2. trova la **zona OMI** corrispondente,
3. mostra i valori OMI **min / med / max €/mq**.
"""
)

with st.form("omi_form"):
    comune = st.text_input("Comune", value="Como", placeholder="Es: Como")
    indirizzo = st.text_input("Indirizzo", value="Via Borgovico 150", placeholder="Es: Via Borgovico 150")

    submit = st.form_submit_button("Calcola valutazione OMI")

if not submit:
    st.stop()

if not comune.strip() or not indirizzo.strip():
    st.error("Inserisci sia il **Comune** che l'**Indirizzo**.")
    st.stop()

with st.spinner("Geocoding indirizzo e ricerca zona OMI..."):
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

st.success("Zona OMI trovata!")

# Dettagli zona
st.subheader("Zona OMI trovata")

col1, col2 = st.columns(2)
with col1:
    st.write(f"**Comune (OMI):** {zona_omi.comune}")
    st.write(f"**Provincia:** {zona_omi.provincia}")
with col2:
    st.write(f"**Zona OMI:** {zona_omi.zona_codice}")
    st.write(f"**Descrizione zona:** {zona_omi.zona_descrizione}")

st.markdown("---")
st.subheader("Valori OMI €/mq")

col_min, col_med, col_max = st.columns(3)
col_min.metric("Minimo", f"{zona_omi.val_min_mq:,.0f} €/mq".replace(",", "."))
col_med.metric("Mediano", f"{zona_omi.val_med_mq:,.0f} €/mq".replace(",", "."))
col_max.metric("Massimo", f"{zona_omi.val_max_mq:,.0f} €/mq".replace(",", "."))

st.caption("Fonte: dati OMI caricati dai file CSV e KML (Agenzia delle Entrate).")
