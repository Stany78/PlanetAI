"""
PLANET AI - Analisi Immobiliare Completa
==========================================
Integrazione:
- Dati OMI (Agenzia delle Entrate)
- Mercato Immobiliare.it
- Report combinato
"""

import streamlit as st
import pandas as pd
import os
from datetime import datetime

# Import moduli locali
from agent_core import geocode_indirizzo
from omi_utils import get_quotazione_omi_da_coordinate, warmup_omi_cache
from immobiliare_scraper import cerca_appartamenti, calcola_statistiche
from report_generator import genera_report_combinato
# Import condizionale per evitare crash se claude_analyzer non esiste
try:
    from claude_analyzer import analizza_con_ai, get_api_key
    CLAUDE_AVAILABLE = True
except ImportError:
    CLAUDE_AVAILABLE = False
    print("⚠️ Modulo claude_analyzer non disponibile - analisi AI disabilitata")
from config import REPORTS_DIR
# Import opzionali per mappe e geocoding
try:
    from map_generator import crea_mappa_interattiva, get_mappa_statistiche
    MAP_AVAILABLE = True
except ImportError as e:
    MAP_AVAILABLE = False
    print(f"⚠️ Moduli mappa non disponibili: {e}")

# Configurazione pagina
st.set_page_config(
    page_title="Planet AI - Analisi Immobiliare",
    page_icon="🏢",
    layout="wide",
)

# Inizializza cache OMI
@st.cache_resource
def init_omi():
    warmup_omi_cache()
    return True

init_omi()

# ========================================
# HEADER
# ========================================
st.title("🏢 Planet AI - Analisi Immobiliare Completa")
st.markdown("""
Analizza una zona immobiliare combinando:
- **📊 Dati OMI** (valori ufficiali rogiti - Agenzia delle Entrate)
- **🏠 Mercato Immobiliare.it** (offerte attuali nuove costruzioni)
""")

st.markdown("---")

# ========================================
# FORM INPUT
# ========================================
with st.form("analisi_form"):
    st.subheader("📍 Parametri Ricerca")
    
    col1, col2, col3 = st.columns(3)
    
    with col1:
        comune = st.text_input("Comune", value="Como", placeholder="Es: Como")
    
    with col2:
        via = st.text_input("Via/Indirizzo", value="Via Anzani", placeholder="Es: Via Anzani")
    
    with col3:
        raggio_km = st.number_input("Raggio (km)", min_value=0.5, max_value=5.0, value=1.0, step=0.5)
    
    st.markdown("---")
    
    # Opzione Analisi AI
    col_ai, col_space = st.columns([1, 3])
    with col_ai:
        abilita_ai = st.checkbox("🤖 Abilita Analisi AI", value=True, help="Genera un'analisi professionale con Claude AI")
    
    st.markdown("---")
    
    col_btn1, col_btn2 = st.columns([1, 4])
    with col_btn1:
        submit = st.form_submit_button("🔍 Avvia Analisi", use_container_width=True)

# Stop se non submit
if not submit:
    st.info("👆 Compila i campi e clicca 'Avvia Analisi' per iniziare")
    st.stop()

# Validazione input
if not comune.strip() or not via.strip():
    st.error("⚠️ Inserisci sia il **Comune** che l'**Indirizzo**.")
    st.stop()

# ========================================
# ELABORAZIONE
# ========================================

# Progress bar
progress_bar = st.progress(0)
status_text = st.empty()

# 1. GEOCODING
status_text.text("🗺️ Geocoding indirizzo...")
progress_bar.progress(20)

lat, lon, geo_info = geocode_indirizzo(comune, via)

# Mostra risultato geocoding
if geo_info['success']:
    st.success(geo_info['message'])
    st.info(f"📍 Coordinate: {lat:.6f}, {lon:.6f}")
else:
    # Geocoding fallito - mostra errore e blocca
    progress_bar.progress(0)
    status_text.text("")
    
    st.error(geo_info['message'])
    
    st.warning("""
    **Come risolvere:**
    - Verifica l'ortografia della via
    - Aggiungi il numero civico (es: "Via Anzani 10")
    - Prova con una via principale vicina
    """)
    
    st.info("💡 **Suggerimento**: Usa Google Maps per verificare il nome esatto della via")
    
    st.stop()  # Blocca qui - non prosegue

# 2. DATI OMI
status_text.text("📊 Ricerca zona OMI...")
progress_bar.progress(40)

zona_omi_obj = get_quotazione_omi_da_coordinate(lat, lon)

# Converti oggetto OMI in dict per facilità
zona_omi = None
if zona_omi_obj:
    zona_omi = {
        'comune': zona_omi_obj.comune,
        'provincia': zona_omi_obj.provincia,
        'zona_codice': zona_omi_obj.zona_codice,
        'zona_descrizione': zona_omi_obj.zona_descrizione,
        'val_min_mq': zona_omi_obj.val_min_mq,
        'val_med_mq': zona_omi_obj.val_med_mq,
        'val_max_mq': zona_omi_obj.val_max_mq,
    }

# 3. SCRAPING IMMOBILIARE.IT
status_text.text("🏠 Scraping Immobiliare.it...")
progress_bar.progress(60)

appartamenti = cerca_appartamenti(lat, lon, raggio_km, max_pagine=5)

# 4. CALCOLO STATISTICHE
status_text.text("📈 Calcolo statistiche...")
progress_bar.progress(80)

stats_immobiliare = calcola_statistiche(appartamenti) if appartamenti else None

progress_bar.progress(100)
status_text.text("✅ Analisi completata!")

# 5. ANALISI AI AUTOMATICA (solo se abilitata dall'utente)
if abilita_ai and CLAUDE_AVAILABLE:
    status_text.text("🤖 Analisi AI in corso...")
    
    try:
        risultato_ai = analizza_con_ai(
            comune=comune,
            via=via,
            zona_omi=zona_omi,
            stats_immobiliare=stats_immobiliare
        )
    except Exception as e:
        print(f"Errore analisi AI: {e}")
        risultato_ai = {'success': False, 'error': str(e)}
else:
    if not abilita_ai:
        print("ℹ️ Analisi AI disabilitata dall'utente")
        risultato_ai = {'success': False, 'error': 'Analisi AI disabilitata dall\'utente'}
    else:
        risultato_ai = {'success': False, 'error': 'Modulo AI non disponibile'}

status_text.text("✅ Completato!")

# 6. GENERA REPORT AUTOMATICAMENTE (con tutti i dati per analisi developer)
status_text.text("📝 Generazione report Word...")

try:
    os.makedirs(REPORTS_DIR, exist_ok=True)
    
    report_filepath = genera_report_combinato(
        comune=comune,
        via=via,
        lat=lat,
        lon=lon,
        raggio_km=raggio_km,
        zona_omi=zona_omi,
        stats_immobiliare=stats_immobiliare,
        appartamenti=appartamenti,
        analisi_ai=risultato_ai,
        output_dir=REPORTS_DIR
    )
    
    # Leggi il file in memoria
    with open(report_filepath, 'rb') as f:
        report_data = f.read()
    
    report_filename = os.path.basename(report_filepath)
    
    status_text.text("✅ Report generato!")
    
except Exception as e:
    print(f"Errore generazione report: {e}")
    report_data = None
    report_filename = None

# SALVA IN SESSION STATE per mantenere i dati
st.session_state.analisi_data = {
    'comune': comune,
    'via': via,
    'lat': lat,
    'lon': lon,
    'raggio_km': raggio_km,
    'zona_omi': zona_omi,
    'appartamenti': appartamenti,
    'stats_immobiliare': stats_immobiliare,
    'report_data': report_data,
    'report_filename': report_filename,
    'analisi_ai': risultato_ai,
}

st.markdown("---")

# ========================================
# VISUALIZZAZIONE RISULTATI
# ========================================

# TAB per organizzare output
tab1, tab2, tab3, tab3b, tab4, tab5, tab6 = st.tabs([
    "📊 Dati OMI", 
    "🏠 Immobiliare.it", 
    "📈 Confronto",
    "🗺️ Mappa",
    "💼 Analisi Developer",
    "🤖 Analisi AI", 
    "📄 Report"
])

# ----------------------------------------
# TAB 1: DATI OMI
# ----------------------------------------
with tab1:
    st.header("📊 Dati OMI (Agenzia delle Entrate)")
    
    if zona_omi and zona_omi['val_med_mq'] is not None:
        
        col1, col2 = st.columns(2)
        
        with col1:
            st.subheader("Zona OMI")
            st.write(f"**Codice:** {zona_omi['zona_codice']}")
            st.write(f"**Descrizione:** {zona_omi['zona_descrizione']}")
            st.write(f"**Comune:** {zona_omi['comune']} ({zona_omi['provincia']})")
        
        with col2:
            st.subheader("Valori €/mq")
            col_min, col_med, col_max = st.columns(3)
            col_min.metric("Minimo", f"€{zona_omi['val_min_mq']:,.0f}".replace(',', '.'))
            col_med.metric("Mediano", f"€{zona_omi['val_med_mq']:,.0f}".replace(',', '.'))
            col_max.metric("Massimo", f"€{zona_omi['val_max_mq']:,.0f}".replace(',', '.'))
        
        st.caption("Fonte: dati ufficiali rogiti - Agenzia delle Entrate (QI 2025/1)")
    else:
        st.warning("⚠️ Dati OMI non disponibili per questa zona.")

# ----------------------------------------
# TAB 2: IMMOBILIARE.IT
# ----------------------------------------
with tab2:
    st.header("🏠 Analisi Mercato Immobiliare.it")
    
    if stats_immobiliare and stats_immobiliare['n_appartamenti'] > 0:
        
        st.subheader(f"📋 Appartamenti analizzati: {stats_immobiliare['n_appartamenti']}")
        
        st.markdown("---")
        
        # Statistiche principali
        col1, col2, col3 = st.columns(3)
        
        with col1:
            st.metric("Prezzo Medio", f"€{stats_immobiliare['prezzo']['medio']:,.0f}".replace(',', '.'))
            st.metric("Prezzo Mediano", f"€{stats_immobiliare['prezzo']['mediano']:,.0f}".replace(',', '.'))
        
        with col2:
            st.metric("Superficie Media", f"{stats_immobiliare['mq']['medio']:.0f} m²")
            st.metric("Superficie Mediana", f"{stats_immobiliare['mq']['mediano']:.0f} m²")
        
        with col3:
            st.metric("Prezzo/mq Medio", f"€{stats_immobiliare['prezzo_mq']['medio']:,.0f}".replace(',', '.'))
            st.metric("Prezzo/mq Mediano", f"€{stats_immobiliare['prezzo_mq']['mediano']:,.0f}".replace(',', '.'))
        
        st.markdown("---")
        
        # Tabella agenzie
        st.subheader("🏢 Analisi per Agenzia")
        
        df = stats_immobiliare['dataframe']
        agenzie = df.groupby('agenzia').agg({
            'prezzo': ['count', 'mean'],
            'mq': 'mean',
            'prezzo_mq': 'mean'
        }).reset_index()
        
        agenzie.columns = ['Agenzia', 'N° Appartamenti', 'Prezzo Medio', 'MQ Medio', 'Prezzo/mq Medio']
        agenzie = agenzie.sort_values('N° Appartamenti', ascending=False)
        
        # Formatta per display
        agenzie_display = agenzie.copy()
        agenzie_display['Prezzo Medio'] = agenzie_display['Prezzo Medio'].apply(lambda x: f"€{x:,.0f}".replace(',', '.'))
        agenzie_display['MQ Medio'] = agenzie_display['MQ Medio'].apply(lambda x: f"{x:.0f} m²")
        agenzie_display['Prezzo/mq Medio'] = agenzie_display['Prezzo/mq Medio'].apply(lambda x: f"€{x:,.0f}/m²".replace(',', '.'))
        
        # Riordina colonne
        agenzie_display = agenzie_display[['Agenzia', 'N° Appartamenti', 'Prezzo Medio', 'MQ Medio', 'Prezzo/mq Medio']]
        
        st.dataframe(agenzie_display, width="stretch", hide_index=True)
        
        st.markdown("---")
        
        # Distribuzione prezzi
        col1, col2 = st.columns(2)
        
        with col1:
            st.subheader("💰 Distribuzione Prezzi")
            fasce_prezzo = pd.DataFrame([
                {'Fascia': 'Fino a €200k', 'N': len(df[df['prezzo'] <= 200000])},
                {'Fascia': '€200k - €350k', 'N': len(df[(df['prezzo'] > 200000) & (df['prezzo'] <= 350000)])},
                {'Fascia': '€350k - €500k', 'N': len(df[(df['prezzo'] > 350000) & (df['prezzo'] <= 500000)])},
                {'Fascia': 'Oltre €500k', 'N': len(df[df['prezzo'] > 500000])},
            ])
            st.bar_chart(fasce_prezzo.set_index('Fascia'))
        
        with col2:
            st.subheader("📏 Distribuzione Superfici")
            fasce_mq = pd.DataFrame([
                {'Fascia': 'Fino a 60 m²', 'N': len(df[df['mq'] <= 60])},
                {'Fascia': '60 - 100 m²', 'N': len(df[(df['mq'] > 60) & (df['mq'] <= 100)])},
                {'Fascia': '100 - 150 m²', 'N': len(df[(df['mq'] > 100) & (df['mq'] <= 150)])},
                {'Fascia': 'Oltre 150 m²', 'N': len(df[df['mq'] > 150])},
            ])
            st.bar_chart(fasce_mq.set_index('Fascia'))
        
    else:
        st.warning("⚠️ Nessun appartamento trovato su Immobiliare.it per questa zona.")

# ----------------------------------------
# TAB 3: CONFRONTO
# ----------------------------------------
with tab3:
    st.header("📈 Confronto OMI vs Mercato")
    
    if zona_omi and zona_omi['val_med_mq'] and stats_immobiliare:
        
        omi_med = zona_omi['val_med_mq']
        mercato_med = stats_immobiliare['prezzo_mq']['mediano']
        gap = ((mercato_med - omi_med) / omi_med) * 100
        
        col1, col2, col3 = st.columns(3)
        
        col1.metric("OMI Mediano", f"€{omi_med:,.0f}/m²".replace(',', '.'))
        col2.metric("Mercato Mediano", f"€{mercato_med:,.0f}/m²".replace(',', '.'))
        col3.metric("Gap", f"{gap:+.1f}%".replace('.', ','), 
                   delta=f"€{mercato_med - omi_med:,.0f}/m²".replace(',', '.'))
        
        st.markdown("---")
        
        # Interpretazione
        st.subheader("💡 Interpretazione")
        
        if gap > 15:
            st.info("""
            **Mercato sopra OMI (+15%+)**
            
            Il mercato quota significativamente sopra i valori OMI. Possibili scenari:
            - Zona ad alta domanda
            - Prezzi di offerta ottimistici
            - Qualità superiore delle nuove costruzioni
            """)
        elif gap > 5:
            st.success("""
            **Mercato moderatamente sopra OMI (+5% - +15%)**
            
            Situazione normale. Le offerte iniziali sono tipicamente più alte dei valori di transazione effettivi.
            """)
        elif gap > -5:
            st.success("""
            **Mercato allineato a OMI (-5% - +5%)**
            
            Prezzi coerenti con le transazioni effettive. Mercato equilibrato.
            """)
        else:
            st.warning("""
            **Mercato sotto OMI (-5%-)**
            
            Il mercato quota sotto i valori OMI. Possibili scenari:
            - Opportunità di acquisto
            - Necessità di rilancio del settore
            - Verifica qualità offerte
            """)
        
        # Grafico comparativo
        st.markdown("---")
        st.subheader("📊 Confronto Visivo")
        
        import plotly.graph_objects as go
        
        fig = go.Figure()
        
        # Barre OMI
        fig.add_trace(go.Bar(
            name='OMI',
            x=['Minimo', 'Mediano', 'Massimo'],
            y=[zona_omi['val_min_mq'], zona_omi['val_med_mq'], zona_omi['val_max_mq']],
            marker_color='lightblue'
        ))
        
        # Barre Mercato
        fig.add_trace(go.Bar(
            name='Mercato',
            x=['Minimo', 'Mediano', 'Massimo'],
            y=[stats_immobiliare['prezzo_mq']['min'], 
               stats_immobiliare['prezzo_mq']['mediano'], 
               stats_immobiliare['prezzo_mq']['max']],
            marker_color='lightcoral'
        ))
        
        fig.update_layout(
            title='Confronto OMI vs Mercato (€/m²)',
            yaxis_title='€/m²',
            barmode='group',
            height=400
        )
        
        st.plotly_chart(fig, use_container_width=True)
        
    else:
        st.warning("⚠️ Dati insufficienti per il confronto.")


# ----------------------------------------
# TAB 3B: MAPPA
# ----------------------------------------
with tab3b:
    st.header("🗺️ Mappa Interattiva Appartamenti")
    
    print(f"[TAB_MAPPA] Entrato nel TAB Mappa")
    print(f"[TAB_MAPPA] appartamenti type: {type(appartamenti)}")
    print(f"[TAB_MAPPA] appartamenti len: {len(appartamenti) if appartamenti else 0}")
    
    if not appartamenti or len(appartamenti) == 0:
        print(f"[TAB_MAPPA] NESSUN appartamento - mostro warning")
        st.warning("⚠️ Nessun appartamento da visualizzare sulla mappa.")
    else:
        print(f"[TAB_MAPPA] Ho {len(appartamenti)} appartamenti - creo mappa")
        try:
            # Crea mappa
            # Usa il dataframe pulito (senza duplicati) invece della lista originale
            appartamenti_per_mappa = stats_immobiliare['dataframe'].to_dict('records') if stats_immobiliare else appartamenti
            
            mappa = crea_mappa_interattiva(
                lat_centro=lat,
                lon_centro=lon,
                via=via,
                comune=comune,
                raggio_km=raggio_km,
                appartamenti=appartamenti_per_mappa,
                stats_immobiliare=stats_immobiliare
            )
            
            # Mostra mappa con streamlit-folium
            try:
                from streamlit_folium import st_folium
                st_folium(mappa, width=1200, height=600, returned_objects=[])
            except ImportError:
                # Fallback se streamlit-folium non installato
                st.warning("⚠️ Modulo streamlit-folium non disponibile")
                st.info("La mappa è stata generata ma non può essere visualizzata. Sarà inclusa nel report Word.")
            
            # Statistiche mappa
            st.markdown("---")
            st.subheader("📊 Statistiche Geografiche")
            
            map_stats = get_mappa_statistiche(appartamenti)
            
            col1, col2, col3 = st.columns(3)
            col1.metric("Totale Appartamenti", map_stats['totale'])
            col2.metric("Con Coordinate", map_stats['con_coordinate'])
            col3.metric("Senza Coordinate", map_stats['senza_coordinate'])
            
            if map_stats['con_coordinate'] > 0:
                st.caption(f"Distribuzione geografica: lat {map_stats['lat_min']:.4f} - {map_stats['lat_max']:.4f}, lon {map_stats['lon_min']:.4f} - {map_stats['lon_max']:.4f}")
            
        except Exception as e:
            st.error(f"❌ Errore nella generazione della mappa: {str(e)}")
            st.info("La mappa potrebbe essere inclusa nel report Word se la generazione riesce.")

# ----------------------------------------
# TAB 4: ANALISI DEVELOPER
# ----------------------------------------
with tab4:
    st.header("💼 Analisi per Developer/Investitori")
    
    if not zona_omi or not stats_immobiliare or stats_immobiliare.get('n_appartamenti', 0) == 0:
        st.warning("⚠️ Esegui prima un'analisi completa per vedere questa sezione.")
    else:
        # ========================================
        # 1. SATURAZIONE MERCATO
        # ========================================
        st.subheader("🎯 1. Saturazione Mercato")
        
        n_app = stats_immobiliare['n_appartamenti']
        n_progetti = stats_immobiliare['n_progetti']
        
        col1, col2 = st.columns(2)
        
        with col1:
            st.metric("Appartamenti in Vendita", n_app)
        
        with col2:
            # Media appartamenti per agenzia top 5
            if stats_immobiliare.get('dataframe') is not None:
                df = stats_immobiliare['dataframe']
                agenzie_stats = df.groupby('agenzia').size()
                top5_media = agenzie_stats.nlargest(5).mean()
                st.metric("App/Agenzia Top 5 (media)", f"{top5_media:.1f}")
            else:
                st.metric("Agenzie Attive", "N/D")
        
        # Valutazione saturazione
        st.markdown("---")
        st.markdown("**Valutazione Saturazione:**")
        
        if n_app < 10:
            st.success("""Mercato LIBERO - Poca concorrenza
            
- Pochi appartamenti in vendita
- Buona opportunità di ingresso
- Minor rischio di invenduto
            """)
        elif n_app < 30:
            st.info("""Mercato MEDIO - Concorrenza normale
            
- Livello di offerta standard
- Necessaria differenziazione
- Attenzione al pricing
            """)
        else:
            st.warning("""Mercato SATURO - Alta concorrenza
            
- Molti appartamenti invenduti
- Rischio absorption rate basso
- Necessaria forte differenziazione o prezzi competitivi
            """)
        
        st.markdown("---")
        
        # ========================================
        # 2. PRICING BENCHMARK
        # ========================================
        st.subheader("💰 2. Pricing Benchmark")
        
        omi_med = zona_omi['val_med_mq']
        mercato_min = stats_immobiliare['prezzo_mq']['min']
        mercato_med = stats_immobiliare['prezzo_mq']['mediano']
        mercato_max = stats_immobiliare['prezzo_mq']['max']
        
        # Tabella comparativa
        col1, col2 = st.columns(2)
        
        with col1:
            st.markdown("**📊 Valori OMI (Rogiti Reali)**")
            st.metric("Valore Mediano OMI", f"€{omi_med:,.0f}/m²".replace(',', '.'))
            st.caption("Baseline ufficiale Agenzia Entrate")
        
        with col2:
            st.markdown("**🏠 Mercato Nuove Costruzioni**")
            st.metric("Range Prezzi", 
                     f"€{mercato_min:,.0f} - €{mercato_max:,.0f}/m²".replace(',', '.'))
            st.metric("Prezzo Mediano", f"€{mercato_med:,.0f}/m²".replace(',', '.'))
        
        st.markdown("---")
        
        # Target pricing consigliato
        st.markdown("**🎯 Target Pricing Consigliato**")
        
        gap_percentuale = ((mercato_med - omi_med) / omi_med) * 100
        
        # Calcola sweet spot
        target_min = mercato_med * 0.95  # -5% dal mediano
        target_max = mercato_med * 1.05  # +5% dal mediano
        
        col1, col2, col3 = st.columns(3)
        
        with col1:
            st.metric("Entry Level", f"€{target_min:,.0f}/m²".replace(',', '.'))
            st.caption("Per vendita veloce")
        
        with col2:
            st.metric("Sweet Spot", f"€{mercato_med:,.0f}/m²".replace(',', '.'))
            st.caption("Consigliato")
        
        with col3:
            st.metric("Premium", f"€{target_max:,.0f}/m²".replace(',', '.'))
            st.caption("Se alta qualità")
        
        st.markdown("---")
        
        # ========================================
        # 4. GAP ANALYSIS STRATEGICO
        # ========================================
        st.subheader("📊 4. Gap Analysis Strategico")
        
        col1, col2, col3 = st.columns(3)
        
        with col1:
            st.metric("OMI Baseline", f"€{omi_med:,.0f}/m²".replace(',', '.'))
        
        with col2:
            st.metric("Mercato Mediano", f"€{mercato_med:,.0f}/m²".replace(',', '.'))
        
        with col3:
            gap_assoluto = mercato_med - omi_med
            st.metric("Gap", 
                     f"{gap_percentuale:+.1f}%".replace('.', ','),
                     delta=f"€{gap_assoluto:,.0f}/m²".replace(',', '.'))
        
        st.markdown("---")
        
        # Interpretazione strategica
        st.markdown("**🎯 Interpretazione Strategica:**")
        
        if gap_percentuale > 50:
            st.error("""GAP MOLTO ALTO (+50% vs OMI)
            
**Significato:**
- Mercato con forte premium su nuove costruzioni
- Possibile sopravvalutazione
- Alto rischio pricing

**Raccomandazione:**
- Verificare qualità effettiva immobili
- Rischio di correzione prezzi
- Considera pricing conservativo
            """)
        elif gap_percentuale > 30:
            st.warning("""GAP SIGNIFICATIVO (+30-50% vs OMI)
            
**Significato:**
- Premium pricing per nuove costruzioni
- Mercato accetta sovrapprezzo elevato
- Margini interessanti ma attenzione

**Raccomandazione:**
- Giustifica premium con alta qualità
- Finiture e servizi eccellenti
- Marketing forte
            """)
        elif gap_percentuale > 15:
            st.success("""GAP NORMALE (+15-30% vs OMI)
            
**Significato:**
- Premium standard nuove costruzioni
- Mercato equilibrato
- Margini sani

**Raccomandazione:**
- Sweet spot ideale per sviluppo
- Buon bilanciamento prezzo/qualità
- Vendibilità ottimale
            """)
        else:
            st.info("""GAP BASSO (<15% vs OMI)
            
**Significato:**
- Prezzi allineati a valori rogiti
- Mercato competitivo
- Margini contenuti

**Raccomandazione:**
- Ottimizza costi costruzione
- Efficienza operativa fondamentale
- Volume vendite importante
            """)
        
        st.markdown("---")
        
        # ========================================
        # 5. ANALISI AGENZIE - VERSIONE CORRETTA
        # ========================================
        st.subheader("🏢 5. Analisi Agenzie/Operatori")
        
        if stats_immobiliare.get('dataframe') is not None:
            st.markdown("**Principali operatori nella zona:**")
            
            # Calcola agenzie dal dataframe direttamente
            df = stats_immobiliare['dataframe']
            agenzie_stats = df.groupby('agenzia').size().reset_index(name='count')
            agenzie_stats = agenzie_stats.sort_values('count', ascending=False).head(10)
            
            if len(agenzie_stats) > 0:
                agenzie_data = []
                for _, row in agenzie_stats.iterrows():
                    percentuale = (row['count'] / n_app * 100)
                    agenzie_data.append({
                        'Agenzia': row['agenzia'],
                        'N° Appartamenti': int(row['count']),
                        '% Mercato': f"{percentuale:.1f}%"
                    })
                
                agenzie_df = pd.DataFrame(agenzie_data)
                
                st.dataframe(
                    agenzie_df,
                    use_container_width=True,
                    hide_index=True
                )
                
                # Analisi concentrazione
                top3_count = agenzie_stats.head(3)['count'].sum()
                top3_share = (top3_count / n_app * 100)
                
                st.markdown("---")
                st.markdown("**📊 Concentrazione Mercato:**")
                
                col1, col2 = st.columns(2)
                
                with col1:
                    st.metric("Top 3 Agenzie", f"{top3_share:.1f}%")
                
                with col2:
                    if top3_share > 60:
                        st.warning("Mercato concentrato - Pochi operatori dominanti")
                    elif top3_share > 40:
                        st.info("Mercato moderato - Mix operatori grandi/piccoli")
                    else:
                        st.success("Mercato frammentato - Molti piccoli operatori")


# ----------------------------------------
# TAB 5: ANALISI AI
# ----------------------------------------
with tab5:
    st.header("🤖 Analisi AI con Claude")
    
    # Controlla se modulo AI disponibile
    if not CLAUDE_AVAILABLE:
        st.error("⚠️ Modulo AI non disponibile")
        st.info("""
        **Il file `claude_analyzer.py` non è stato trovato nel repository.**
        
        Per abilitare l'analisi AI:
        1. Crea il file `claude_analyzer.py` con le funzioni necessarie
        2. Carica il file nel repository GitHub
        3. Fai push e riavvia l'app Streamlit
        """)
        st.stop()
    
    # Verifica dati disponibili
    if 'analisi_data' not in st.session_state:
        st.warning("⚠️ Esegui prima un'analisi per vedere l'analisi AI")
        st.stop()
    
    data = st.session_state.analisi_data
    
    # Verifica se l'utente ha abilitato l'AI durante l'analisi
    if 'analisi_ai' in data and data['analisi_ai']:
        risultato = data['analisi_ai']
        
        # Controlla se è stata disabilitata dall'utente
        if not risultato.get('success') and 'disabilitata dall\'utente' in risultato.get('error', ''):
            st.info("""
            ℹ️ **Analisi AI disabilitata**
            
            Hai scelto di non eseguire l'analisi AI durante questa ricerca.
            
            Per ottenere un'analisi AI:
            1. Torna alla Home
            2. Seleziona la checkbox "🤖 Abilita Analisi AI"
            3. Avvia una nuova analisi
            """)
            st.stop()
        
        if risultato['success']:
            
            # Gap Analysis
            if risultato.get('gap_analysis'):
                st.subheader("📊 Gap Analysis")
                gap = risultato['gap_analysis']
                
                col1, col2, col3 = st.columns(3)
                col1.metric("OMI Mediano", f"€{gap['omi_mediano']:,.0f}/m²".replace(',', '.'))
                col2.metric("Mercato Mediano", f"€{gap['mercato_mediano']:,.0f}/m²".replace(',', '.'))
                col3.metric("Gap", f"{gap['gap_percentuale']:+.1f}%".replace('.', ','),
                           delta=f"€{gap['gap_assoluto']:,.0f}/m²".replace(',', '.'))
                
                st.markdown("---")
            
            # Analisi completa
            st.subheader("📝 Analisi Completa")
            st.markdown(risultato['analisi_completa'])
            
            st.markdown("---")
            
            # Download analisi come testo
            st.subheader("💾 Salva Analisi")
            
            # Prepara testo completo
            testo_completo = f"""ANALISI AI - Planet AI
{'='*70}
Località: {data['via']}, {data['comune']}
Data: {datetime.now().strftime('%d/%m/%Y %H:%M')}
{'='*70}

"""
            
            if risultato.get('gap_analysis'):
                gap = risultato['gap_analysis']
                testo_completo += f"""
GAP ANALYSIS
------------
OMI Mediano: €{gap['omi_mediano']:,.0f}/m²
Mercato Mediano: €{gap['mercato_mediano']:,.0f}/m²
Gap: {gap['gap_percentuale']:+.1f}% (€{gap['gap_assoluto']:,.0f}/m²)

"""
            
            testo_completo += f"""
ANALISI DETTAGLIATA
-------------------
{risultato['analisi_completa']}

"""
            
            if risultato.get('raccomandazioni'):
                testo_completo += "\nRACCOMANDAZIONI\n---------------\n"
                for i, racc in enumerate(risultato['raccomandazioni'], 1):
                    testo_completo += f"{i}. {racc}\n"
            
            testo_completo += f"\n{'='*70}\nGenerato da Planet AI - Powered by Claude (Anthropic)\n"
            
            st.download_button(
                label="📥 Scarica Analisi AI (TXT)",
                data=testo_completo,
                file_name=f"analisi_ai_{data['comune']}_{datetime.now().strftime('%Y%m%d_%H%M%S')}.txt",
                mime="text/plain",
                key="download_ai",
                use_container_width=True
            )
            
            # Raccomandazioni
            if risultato.get('raccomandazioni'):
                st.markdown("---")
                st.subheader("💡 Raccomandazioni")
                for i, racc in enumerate(risultato['raccomandazioni'], 1):
                    st.markdown(f"**{i}.** {racc}")
        
        else:
            st.error(f"❌ Errore nell'analisi AI: {risultato.get('error', 'Errore sconosciuto')}")
    
    else:
        st.info("⏳ L'analisi AI verrà generata automaticamente al prossimo avvio dell'analisi.")
        st.caption("Se hai appena fatto un'analisi e non vedi i risultati, verifica che l'API key sia configurata correttamente.")

# ----------------------------------------
# TAB 6: REPORT
# ----------------------------------------
with tab6:
    st.header("📄 Download Report e Dati")
    
    # Controlla se ci sono dati salvati
    if 'analisi_data' not in st.session_state:
        st.warning("⚠️ Esegui prima un'analisi per generare il report.")
        st.stop()
    
    # Recupera dati da session state
    data = st.session_state.analisi_data
    comune_report = data['comune']
    appartamenti_report = data['appartamenti']
    report_data = data.get('report_data')
    report_filename = data.get('report_filename')
    
    # REPORT WORD
    st.subheader("📝 Report Word Completo")
    st.write("Report con dati OMI, analisi mercato Immobiliare.it, analisi Developer e confronto completo.")
    
    if report_data and report_filename:
        
        col1, col2 = st.columns([3, 1])
        with col1:
            st.download_button(
                label="📥 Scarica Report Word",
                data=report_data,
                file_name=report_filename,
                mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document",
                key="download_report",
                use_container_width=True
            )
        with col2:
            st.caption(f"📄 {report_filename}")
        
        st.success("✅ Report pronto per il download!")
        
    else:
        st.error("❌ Errore nella generazione del report. Riprova l'analisi.")
    
    st.markdown("---")
    
    # Download CSV
    st.subheader("📊 Dati Raw (CSV)")
    st.write("Dati grezzi degli appartamenti per analisi personalizzate.")
    
    if appartamenti_report:
        df_export = pd.DataFrame(appartamenti_report)
        csv = df_export.to_csv(index=False, encoding='utf-8-sig')
        
        col1, col2 = st.columns([3, 1])
        with col1:
            st.download_button(
                label="📥 Scarica CSV Appartamenti",
                data=csv,
                file_name=f"appartamenti_{comune_report}_{datetime.now().strftime('%Y%m%d_%H%M%S')}.csv",
                mime="text/csv",
                key="download_csv",
                use_container_width=True
            )
        with col2:
            st.caption(f"📊 {len(appartamenti_report)} record")
        
        st.success("✅ CSV pronto per il download!")
    else:
        st.info("Nessun dato disponibile per l'export CSV.")

# ========================================
# FOOTER
# ========================================
st.markdown("---")
st.caption("Planet AI - Analisi Immobiliare | Dati OMI: Agenzia delle Entrate QI 2025/1 | Mercato: Immobiliare.it")