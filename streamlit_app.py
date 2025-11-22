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
    
    col_btn1, col_btn2 = st.columns([1, 4])
    with col_btn1:
        submit = st.form_submit_button("🔍 Avvia Analisi", width="stretch")

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

lat, lon = geocode_indirizzo(comune, via)

st.success(f"✅ Coordinate: {lat:.6f}, {lon:.6f}")

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

# 5. GENERA REPORT AUTOMATICAMENTE
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
        output_dir=REPORTS_DIR
    )
    
    # Leggi il file in memoria
    with open(report_filepath, 'rb') as f:
        report_data = f.read()
    
    report_filename = os.path.basename(report_filepath)
    
except Exception as e:
    print(f"Errore generazione report: {e}")
    report_data = None
    report_filename = None

# 6. ANALISI AI AUTOMATICA (solo se disponibile)
if CLAUDE_AVAILABLE:
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
    risultato_ai = {'success': False, 'error': 'Modulo AI non disponibile'}

status_text.text("✅ Completato!")

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
tab1, tab2, tab3, tab4, tab5 = st.tabs(["📊 Dati OMI", "🏠 Immobiliare.it", "📈 Confronto", "🤖 Analisi AI", "📄 Report"])

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
        
        col_info1, col_info2 = st.columns(2)
        with col_info1:
            st.metric("Progetti immobiliari", stats_immobiliare['n_progetti'])
        with col_info2:
            st.metric("App. per progetto (media)", f"{stats_immobiliare['n_appartamenti'] / stats_immobiliare['n_progetti']:.1f}")
        
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
            'prezzo_mq': 'mean',
            'progetto_id': 'nunique'
        }).reset_index()
        
        agenzie.columns = ['Agenzia', 'N° Appartamenti', 'Prezzo Medio', 'MQ Medio', 'Prezzo/mq Medio', 'N° Progetti']
        agenzie = agenzie.sort_values('N° Appartamenti', ascending=False)
        
        # Formatta per display
        agenzie_display = agenzie.copy()
        agenzie_display['Prezzo Medio'] = agenzie_display['Prezzo Medio'].apply(lambda x: f"€{x:,.0f}".replace(',', '.'))
        agenzie_display['MQ Medio'] = agenzie_display['MQ Medio'].apply(lambda x: f"{x:.0f} m²")
        agenzie_display['Prezzo/mq Medio'] = agenzie_display['Prezzo/mq Medio'].apply(lambda x: f"€{x:,.0f}/m²".replace(',', '.'))
        
        # Riordina colonne
        agenzie_display = agenzie_display[['Agenzia', 'N° Progetti', 'N° Appartamenti', 'Prezzo Medio', 'MQ Medio', 'Prezzo/mq Medio']]
        
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
        
        st.plotly_chart(fig, width="stretch")
        
    else:
        st.warning("⚠️ Dati insufficienti per il confronto.")

# ----------------------------------------
# TAB 4: ANALISI AI
# ----------------------------------------
with tab4:
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
    
    # Controlla API key
    api_key = get_api_key()
    
    if not api_key:
        st.warning("⚠️ API key Anthropic non configurata")
        st.info("""
        **Come configurare:**
        
        **Locale:**
        Crea file `.env` con:
        ```
        ANTHROPIC_API_KEY=sk-ant-api03-...
        ```
        
        **Streamlit Cloud:**
        Settings → Secrets → Aggiungi:
        ```toml
        ANTHROPIC_API_KEY = "sk-ant-api03-..."
        ```
        
        **Ottieni API key:** https://console.anthropic.com
        """)
        st.stop()
    
    # Verifica dati disponibili
    if 'analisi_data' not in st.session_state:
        st.warning("⚠️ Esegui prima un'analisi per vedere l'analisi AI")
        st.stop()
    
    data = st.session_state.analisi_data
    
    # Mostra analisi (già generata automaticamente durante l'elaborazione)
    if 'analisi_ai' in data and data['analisi_ai']:
        risultato = data['analisi_ai']
        
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
                width="stretch"
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
# TAB 5: REPORT
# ----------------------------------------
with tab5:
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
    st.write("Report con dati OMI, analisi mercato Immobiliare.it e confronto completo.")
    
    if report_data and report_filename:
        
        col1, col2 = st.columns([3, 1])
        with col1:
            st.download_button(
                label="📥 Scarica Report Word",
                data=report_data,
                file_name=report_filename,
                mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document",
                key="download_report",
                width="stretch"
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
                width="stretch"
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