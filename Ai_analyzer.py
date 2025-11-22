"""
AI Analyzer - Analisi immobiliare con Claude (Anthropic)
"""

import os
from typing import Dict, Optional
import anthropic


def get_api_key() -> Optional[str]:
    """
    Recupera API key da environment o Streamlit secrets
    """
    # Prova environment variable
    api_key = os.getenv('ANTHROPIC_API_KEY')
    
    # Prova Streamlit secrets (solo se streamlit è disponibile)
    if not api_key:
        try:
            import streamlit as st
            if hasattr(st, 'secrets') and 'ANTHROPIC_API_KEY' in st.secrets:
                api_key = st.secrets['ANTHROPIC_API_KEY']
        except:
            pass
    
    return api_key


def analizza_con_ai(
    comune: str,
    via: str,
    zona_omi: Optional[Dict],
    stats_immobiliare: Optional[Dict]
) -> Dict:
    """
    Analizza i dati immobiliari con Claude AI
    
    Args:
        comune: Nome comune
        via: Nome via
        zona_omi: Dati OMI (dict con val_min_mq, val_med_mq, val_max_mq)
        stats_immobiliare: Statistiche Immobiliare.it
    
    Returns:
        Dict con:
        - analisi_completa (str): Analisi testuale completa
        - gap_analysis (dict): Numeri chiave
        - raccomandazioni (list): Lista raccomandazioni
        - pricing_strategy (dict): Prezzi consigliati
        - success (bool): Se l'analisi è riuscita
        - error (str): Eventuale errore
    """
    
    api_key = get_api_key()
    
    if not api_key:
        return {
            'success': False,
            'error': 'API key Anthropic non configurata',
            'analisi_completa': None,
        }
    
    # Prepara i dati per l'AI
    prompt = _prepara_prompt(comune, via, zona_omi, stats_immobiliare)
    
    try:
        # Chiama Claude API
        client = anthropic.Anthropic(api_key=api_key)
        
        message = client.messages.create(
            model="claude-sonnet-4-20250514",
            max_tokens=2000,
            temperature=0.7,
            messages=[
                {
                    "role": "user",
                    "content": prompt
                }
            ]
        )
        
        # Estrai risposta
        analisi_testo = message.content[0].text
        
        # Parsing della risposta
        risultato = _parse_risposta_ai(analisi_testo, zona_omi, stats_immobiliare)
        risultato['success'] = True
        risultato['analisi_completa'] = analisi_testo
        
        return risultato
        
    except Exception as e:
        return {
            'success': False,
            'error': str(e),
            'analisi_completa': None,
        }


def _prepara_prompt(
    comune: str,
    via: str,
    zona_omi: Optional[Dict],
    stats_immobiliare: Optional[Dict]
) -> str:
    """
    Prepara il prompt per Claude con tutti i dati
    """
    
    prompt = f"""Sei un esperto analista immobiliare italiano. Analizza i seguenti dati per una zona specifica e fornisci insights professionali.

LOCALITÀ: {via}, {comune}

"""
    
    # Dati OMI
    if zona_omi and zona_omi.get('val_med_mq'):
        prompt += f"""DATI OMI (Agenzia delle Entrate - valori ufficiali rogiti):
- Zona: {zona_omi['zona_codice']}
- Descrizione: {zona_omi['zona_descrizione']}
- Comune: {zona_omi['comune']} ({zona_omi['provincia']})
- Valori €/mq:
  * Minimo: €{zona_omi['val_min_mq']:,.0f}/m²
  * Mediano: €{zona_omi['val_med_mq']:,.0f}/m²
  * Massimo: €{zona_omi['val_max_mq']:,.0f}/m²

"""
    else:
        prompt += "DATI OMI: Non disponibili per questa zona\n\n"
    
    # Dati Immobiliare.it
    if stats_immobiliare and stats_immobiliare['n_appartamenti'] > 0:
        prompt += f"""DATI MERCATO (Immobiliare.it - nuove costruzioni):
- Appartamenti analizzati: {stats_immobiliare['n_appartamenti']}
- Progetti immobiliari: {stats_immobiliare['n_progetti']}
- Media appartamenti per progetto: {stats_immobiliare['n_appartamenti'] / stats_immobiliare['n_progetti']:.1f}

PREZZI:
- Prezzo medio: €{stats_immobiliare['prezzo']['medio']:,.0f}
- Prezzo mediano: €{stats_immobiliare['prezzo']['mediano']:,.0f}
- Range: €{stats_immobiliare['prezzo']['min']:,.0f} - €{stats_immobiliare['prezzo']['max']:,.0f}

SUPERFICI:
- Superficie media: {stats_immobiliare['mq']['medio']:.0f} m²
- Superficie mediana: {stats_immobiliare['mq']['mediano']:.0f} m²
- Range: {stats_immobiliare['mq']['min']:.0f} - {stats_immobiliare['mq']['max']:.0f} m²

PREZZO/MQ:
- Prezzo/mq medio: €{stats_immobiliare['prezzo_mq']['medio']:,.0f}/m²
- Prezzo/mq mediano: €{stats_immobiliare['prezzo_mq']['mediano']:,.0f}/m²
- Range: €{stats_immobiliare['prezzo_mq']['min']:,.0f} - €{stats_immobiliare['prezzo_mq']['max']:,.0f}/m²

"""
        
        # Agenzie top 3
        if stats_immobiliare.get('agenzie'):
            agenzie_list = sorted(
                stats_immobiliare['agenzie'],
                key=lambda x: x[1],  # N° Appartamenti
                reverse=True
            )[:3]
            
            prompt += "PRINCIPALI AGENZIE:\n"
            for ag in agenzie_list:
                prompt += f"- {ag[0]}: {int(ag[1])} appartamenti, {int(ag[3])} progetti\n"
            prompt += "\n"
    else:
        prompt += "DATI MERCATO: Nessun annuncio trovato su Immobiliare.it\n\n"
    
    # Richiesta analisi
    prompt += """
RICHIESTA:
Fornisci un'analisi professionale strutturata così:

1. GAP ANALYSIS
Confronta i valori OMI con i prezzi di mercato. Calcola il gap percentuale e spiega cosa significa.

2. VALUTAZIONE ZONA
- È una zona calda, equilibrata o fredda?
- Analizza domanda vs offerta
- Considera il numero di progetti attivi
- Valuta la competizione tra agenzie

3. PRICING STRATEGY
Per una nuova costruzione in questa zona, suggerisci:
- Prezzo/mq PRUDENTE (conservativo)
- Prezzo/mq CENTRALE (realistico)
- Prezzo/mq AGGRESSIVO (ottimistico)

Giustifica ogni range in base ai dati.

4. OPPORTUNITÀ E RISCHI
- Opportunità: cosa rende interessante questa zona
- Rischi: cosa potrebbe impattare negativamente

5. RACCOMANDAZIONI
3-5 azioni concrete per un investitore/costruttore.

Usa linguaggio professionale ma chiaro. Fornisci numeri precisi dove possibile.
"""
    
    return prompt


def _parse_risposta_ai(
    analisi_testo: str,
    zona_omi: Optional[Dict],
    stats_immobiliare: Optional[Dict]
) -> Dict:
    """
    Estrae informazioni strutturate dalla risposta AI
    """
    
    risultato = {
        'gap_analysis': {},
        'raccomandazioni': [],
        'pricing_strategy': {},
    }
    
    # Calcola gap se possibile
    if zona_omi and zona_omi.get('val_med_mq') and stats_immobiliare:
        omi_med = zona_omi['val_med_mq']
        mercato_med = stats_immobiliare['prezzo_mq']['mediano']
        gap_pct = ((mercato_med - omi_med) / omi_med) * 100
        
        risultato['gap_analysis'] = {
            'omi_mediano': omi_med,
            'mercato_mediano': mercato_med,
            'gap_assoluto': mercato_med - omi_med,
            'gap_percentuale': gap_pct,
        }
    
    # Estrai raccomandazioni (cerca linee che iniziano con -, •, o numeri)
    linee = analisi_testo.split('\n')
    in_raccomandazioni = False
    
    for linea in linee:
        if 'RACCOMANDAZIONI' in linea.upper() or 'AZIONI' in linea.upper():
            in_raccomandazioni = True
            continue
        
        if in_raccomandazioni:
            linea_clean = linea.strip()
            if linea_clean and (linea_clean.startswith('-') or 
                               linea_clean.startswith('•') or 
                               linea_clean[0].isdigit()):
                # Rimuovi bullet/numeri iniziali
                raccomandazione = linea_clean.lstrip('-•0123456789. ').strip()
                if raccomandazione:
                    risultato['raccomandazioni'].append(raccomandazione)
    
    return risultato