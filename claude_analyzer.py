"""
Claude Analyzer - Analisi AI per dati immobiliari
==================================================
Utilizza l'API Anthropic Claude per generare analisi approfondite
confrontando dati OMI e mercato Immobiliare.it
"""

import os
import json
from typing import Dict, Optional, List
import anthropic


def get_api_key() -> Optional[str]:
    """
    Recupera la API key di Anthropic.
    
    Cerca in ordine:
    1. Variabile d'ambiente ANTHROPIC_API_KEY
    2. Streamlit secrets (se disponibile)
    3. File .env (se disponibile)
    
    Returns:
        str: API key se trovata, None altrimenti
    """
    # 1. Variabile d'ambiente
    api_key = os.getenv('ANTHROPIC_API_KEY')
    if api_key:
        return api_key
    
    # 2. Streamlit secrets
    try:
        import streamlit as st
        if 'ANTHROPIC_API_KEY' in st.secrets:
            return st.secrets['ANTHROPIC_API_KEY']
    except:
        pass
    
    # 3. File .env
    try:
        from dotenv import load_dotenv
        load_dotenv()
        api_key = os.getenv('ANTHROPIC_API_KEY')
        if api_key:
            return api_key
    except:
        pass
    
    return None


def calcola_gap_analysis(zona_omi: Dict, stats_immobiliare: Dict) -> Optional[Dict]:
    """
    Calcola il gap tra valori OMI e mercato.
    
    Args:
        zona_omi: Dati zona OMI
        stats_immobiliare: Statistiche mercato Immobiliare.it
    
    Returns:
        Dict con gap analysis o None se dati insufficienti
    """
    if not zona_omi or not stats_immobiliare:
        return None
    
    if zona_omi.get('val_med_mq') is None:
        return None
    
    if not stats_immobiliare.get('prezzo_mq') or stats_immobiliare['prezzo_mq'].get('mediano') is None:
        return None
    
    omi_mediano = zona_omi['val_med_mq']
    mercato_mediano = stats_immobiliare['prezzo_mq']['mediano']
    
    gap_assoluto = mercato_mediano - omi_mediano
    gap_percentuale = (gap_assoluto / omi_mediano) * 100 if omi_mediano > 0 else 0
    
    return {
        'omi_mediano': omi_mediano,
        'mercato_mediano': mercato_mediano,
        'gap_assoluto': gap_assoluto,
        'gap_percentuale': gap_percentuale
    }


def prepara_prompt_analisi(
    comune: str,
    via: str,
    zona_omi: Optional[Dict],
    stats_immobiliare: Optional[Dict],
    gap_analysis: Optional[Dict]
) -> str:
    """
    Prepara il prompt per Claude con tutti i dati.
    
    Args:
        comune: Nome comune
        via: Nome via
        zona_omi: Dati OMI
        stats_immobiliare: Statistiche mercato
        gap_analysis: Analisi gap
    
    Returns:
        str: Prompt formattato
    """
    prompt = f"""Sei un esperto analista immobiliare italiano specializzato nell'analisi di mercato.

Devi analizzare i dati di una zona immobiliare e fornire un'analisi professionale approfondita.

**LOCALITÀ ANALIZZATA:**
- Comune: {comune}
- Via/Zona: {via}

---

**DATI OMI (Osservatorio Mercato Immobiliare - Agenzia delle Entrate):**
"""
    
    if zona_omi and zona_omi.get('val_med_mq'):
        prompt += f"""
- Zona OMI: {zona_omi['zona_codice']} - {zona_omi['zona_descrizione']}
- Valori €/m² (rogiti):
  * Minimo: €{zona_omi['val_min_mq']:,.0f}
  * Mediano: €{zona_omi['val_med_mq']:,.0f}
  * Massimo: €{zona_omi['val_max_mq']:,.0f}
"""
    else:
        prompt += "\n- Dati OMI non disponibili per questa zona\n"
    
    prompt += "\n---\n\n**DATI MERCATO (Immobiliare.it - Nuove Costruzioni):**\n"
    
    if stats_immobiliare and stats_immobiliare.get('n_appartamenti', 0) > 0:
        prompt += f"""
- Numero appartamenti analizzati: {stats_immobiliare.get('n_appartamenti', 'N/D')}
- Progetti immobiliari: {stats_immobiliare.get('n_progetti', 'N/D')}
"""
        
        # Prezzi totali (se disponibili)
        if stats_immobiliare.get('prezzo_totale'):
            prezzo_tot = stats_immobiliare['prezzo_totale']
            prompt += f"""
**Prezzi totali:**
- Minimo: €{prezzo_tot.get('min', 0):,.0f}
- Mediano: €{prezzo_tot.get('mediano', 0):,.0f}
- Massimo: €{prezzo_tot.get('max', 0):,.0f}
"""
        
        # Superfici (se disponibili)
        if stats_immobiliare.get('superficie'):
            superficie = stats_immobiliare['superficie']
            prompt += f"""
**Superfici (m²):**
- Minima: {superficie.get('min', 0)} m²
- Mediana: {superficie.get('mediano', 0)} m²
- Massima: {superficie.get('max', 0)} m²
"""
        
        # Prezzi al m² (se disponibili)
        if stats_immobiliare.get('prezzo_mq'):
            prezzo_mq = stats_immobiliare['prezzo_mq']
            prompt += f"""
**Prezzi al m²:**
- Minimo: €{prezzo_mq.get('min', 0):,.0f}/m²
- Mediano: €{prezzo_mq.get('mediano', 0):,.0f}/m²
- Massimo: €{prezzo_mq.get('max', 0):,.0f}/m²
"""
        
        # Agenzie immobiliari (se disponibili)
        if stats_immobiliare.get('top_agenzie'):
            prompt += "\n**Agenzie immobiliari:**\n"
            for agenzia in stats_immobiliare['top_agenzie'][:5]:
                nome_agenzia = agenzia.get('agenzia', 'N/D')
                count_agenzia = agenzia.get('count', 0)
                prompt += f"- {nome_agenzia}: {count_agenzia} appartamenti\n"
    else:
        prompt += "\n- Nessun dato disponibile dal mercato Immobiliare.it\n"
    
    if gap_analysis:
        prompt += f"""
---

**GAP ANALYSIS:**
- OMI Mediano: €{gap_analysis['omi_mediano']:,.0f}/m²
- Mercato Mediano: €{gap_analysis['mercato_mediano']:,.0f}/m²
- Gap Assoluto: €{gap_analysis['gap_assoluto']:,.0f}/m²
- Gap Percentuale: {gap_analysis['gap_percentuale']:+.1f}%
"""
    
    prompt += """

---

**COMPITO:**

Fornisci un'analisi professionale approfondita in formato Markdown che includa:

1. **Sintesi Esecutiva** (2-3 paragrafi)
   - Overview della zona e del mercato
   - Principale insight dal confronto OMI vs Mercato

2. **Analisi Gap OMI vs Mercato**
   - Interpretazione del gap percentuale
   - Cosa significa questo gap per investitori/acquirenti
   - Fattori che potrebbero spiegare la differenza

3. **Analisi del Mercato Immobiliare.it**
   - Valutazione dell'offerta (numero progetti, distribuzione prezzi)
   - Analisi dei prezzi al m² (alta/media/bassa gamma)
   - Concentrazione agenzie (mercato frammentato o concentrato)

4. **Posizionamento della Zona**
   - Come si posiziona rispetto ai valori OMI
   - Fascia di mercato (luxury, medio-alta, media, accessibile)
   - Dinamiche di prezzo rilevate

5. **Raccomandazioni Strategiche** (lista puntata chiara)
   - Per acquirenti/investitori
   - Per venditori/costruttori
   - Timing di mercato

**STILE:**
- Professionale ma accessibile
- Dati concreti e numerici
- Evita genericità
- Usa terminologia tecnica immobiliare italiana
- Sii specifico sui numeri
- Non usare emoji

**IMPORTANTE:** 
- Se i dati sono limitati, menzionalo e fornisci comunque l'analisi disponibile
- Se il gap è significativo (>±10%), spiegalo in dettaglio
- Contestualizza sempre i numeri rispetto al mercato locale
"""
    
    return prompt


def analizza_con_ai(
    comune: str,
    via: str,
    zona_omi: Optional[Dict],
    stats_immobiliare: Optional[Dict]
) -> Dict:
    """
    Esegue l'analisi AI tramite Claude.
    
    Args:
        comune: Nome comune
        via: Nome via
        zona_omi: Dati OMI (dict o None)
        stats_immobiliare: Statistiche mercato (dict o None)
    
    Returns:
        Dict con risultati analisi:
        {
            'success': bool,
            'analisi_completa': str (markdown),
            'gap_analysis': dict (se disponibile),
            'raccomandazioni': list[str] (se disponibile),
            'error': str (se success=False)
        }
    """
    # Recupera API key
    api_key = get_api_key()
    
    if not api_key:
        return {
            'success': False,
            'error': 'API key Anthropic non configurata'
        }
    
    # Verifica che ci siano dati da analizzare
    if not zona_omi and not stats_immobiliare:
        return {
            'success': False,
            'error': 'Nessun dato disponibile per l\'analisi'
        }
    
    try:
        # Debug: stampa struttura dati ricevuti
        print(f"[DEBUG] zona_omi keys: {zona_omi.keys() if zona_omi else 'None'}")
        print(f"[DEBUG] stats_immobiliare keys: {stats_immobiliare.keys() if stats_immobiliare else 'None'}")
        
        # Calcola gap analysis se possibile
        gap_analysis = calcola_gap_analysis(zona_omi, stats_immobiliare)
        
        # Prepara prompt
        prompt = prepara_prompt_analisi(
            comune=comune,
            via=via,
            zona_omi=zona_omi,
            stats_immobiliare=stats_immobiliare,
            gap_analysis=gap_analysis
        )
        
        # Inizializza client Anthropic
        client = anthropic.Anthropic(api_key=api_key)
        
        # Chiamata API
        message = client.messages.create(
            model="claude-sonnet-4-20250514",
            max_tokens=4000,
            temperature=0.7,
            messages=[
                {
                    "role": "user",
                    "content": prompt
                }
            ]
        )
        
        # Estrai il testo della risposta
        analisi_completa = message.content[0].text
        
        # Prova a estrarre raccomandazioni (cerca sezione con lista puntata)
        raccomandazioni = []
        if "Raccomandazioni" in analisi_completa or "RACCOMANDAZIONI" in analisi_completa:
            # Estrazione semplice delle raccomandazioni
            lines = analisi_completa.split('\n')
            in_raccomandazioni = False
            for line in lines:
                if 'raccomandazioni' in line.lower():
                    in_raccomandazioni = True
                    continue
                if in_raccomandazioni:
                    line = line.strip()
                    if line.startswith('-') or line.startswith('*') or line.startswith('•'):
                        racc = line.lstrip('-*•').strip()
                        if racc:
                            raccomandazioni.append(racc)
                    elif line.startswith('#'):
                        # Nuova sezione, stop
                        break
        
        return {
            'success': True,
            'analisi_completa': analisi_completa,
            'gap_analysis': gap_analysis,
            'raccomandazioni': raccomandazioni if raccomandazioni else None
        }
        
    except anthropic.AuthenticationError:
        return {
            'success': False,
            'error': 'API key non valida. Verifica la configurazione.'
        }
    except anthropic.RateLimitError:
        return {
            'success': False,
            'error': 'Limite rate API raggiunto. Riprova tra qualche minuto.'
        }
    except Exception as e:
        return {
            'success': False,
            'error': f'Errore durante l\'analisi AI: {str(e)}'
        }


if __name__ == "__main__":
    """
    Test del modulo con dati di esempio
    """
    print("🧪 Test Claude Analyzer")
    print("="*50)
    
    # Test get_api_key
    api_key = get_api_key()
    if api_key:
        print(f"✅ API Key trovata: {api_key[:10]}...")
    else:
        print("❌ API Key non trovata")
    
    # Dati di test
    zona_omi_test = {
        'comune': 'Como',
        'provincia': 'CO',
        'zona_codice': 'B1',
        'zona_descrizione': 'Centro storico',
        'val_min_mq': 2500,
        'val_med_mq': 3200,
        'val_max_mq': 4000
    }
    
    stats_test = {
        'n_appartamenti': 15,
        'n_progetti': 3,
        'prezzo_totale': {'min': 250000, 'mediano': 450000, 'max': 850000},
        'superficie': {'min': 65, 'mediano': 95, 'max': 150},
        'prezzo_mq': {'min': 3200, 'mediano': 4500, 'max': 6000},
        'top_agenzie': [
            {'agenzia': 'Immobiliare Como Centro', 'count': 8},
            {'agenzia': 'Luxury Homes', 'count': 4}
        ]
    }
    
    print("\n📊 Calcolo Gap Analysis...")
    gap = calcola_gap_analysis(zona_omi_test, stats_test)
    if gap:
        print(f"OMI Mediano: €{gap['omi_mediano']:,.0f}/m²")
        print(f"Mercato Mediano: €{gap['mercato_mediano']:,.0f}/m²")
        print(f"Gap: {gap['gap_percentuale']:+.1f}%")
    
    print("\n✅ Test completato!")