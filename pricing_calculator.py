"""
Pricing Calculator - Calcolo intelligente pricing benchmark
============================================================
Considera: saturazione, gap OMI, premium nuova costruzione, rischio
"""

from typing import Dict, Optional


def calcola_pricing_intelligente(
    zona_omi: Optional[Dict],
    stats_immobiliare: Optional[Dict]
) -> Dict:
    """
    Calcola pricing benchmark intelligente per nuove costruzioni.
    
    Formula:
    Base = OMI Mediano (transazioni reali)
    + Premium Nuova Costruzione (15-25%)
    ± Aggiustamento Saturazione (-10% a +10%)
    ± Aggiustamento Gap Mercato
    
    Returns:
        Dict con:
        - prezzo_minimo: Prezzo entry conservative
        - prezzo_ottimale: Sweet spot consigliato
        - prezzo_massimo: Ceiling realistico
        - logica_calcolo: Spiegazione dettagliata
        - fattori: Dict con i vari aggiustamenti applicati
    """
    
    if not zona_omi or not zona_omi.get('val_med_mq'):
        return {
            'prezzo_minimo': None,
            'prezzo_ottimale': None,
            'prezzo_massimo': None,
            'logica_calcolo': 'Dati OMI non disponibili - impossibile calcolare pricing',
            'fattori': {}
        }
    
    # BASE: OMI Mediano (transazioni reali)
    base_omi = zona_omi['val_med_mq']
    
    # FATTORI DI AGGIUSTAMENTO
    fattori = {
        'base_omi': base_omi,
        'premium_nuova_costruzione': 0.0,
        'aggiustamento_saturazione': 0.0,
        'aggiustamento_gap': 0.0,
        'fattore_rischio': 0.0
    }
    
    # 1. PREMIUM NUOVA COSTRUZIONE (standard mercato italiano)
    # Baseline: +20% per nuove costruzioni vs usato
    premium_base = 0.20
    
    # Se abbiamo dati mercato, calibriamo il premium
    if stats_immobiliare and stats_immobiliare.get('n_appartamenti', 0) > 0:
        mercato_med = stats_immobiliare['prezzo_mq']['mediano']
        gap_attuale = ((mercato_med - base_omi) / base_omi)
        
        # Se il gap attuale è molto alto, il mercato accetta premium elevati
        if gap_attuale > 0.40:  # +40%
            premium_base = 0.25  # +25%
        elif gap_attuale > 0.25:  # +25%
            premium_base = 0.22  # +22%
        elif gap_attuale < 0.10:  # +10%
            premium_base = 0.15  # +15% (mercato conservativo)
    
    fattori['premium_nuova_costruzione'] = premium_base
    
    # 2. AGGIUSTAMENTO SATURAZIONE MERCATO
    aggiustamento_saturazione = 0.0
    
    if stats_immobiliare and stats_immobiliare.get('n_appartamenti'):
        n_appartamenti = stats_immobiliare['n_appartamenti']
        
        if n_appartamenti < 10:
            # MERCATO LIBERO - Poca concorrenza, puoi chiedere di più
            aggiustamento_saturazione = 0.08  # +8%
            saturazione_label = "LIBERO (poca concorrenza)"
            
        elif n_appartamenti < 30:
            # MERCATO MEDIO - Concorrenza normale
            aggiustamento_saturazione = 0.00  # Neutro
            saturazione_label = "MEDIO (concorrenza normale)"
            
        elif n_appartamenti < 60:
            # MERCATO AFFOLLATO - Devi essere competitivo
            aggiustamento_saturazione = -0.05  # -5%
            saturazione_label = "AFFOLLATO (alta concorrenza)"
            
        else:
            # MERCATO SATURO - Rischio invenduto alto
            aggiustamento_saturazione = -0.10  # -10%
            saturazione_label = "SATURO (rischio invenduto)"
        
        fattori['aggiustamento_saturazione'] = aggiustamento_saturazione
        fattori['saturazione_label'] = saturazione_label
    
    # 3. AGGIUSTAMENTO GAP MERCATO
    # Se il mercato attuale è molto sopra OMI, c'è rischio bolla
    aggiustamento_gap = 0.0
    
    if stats_immobiliare and stats_immobiliare.get('prezzo_mq'):
        mercato_med = stats_immobiliare['prezzo_mq']['mediano']
        gap_percentuale = ((mercato_med - base_omi) / base_omi) * 100
        
        if gap_percentuale > 50:
            # GAP MOLTO ALTO - Possibile bolla, riduci ambizioni
            aggiustamento_gap = -0.08  # -8% (sii conservativo)
            gap_label = "MOLTO ALTO (+50%) - Rischio bolla"
            
        elif gap_percentuale > 35:
            # GAP ALTO - Mercato caldo ma sostenibile
            aggiustamento_gap = -0.03  # -3% (prudenza)
            gap_label = "ALTO (+35-50%) - Mercato caldo"
            
        elif gap_percentuale > 20:
            # GAP NORMALE - Premium standard nuove costruzioni
            aggiustamento_gap = 0.00  # Neutro
            gap_label = "NORMALE (+20-35%) - Premium sano"
            
        elif gap_percentuale > 0:
            # GAP BASSO - Mercato sottovalutato, opportunità
            aggiustamento_gap = 0.05  # +5% (puoi osare)
            gap_label = "BASSO (0-20%) - Mercato sottovalutato"
            
        else:
            # NEGATIVO - Mercato sotto OMI (anomalia o zona in declino)
            aggiustamento_gap = -0.05  # -5% (cautela)
            gap_label = "NEGATIVO - Mercato debole"
        
        fattori['aggiustamento_gap'] = aggiustamento_gap
        fattori['gap_label'] = gap_label
        fattori['gap_percentuale'] = gap_percentuale
    
    # 4. FATTORE RISCHIO (concentrazione agenzie)
    # Se poche agenzie dominano, c'è controllo prezzi
    fattore_rischio = 0.0
    
    if stats_immobiliare and stats_immobiliare.get('dataframe') is not None:
        import pandas as pd
        df = stats_immobiliare['dataframe']
        
        if len(df) > 0:
            agenzie_stats = df.groupby('agenzia').size()
            n_appartamenti = len(df)
            
            # Calcola concentrazione Top 3
            top3_count = agenzie_stats.nlargest(3).sum()
            concentrazione = (top3_count / n_appartamenti * 100) if n_appartamenti > 0 else 0
            
            if concentrazione > 75:
                # ALTA CONCENTRAZIONE - Mercato controllato, prezzi meno flessibili
                fattore_rischio = -0.03  # -3%
                concentrazione_label = "ALTA (>75%) - Mercato controllato"
            elif concentrazione > 50:
                # MEDIA CONCENTRAZIONE - Mix operatori
                fattore_rischio = 0.00  # Neutro
                concentrazione_label = "MEDIA (50-75%) - Mix operatori"
            else:
                # BASSA CONCENTRAZIONE - Mercato frammentato, più flessibilità
                fattore_rischio = 0.02  # +2%
                concentrazione_label = "BASSA (<50%) - Mercato frammentato"
            
            fattori['fattore_rischio'] = fattore_rischio
            fattori['concentrazione'] = concentrazione
            fattori['concentrazione_label'] = concentrazione_label
    
    # CALCOLO FINALE
    fattore_totale = (
        1.0 +
        fattori['premium_nuova_costruzione'] +
        fattori['aggiustamento_saturazione'] +
        fattori['aggiustamento_gap'] +
        fattori['fattore_rischio']
    )
    
    prezzo_ottimale = base_omi * fattore_totale
    
    # RANGE CONSIGLIATO (±8% dal prezzo ottimale)
    prezzo_minimo = prezzo_ottimale * 0.92  # -8%
    prezzo_massimo = prezzo_ottimale * 1.08  # +8%
    
    # LOGICA DI CALCOLO
    logica_calcolo = f"""
CALCOLO PRICING BENCHMARK:

Base OMI: €{base_omi:,.0f}/m²

Aggiustamenti applicati:
• Premium nuova costruzione: +{fattori['premium_nuova_costruzione']*100:.1f}%
• Saturazione mercato: {fattori['aggiustamento_saturazione']*100:+.1f}% ({fattori.get('saturazione_label', 'N/D')})
• Gap vs mercato: {fattori['aggiustamento_gap']*100:+.1f}% ({fattori.get('gap_label', 'N/D')})
• Fattore rischio: {fattori['fattore_rischio']*100:+.1f}% ({fattori.get('concentrazione_label', 'N/D')})

Fattore totale: {(fattore_totale-1)*100:+.1f}%

RANGE CONSIGLIATO:
• Minimo (entry conservative): €{prezzo_minimo:,.0f}/m²
• Ottimale (sweet spot): €{prezzo_ottimale:,.0f}/m²
• Massimo (ceiling realistico): €{prezzo_massimo:,.0f}/m²
"""
    
    return {
        'prezzo_minimo': round(prezzo_minimo, 0),
        'prezzo_ottimale': round(prezzo_ottimale, 0),
        'prezzo_massimo': round(prezzo_massimo, 0),
        'logica_calcolo': logica_calcolo.strip(),
        'fattori': fattori
    }


if __name__ == "__main__":
    # TEST
    print("🧮 Test Pricing Calculator")
    print("="*60)
    
    # Dati test
    zona_omi_test = {
        'val_med_mq': 2000
    }
    
    stats_test = {
        'n_appartamenti': 15,
        'prezzo_mq': {
            'mediano': 2400
        }
    }
    
    result = calcola_pricing_intelligente(zona_omi_test, stats_test)
    
    print(f"Minimo: €{result['prezzo_minimo']:,.0f}/m²")
    print(f"Ottimale: €{result['prezzo_ottimale']:,.0f}/m²")
    print(f"Massimo: €{result['prezzo_massimo']:,.0f}/m²")
    print(f"\n{result['logica_calcolo']}")