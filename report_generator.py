"""
Generatore report Word combinato:
- Dati OMI (Agenzia Entrate)
- Dati Immobiliare.it
- Confronto e analisi
"""

import os
import re
from datetime import datetime
from typing import Dict, Optional
import pandas as pd

from docx import Document
from docx.shared import Pt, RGBColor, Inches
from docx.enum.text import WD_ALIGN_PARAGRAPH




def aggiungi_testo_formattato(paragraph, text):
    """
    Aggiunge testo a un paragrafo gestendo la formattazione Markdown (bold).
    """
    # Pattern per **bold**
    parts = re.split(r'(\*\*.*?\*\*)', text)
    
    for part in parts:
        if part.startswith('**') and part.endswith('**'):
            # Testo bold
            bold_text = part[2:-2]
            run = paragraph.add_run(bold_text)
            run.bold = True
        else:
            # Testo normale
            paragraph.add_run(part)


def converti_markdown_a_docx(doc, markdown_text):
    """
    Converte testo Markdown in paragrafi Word formattati.
    
    Gestisce:
    - Headers (##, ###, ####)
    - Bold (**text**)
    - Liste puntate (-, *, â€¢)
    - Paragrafi normali
    """
    lines = markdown_text.split('\n')
    
    i = 0
    while i < len(lines):
        line = lines[i].strip()
        
        # Salta righe vuote
        if not line:
            i += 1
            continue
        
        # Headers ## (livello 1)
        if line.startswith('## '):
            text = line.replace('## ', '').strip()
            heading = doc.add_heading(text, level=1)
            heading_run = heading.runs[0]
            heading_run.font.color.rgb = RGBColor(0, 51, 102)
            i += 1
            continue
        
        # Headers ### (livello 2)
        if line.startswith('### '):
            text = line.replace('### ', '').strip()
            heading = doc.add_heading(text, level=2)
            heading_run = heading.runs[0]
            heading_run.font.color.rgb = RGBColor(68, 114, 196)
            i += 1
            continue
        
        # Headers #### (livello 3)
        if line.startswith('#### '):
            text = line.replace('#### ', '').strip()
            heading = doc.add_heading(text, level=3)
            i += 1
            continue
        
        # Liste puntate
        if line.startswith('- ') or line.startswith('* ') or line.startswith('â€¢ '):
            text = line.lstrip('-*â€¢ ').strip()
            p = doc.add_paragraph(style='List Bullet')
            aggiungi_testo_formattato(p, text)
            i += 1
            continue
        
        # Paragrafi normali
        p = doc.add_paragraph()
        aggiungi_testo_formattato(p, line)
        i += 1


def genera_report_combinato(
    comune: str,
    via: str,
    lat: float,
    lon: float,
    raggio_km: float,
    zona_omi: Optional[Dict],
    stats_immobiliare: Optional[Dict],
    analisi_ai: Optional[Dict] = None,
    output_dir: str = "reports"
) -> str:
    """
    Genera report Word combinato con dati OMI + Immobiliare.it
    
    Args:
        comune: Nome comune
        via: Nome via
        lat: Latitudine
        lon: Longitudine
        raggio_km: Raggio ricerca
        zona_omi: Dict con dati OMI (da omi_utils.OMIQuotazione)
        stats_immobiliare: Dict con statistiche Immobiliare.it
        output_dir: Directory output
    
    Returns:
        Path del file generato
    """
    
    print("\n📝 Generazione report Word combinato...")
    
    # Crea documento
    doc = Document()
    
    # Stile base
    style = doc.styles['Normal']
    font = style.font
    font.name = 'Calibri'
    font.size = Pt(11)
    
    now = datetime.now()
    
    # ========================================
    # TITOLO
    # ========================================
    title = doc.add_heading('REPORT ANALISI IMMOBILIARE', 0)
    title.alignment = WD_ALIGN_PARAGRAPH.CENTER
    title_run = title.runs[0]
    title_run.font.color.rgb = RGBColor(0, 51, 102)
    
    subtitle = doc.add_heading('Dati OMI + Mercato Immobiliare.it', 2)
    subtitle.alignment = WD_ALIGN_PARAGRAPH.CENTER
    subtitle_run = subtitle.runs[0]
    subtitle_run.font.color.rgb = RGBColor(100, 100, 100)
    
    doc.add_paragraph()
    
    # ========================================
    # SEZIONE 1: INFORMAZIONI RICERCA
    # ========================================
    doc.add_heading('📋 Informazioni Ricerca', 1)
    
    info_table = doc.add_table(rows=5, cols=2)
    info_table.style = 'Light Grid Accent 1'
    
    info_data = [
        ['Data estrazione', now.strftime('%d/%m/%Y %H:%M')],
        ['Località', f'{via}, {comune}'],
        ['Coordinate', f'lat: {lat:.6f}, lon: {lon:.6f}'],
        ['Raggio di ricerca', f'{raggio_km} km'],
        ['Fonte dati', 'OMI (Agenzia Entrate) + Immobiliare.it']
    ]
    
    for i, (label, value) in enumerate(info_data):
        info_table.rows[i].cells[0].text = label
        info_table.rows[i].cells[1].text = value
        info_table.rows[i].cells[0].paragraphs[0].runs[0].font.bold = True
    
    doc.add_paragraph()
    
    
    # ========================================
    # ANALISI AI (se disponibile)
    # ========================================
    if analisi_ai and analisi_ai.get('success'):
        doc.add_page_break()
        
        # Titolo sezione
        heading = doc.add_heading('🤖 ANALISI AI - EXECUTIVE REPORT', 1)
        heading_run = heading.runs[0]
        heading_run.font.color.rgb = RGBColor(0, 102, 204)
        
        p = doc.add_paragraph()
        run = p.add_run('Analisi professionale generata da Claude (Anthropic)')
        run.font.size = Pt(10)
        run.font.italic = True
        run.font.color.rgb = RGBColor(100, 100, 100)
        
        doc.add_paragraph()
        
        # Gap Analysis (se disponibile)
        if analisi_ai.get('gap_analysis'):
            gap = analisi_ai['gap_analysis']
            
            doc.add_heading('Gap Analysis: OMI vs Mercato', 2)
            
            gap_table = doc.add_table(rows=4, cols=2)
            gap_table.style = 'Medium Shading 1 Accent 1'
            
            gap_cells = [
                ('OMI Mediano', f"€{gap['omi_mediano']:,.0f}/m²"),
                ('Mercato Mediano', f"€{gap['mercato_mediano']:,.0f}/m²"),
                ('Gap Assoluto', f"€{gap['gap_assoluto']:,.0f}/m²"),
                ('Gap Percentuale', f"{gap['gap_percentuale']:+.1f}%")
            ]
            
            for i, (label, value) in enumerate(gap_cells):
                gap_table.rows[i].cells[0].text = label
                gap_table.rows[i].cells[0].paragraphs[0].runs[0].font.bold = True
                gap_table.rows[i].cells[1].text = value
                
                # Colora il gap percentuale
                if i == 3:  # Gap percentuale
                    value_run = gap_table.rows[i].cells[1].paragraphs[0].runs[0]
                    if gap['gap_percentuale'] > 10:
                        value_run.font.color.rgb = RGBColor(192, 0, 0)  # Rosso
                        value_run.font.bold = True
                    elif gap['gap_percentuale'] < -10:
                        value_run.font.color.rgb = RGBColor(0, 176, 80)  # Verde
                        value_run.font.bold = True
            
            doc.add_paragraph()
        
        # Analisi completa in Markdown
        if analisi_ai.get('analisi_completa'):
            converti_markdown_a_docx(doc, analisi_ai['analisi_completa'])
        
        doc.add_paragraph()
        
        # Note finali
        p = doc.add_paragraph()
        run = p.add_run('Nota: ')
        run.font.bold = True
        run.font.size = Pt(9)
        run2 = p.add_run('Questa analisi è stata generata automaticamente da intelligenza artificiale e deve essere considerata come supporto decisionale. Si raccomanda di validare le conclusioni con analisi approfondite e consulenza professionale.')
        run2.font.size = Pt(9)
        run2.font.italic = True
        run2.font.color.rgb = RGBColor(100, 100, 100)
    
    
    # ========================================
    # SEZIONE 2: DATI OMI
    # ========================================
    doc.add_heading('🏛️ Dati OMI (Agenzia delle Entrate)', 1)
    
    if zona_omi and zona_omi.get('val_med_mq') is not None:
        doc.add_paragraph(f"Zona OMI identificata: {zona_omi['zona_codice']}")
        doc.add_paragraph(f"Descrizione: {zona_omi['zona_descrizione']}")
        doc.add_paragraph(f"Comune: {zona_omi['comune']} ({zona_omi['provincia']})")
        doc.add_paragraph()
        
        # Valori OMI
        doc.add_paragraph('Valori €/mq (dati ufficiali rogiti):', style='Heading 3')
        
        omi_table = doc.add_table(rows=2, cols=3)
        omi_table.style = 'Light Grid Accent 1'
        
        # Header
        headers = ['Minimo', 'Mediano', 'Massimo']
        for i, header in enumerate(headers):
            cell = omi_table.rows[0].cells[i]
            cell.text = header
            cell.paragraphs[0].runs[0].font.bold = True
            cell.paragraphs[0].alignment = WD_ALIGN_PARAGRAPH.CENTER
        
        # Valori
        omi_table.rows[1].cells[0].text = f"€{zona_omi['val_min_mq']:,.0f}".replace(',', '.')
        omi_table.rows[1].cells[1].text = f"€{zona_omi['val_med_mq']:,.0f}".replace(',', '.')
        omi_table.rows[1].cells[2].text = f"€{zona_omi['val_max_mq']:,.0f}".replace(',', '.')
        
        for i in range(3):
            omi_table.rows[1].cells[i].paragraphs[0].alignment = WD_ALIGN_PARAGRAPH.CENTER
    else:
        doc.add_paragraph('⚠️ Dati OMI non disponibili per questa zona.')
    
    doc.add_paragraph()
    
    # ========================================
    # SEZIONE 3: DATI IMMOBILIARE.IT
    # ========================================
    doc.add_heading('🏢 Analisi Mercato (Immobiliare.it)', 1)
    
    if stats_immobiliare and stats_immobiliare['n_appartamenti'] > 0:
        doc.add_paragraph(f"Appartamenti analizzati: {stats_immobiliare['n_appartamenti']}")
        doc.add_paragraph()
        
        # Statistiche Prezzi
        doc.add_paragraph('Statistiche Prezzi:', style='Heading 3')
        doc.add_paragraph(f"Prezzo medio: €{stats_immobiliare['prezzo']['medio']:,.0f}".replace(',', '.'))
        doc.add_paragraph(f"Prezzo mediano: €{stats_immobiliare['prezzo']['mediano']:,.0f}".replace(',', '.'))
        doc.add_paragraph(f"Prezzo minimo: €{stats_immobiliare['prezzo']['min']:,.0f}".replace(',', '.'))
        doc.add_paragraph(f"Prezzo massimo: €{stats_immobiliare['prezzo']['max']:,.0f}".replace(',', '.'))
        doc.add_paragraph()
        
        # Statistiche Superfici
        doc.add_paragraph('Statistiche Superfici:', style='Heading 3')
        doc.add_paragraph(f"Superficie media: {stats_immobiliare['mq']['medio']:.0f} m²")
        doc.add_paragraph(f"Superficie mediana: {stats_immobiliare['mq']['mediano']:.0f} m²")
        doc.add_paragraph(f"Superficie minima: {stats_immobiliare['mq']['min']:.0f} m²")
        doc.add_paragraph(f"Superficie massima: {stats_immobiliare['mq']['max']:.0f} m²")
        doc.add_paragraph()
        
        # Prezzo al MQ
        doc.add_paragraph('Prezzo al Metro Quadro:', style='Heading 3')
        doc.add_paragraph(f"Prezzo/mq medio: €{stats_immobiliare['prezzo_mq']['medio']:,.0f}/m²".replace(',', '.'))
        doc.add_paragraph(f"Prezzo/mq mediano: €{stats_immobiliare['prezzo_mq']['mediano']:,.0f}/m²".replace(',', '.'))
        doc.add_paragraph(f"Prezzo/mq minimo: €{stats_immobiliare['prezzo_mq']['min']:,.0f}/m²".replace(',', '.'))
        doc.add_paragraph(f"Prezzo/mq massimo: €{stats_immobiliare['prezzo_mq']['max']:,.0f}/m²".replace(',', '.'))
        doc.add_paragraph()
        
        # Distribuzione per fasce prezzo
        doc.add_paragraph('Distribuzione per fasce di prezzo:', style='Heading 3')
        df = stats_immobiliare['dataframe']
        
        fasce_prezzo = [
            ('Fino a €200.000', df[df['prezzo'] <= 200000].shape[0]),
            ('€200.000 - €350.000', df[(df['prezzo'] > 200000) & (df['prezzo'] <= 350000)].shape[0]),
            ('€350.000 - €500.000', df[(df['prezzo'] > 350000) & (df['prezzo'] <= 500000)].shape[0]),
            ('Oltre €500.000', df[df['prezzo'] > 500000].shape[0])
        ]
        
        for fascia, count in fasce_prezzo:
            doc.add_paragraph(f"  • {fascia}: {count} appartamenti", style='List Bullet')
        
        doc.add_paragraph()
        
        # Distribuzione per superfici
        doc.add_paragraph('Distribuzione per fasce di superficie:', style='Heading 3')
        
        fasce_mq = [
            ('Fino a 60 m²', df[df['mq'] <= 60].shape[0]),
            ('60 - 100 m²', df[(df['mq'] > 60) & (df['mq'] <= 100)].shape[0]),
            ('100 - 150 m²', df[(df['mq'] > 100) & (df['mq'] <= 150)].shape[0]),
            ('Oltre 150 m²', df[df['mq'] > 150].shape[0])
        ]
        
        for fascia, count in fasce_mq:
            doc.add_paragraph(f"  • {fascia}: {count} appartamenti", style='List Bullet')
        
        doc.add_paragraph()
        
    else:
        doc.add_paragraph('⚠️ Nessun appartamento trovato su Immobiliare.it per questa zona.')
    
    # ========================================
    # SEZIONE 4: ANALISI PER AGENZIA
    # ========================================
    if stats_immobiliare and stats_immobiliare['n_appartamenti'] > 0:
        doc.add_heading('🏢 Analisi per Agenzia', 1)
        
        # Prepara dati agenzie
        df = stats_immobiliare['dataframe']
        agenzie = df.groupby('agenzia').agg({
            'prezzo': ['count', 'mean'],
            'mq': 'mean'
        }).reset_index()
        
        agenzie.columns = ['Agenzia', 'N° Appartamenti', 'Prezzo Medio', 'MQ Medio']
        agenzie = agenzie.sort_values('N° Appartamenti', ascending=False)
        
        # Tabella agenzie
        table = doc.add_table(rows=len(agenzie)+1, cols=4)
        table.style = 'Light Grid Accent 1'
        
        # Header
        headers = ['Agenzia', 'N° Appartamenti', 'Prezzo Medio', 'MQ Medio']
        for i, header in enumerate(headers):
            cell = table.rows[0].cells[i]
            cell.text = header
            cell.paragraphs[0].runs[0].font.bold = True
            cell.paragraphs[0].alignment = WD_ALIGN_PARAGRAPH.CENTER
        
        # Dati
        for idx, row in agenzie.iterrows():
            table.rows[idx+1].cells[0].text = str(row['Agenzia'])
            table.rows[idx+1].cells[1].text = str(int(row['N° Appartamenti']))
            table.rows[idx+1].cells[2].text = f"€{row['Prezzo Medio']:,.0f}".replace(',', '.')
            table.rows[idx+1].cells[3].text = f"{row['MQ Medio']:.0f} m²"
            
            table.rows[idx+1].cells[1].paragraphs[0].alignment = WD_ALIGN_PARAGRAPH.CENTER
            table.rows[idx+1].cells[2].paragraphs[0].alignment = WD_ALIGN_PARAGRAPH.RIGHT
            table.rows[idx+1].cells[3].paragraphs[0].alignment = WD_ALIGN_PARAGRAPH.RIGHT
        
        doc.add_paragraph()
    
    # ========================================
    # SEZIONE 5: CONFRONTO OMI vs MERCATO
    # ========================================
    doc.add_heading('📊 Confronto OMI vs Mercato', 1)
    
    if zona_omi and zona_omi.get('val_med_mq') and stats_immobiliare:
        omi_med = zona_omi['val_med_mq']
        mercato_med = stats_immobiliare['prezzo_mq']['mediano']
        
        gap = ((mercato_med - omi_med) / omi_med) * 100
        
        doc.add_paragraph(f"Valore OMI mediano: €{omi_med:,.0f}/m²".replace(',', '.'))
        doc.add_paragraph(f"Prezzo mercato mediano: €{mercato_med:,.0f}/m²".replace(',', '.'))
        doc.add_paragraph(f"Gap: {gap:+.1f}%".replace('.', ','))
        doc.add_paragraph()
        
        # Interpretazione
        doc.add_paragraph('Interpretazione:', style='Heading 3')
        if gap > 15:
            doc.add_paragraph('Il mercato quota significativamente sopra i valori OMI. Possibile zona ad alta domanda o prezzi di offerta ottimistici.')
        elif gap > 5:
            doc.add_paragraph('Il mercato quota moderatamente sopra i valori OMI. Situazione normale per offerte iniziali.')
        elif gap > -5:
            doc.add_paragraph('Il mercato è allineato ai valori OMI. Prezzi coerenti con le transazioni effettive.')
        else:
            doc.add_paragraph('Il mercato quota sotto i valori OMI. Possibili opportunità o necessità di rilancio del settore.')
        
        doc.add_paragraph()
    
    # ========================================
    # FOOTER
    # ========================================
    doc.add_page_break()
    footer_para = doc.add_paragraph()
    footer_para.add_run(f'Report generato il {now.strftime("%d/%m/%Y alle %H:%M")}').italic = True
    footer_para.alignment = WD_ALIGN_PARAGRAPH.CENTER
    footer_para.add_run('\nPlanet AI - Analisi Immobiliare').italic = True
    
    # Salva
    timestamp = now.strftime("%Y%m%d_%H%M%S")
    filename = f"report_combinato_{comune}_{timestamp}.docx"
    filepath = os.path.join(output_dir, filename)
    
    doc.save(filepath)
    
    print(f"   ✅ Report salvato: {filename}\n")
    
    return filepath