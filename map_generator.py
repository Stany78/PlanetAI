"""
Map Generator - Generazione mappe interattive per Planet AI
============================================================
Crea mappe con Folium mostrando:
- Centro ricerca (via cercata)
- Appartamenti trovati (con colori per fascia prezzo)
- Raggio di ricerca
"""

import folium
from folium import plugins
import io
from PIL import Image
import pandas as pd
from typing import Optional, Dict, List


def get_color_by_price(prezzo_mq: float, stats: Dict) -> str:
    """
    Restituisce un colore in base alla fascia di prezzo/mq.
    
    Args:
        prezzo_mq: Prezzo al metro quadro
        stats: Statistiche con min, mediano, max
    
    Returns:
        str: Nome colore per Folium
    """
    mediano = stats['prezzo_mq']['mediano']
    
    if prezzo_mq < mediano * 0.85:
        return 'green'  # Economico
    elif prezzo_mq < mediano * 1.15:
        return 'blue'   # Medio
    elif prezzo_mq < mediano * 1.35:
        return 'orange' # Alto
    else:
        return 'red'    # Molto alto


def crea_mappa_interattiva(
    lat_centro: float,
    lon_centro: float,
    via: str,
    comune: str,
    raggio_km: float,
    appartamenti: List[Dict],
    stats_immobiliare: Optional[Dict] = None
) -> folium.Map:
    """
    Crea mappa interattiva con Folium.
    
    Args:
        lat_centro: Latitudine centro ricerca
        lon_centro: Longitudine centro ricerca
        via: Via/indirizzo cercato
        comune: Comune
        raggio_km: Raggio ricerca in km
        appartamenti: Lista appartamenti trovati
        stats_immobiliare: Statistiche per determinare colori
    
    Returns:
        folium.Map: Oggetto mappa Folium
    """
    
    # Crea mappa centrata sul punto di ricerca
    mappa = folium.Map(
        location=[lat_centro, lon_centro],
        zoom_start=14,
        tiles='OpenStreetMap'
    )
    
    # CERCHIO - Raggio di ricerca
    folium.Circle(
        location=[lat_centro, lon_centro],
        radius=raggio_km * 1000,  # Converti km in metri
        color='blue',
        fill=True,
        fillColor='lightblue',
        fillOpacity=0.2,
        popup=f'Raggio ricerca: {raggio_km} km',
        tooltip=f'Raggio ricerca: {raggio_km} km'
    ).add_to(mappa)
    
    # PIN ROSSO - Centro ricerca (via cercata)
    folium.Marker(
        location=[lat_centro, lon_centro],
        popup=f"<b>📍 Centro Ricerca</b><br>{via}, {comune}",
        tooltip=f"Centro ricerca: {via}",
        icon=folium.Icon(color='red', icon='home', prefix='fa')
    ).add_to(mappa)
    
    # PIN APPARTAMENTI
    if appartamenti:
        for idx, app in enumerate(appartamenti, 1):
            lat = app.get('latitudine')
            lon = app.get('longitudine')
            
            if lat and lon:
                prezzo = app.get('prezzo', 0)
                mq = app.get('mq', 0)
                prezzo_mq = app.get('prezzo_mq', 0)
                agenzia = app.get('agenzia', 'N/D')
                url = app.get('url', '')
                
                # Determina colore in base al prezzo
                if stats_immobiliare:
                    color = get_color_by_price(prezzo_mq, stats_immobiliare)
                else:
                    color = 'blue'
                
                # Crea popup con info dettagliate
                popup_html = f"""
                <div style="width:200px">
                    <b>🏗️ Appartamento #{idx}</b><br>
                    <hr style="margin:5px 0">
                    💰 <b>Prezzo:</b> €{prezzo:,.0f}<br>
                    📐 <b>Superficie:</b> {mq:.0f} m²<br>
                    💵 <b>€/m²:</b> €{prezzo_mq:,.0f}/m²<br>
                    🏢 <b>Agenzia:</b> {agenzia}<br>
                """
                
                if url:
                    popup_html += f'<br><a href="{url}" target="_blank">🔗 Vedi annuncio</a>'
                
                popup_html += "</div>"
                
                # Tooltip breve (hover)
                tooltip_text = f"#{idx}: €{prezzo:,.0f} - {mq:.0f}m² - {agenzia}"
                
                # Aggiungi marker
                folium.Marker(
                    location=[lat, lon],
                    popup=folium.Popup(popup_html, max_width=250),
                    tooltip=tooltip_text,
                    icon=folium.Icon(
                        color=color,
                        icon='building',
                        prefix='fa'
                    )
                ).add_to(mappa)
    
    # LEGENDA
    legend_html = '''
    <div style="position: fixed; 
                bottom: 50px; left: 50px; width: 220px; height: auto; 
                background-color: white; z-index:9999; font-size:14px;
                border:2px solid grey; border-radius: 5px; padding: 10px">
        <b>📊 Legenda Prezzi/m²</b><br>
        <i class="fa fa-map-marker" style="color:green"></i> Economico (&lt;-15% mediano)<br>
        <i class="fa fa-map-marker" style="color:blue"></i> Medio (±15% mediano)<br>
        <i class="fa fa-map-marker" style="color:orange"></i> Alto (+15-35% mediano)<br>
        <i class="fa fa-map-marker" style="color:red"></i> Molto alto (&gt;+35% mediano)<br>
        <hr style="margin:5px 0">
        <i class="fa fa-home" style="color:red"></i> Centro ricerca<br>
    </div>
    '''
    mappa.get_root().html.add_child(folium.Element(legend_html))
    
    # Aggiungi fullscreen button
    plugins.Fullscreen().add_to(mappa)
    
    return mappa


def salva_mappa_come_immagine(
    mappa: folium.Map,
    output_path: str,
    width: int = 1200,
    height: int = 800
) -> str:
    """
    Salva la mappa come immagine PNG (per report Word).
    
    Args:
        mappa: Oggetto folium.Map
        output_path: Percorso file output
        width: Larghezza immagine
        height: Altezza immagine
    
    Returns:
        str: Percorso file salvato
    """
    try:
        # Salva come HTML temporaneo
        html_path = output_path.replace('.png', '_temp.html')
        mappa.save(html_path)
        
        # Nota: Per convertire HTML in PNG serve selenium o altra libreria
        # Per ora salviamo solo l'HTML e nel report mettiamo un link
        # In alternativa si può usare selenium + chromedriver
        
        print(f"[MAP] Mappa salvata come HTML: {html_path}")
        return html_path
        
    except Exception as e:
        print(f"[MAP][ERROR] Errore salvataggio mappa: {e}")
        return None


def get_mappa_statistiche(appartamenti: List[Dict]) -> Dict:
    """
    Calcola statistiche geografiche degli appartamenti.
    
    Args:
        appartamenti: Lista appartamenti con coordinate
    
    Returns:
        Dict con statistiche geografiche
    """
    appartamenti_con_coord = [
        app for app in appartamenti 
        if app.get('latitudine') and app.get('longitudine')
    ]
    
    if not appartamenti_con_coord:
        return {
            'totale': 0,
            'con_coordinate': 0,
            'senza_coordinate': len(appartamenti)
        }
    
    return {
        'totale': len(appartamenti),
        'con_coordinate': len(appartamenti_con_coord),
        'senza_coordinate': len(appartamenti) - len(appartamenti_con_coord),
        'lat_min': min(app['latitudine'] for app in appartamenti_con_coord),
        'lat_max': max(app['latitudine'] for app in appartamenti_con_coord),
        'lon_min': min(app['longitudine'] for app in appartamenti_con_coord),
        'lon_max': max(app['longitudine'] for app in appartamenti_con_coord),
    }


if __name__ == "__main__":
    """
    Test del modulo
    """
    print("🗺️ Test Map Generator")
    print("="*50)
    
    # Dati di test
    lat_centro = 45.8081
    lon_centro = 9.0852
    
    appartamenti_test = [
        {
            'latitudine': 45.8100,
            'longitudine': 9.0870,
            'prezzo': 350000,
            'mq': 85,
            'prezzo_mq': 4117,
            'agenzia': 'Immobiliare Test 1',
            'url': 'https://www.immobiliare.it/test1'
        },
        {
            'latitudine': 45.8050,
            'longitudine': 9.0830,
            'prezzo': 280000,
            'mq': 70,
            'prezzo_mq': 4000,
            'agenzia': 'Immobiliare Test 2',
            'url': 'https://www.immobiliare.it/test2'
        }
    ]
    
    stats_test = {
        'prezzo_mq': {
            'min': 3000,
            'mediano': 4000,
            'max': 5000
        }
    }
    
    # Crea mappa
    mappa = crea_mappa_interattiva(
        lat_centro=lat_centro,
        lon_centro=lon_centro,
        via="Via Anzani",
        comune="Como",
        raggio_km=1.0,
        appartamenti=appartamenti_test,
        stats_immobiliare=stats_test
    )
    
    # Salva
    mappa.save('/tmp/test_mappa.html')
    print("✅ Mappa test creata: /tmp/test_mappa.html")
    
    # Statistiche
    stats = get_mappa_statistiche(appartamenti_test)
    print(f"📊 Statistiche: {stats}")