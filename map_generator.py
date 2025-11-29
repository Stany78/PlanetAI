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
    
    # PIN APPARTAMENTI - RAGGRUPPATI PER EDIFICIO
    print(f"[MAP] Totale appartamenti ricevuti: {len(appartamenti) if appartamenti else 0}")
    
    if appartamenti:
        from collections import defaultdict
        
        # Raggruppa appartamenti per coordinate (stesso edificio)
        edifici = defaultdict(list)
        for app in appartamenti:
            lat = app.get('latitudine')
            lon = app.get('longitudine')
            if lat and lon:
                # Arrotonda a 5 decimali (~1 metro) per raggruppare stesso edificio
                key = (round(lat, 5), round(lon, 5))
                edifici[key].append(app)
        
        print(f"[MAP] {len(edifici)} edifici diversi con {len(appartamenti)} appartamenti totali")
        
        # PIN EDIFICI (raggruppati)
        appartamenti_con_coord = 0
        for (lat_edificio, lon_edificio), apps_edificio in edifici.items():
            n_apps = len(apps_edificio)
            appartamenti_con_coord += n_apps
            
            # Calcola prezzo medio dell'edificio
            prezzi = [a['prezzo'] for a in apps_edificio]
            prezzo_medio_edificio = sum(prezzi) / len(prezzi)
            
            # Calcola prezzo/mq medio per determinare colore
            if stats_immobiliare:
                prezzi_mq = [a['prezzo']/a['mq'] for a in apps_edificio if a.get('mq', 0) > 0]
                prezzo_mq_medio = sum(prezzi_mq) / len(prezzi_mq) if prezzi_mq else 0
                color = get_color_by_price(prezzo_mq_medio, stats_immobiliare)
            else:
                color = 'blue'
            
            # HTML popup con LISTA appartamenti
            popup_html = f"""
            <div style="width:280px; max-height:400px; overflow-y:auto">
                <b>🏢 {n_apps} Appartament{'o' if n_apps == 1 else 'i'}</b>
                <hr style="margin:5px 0">
            """
            
            for idx, app in enumerate(apps_edificio, 1):
                agenzia = app.get('agenzia', 'N/D')
                prezzo = app.get('prezzo', 0)
                mq = app.get('mq', 0)
                prezzo_mq = int(prezzo / mq) if mq > 0 else 0
                
                popup_html += f"""
                <div style="border-bottom:1px solid #eee; padding:5px 0; font-size:11px">
                    <b>Unità #{idx}</b><br>
                    💰 <b>€{prezzo:,}</b><br>
                    📐 {mq} m² · €{prezzo_mq:,}/m²<br>
                    🏢 {agenzia[:30]}
                </div>
                """
            
            popup_html += "</div>"
            
            # Determina colore CSS per il pin
            if color == 'green':
                bg_color = '#22c55e'
            elif color == 'blue':
                bg_color = '#3b82f6'
            elif color == 'orange':
                bg_color = '#f97316'
            else:  # red
                bg_color = '#ef4444'
            
            # Icona con NUMERO di appartamenti
            folium.Marker(
                location=[lat_edificio, lon_edificio],
                popup=folium.Popup(popup_html, max_width=300),
                tooltip=f"{n_apps} app. - €{int(prezzo_medio_edificio):,} - [{lat_edificio:.5f}, {lon_edificio:.5f}]",
                icon=folium.DivIcon(html=f"""
                    <div style="
                        background-color: {bg_color};
                        border: 2px solid white;
                        border-radius: 50%;
                        width: 34px;
                        height: 34px;
                        display: flex;
                        align-items: center;
                        justify-content: center;
                        color: white;
                        font-weight: bold;
                        font-size: 14px;
                        box-shadow: 0 2px 5px rgba(0,0,0,0.4);
                    ">{n_apps}</div>
                """)
            ).add_to(mappa)
        
        print(f"[MAP] Appartamenti con coordinate aggiunti alla mappa: {appartamenti_con_coord}/{len(appartamenti)}")
    
    # LEGENDA
    legend_html = '''
    <div style="position: fixed; 
                bottom: 50px; left: 50px; width: 250px; height: auto; 
                background-color: white; z-index:9999; font-size:14px;
                border:2px solid grey; border-radius: 5px; padding: 10px">
        <b>📊 Legenda Mappa</b><br>
        <hr style="margin:3px 0">
        <b>Colori (prezzo/m² medio):</b><br>
        <span style="color:green">●</span> Economico (&lt;-15% mediano)<br>
        <span style="color:blue">●</span> Medio (±15% mediano)<br>
        <span style="color:orange">●</span> Alto (+15-35% mediano)<br>
        <span style="color:red">●</span> Molto alto (&gt;+35% mediano)<br>
        <hr style="margin:3px 0">
        <b>Numero sul pin:</b> appartamenti nell'edificio<br>
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