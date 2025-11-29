"""
Modulo per scraping dati da Immobiliare.it
Estrae: Prezzo, MQ, Agenzia, Coordinate GPS per nuove costruzioni
"""

import requests
import time
from typing import List, Dict, Optional


def cerca_appartamenti(lat: float, lon: float, raggio_km: float, max_pagine: int = 5) -> List[Dict]:
    """
    Chiama API Immobiliare.it e estrae appartamenti nuove costruzioni
    
    Args:
        lat: Latitudine centro ricerca
        lon: Longitudine centro ricerca
        raggio_km: Raggio di ricerca in km
        max_pagine: Numero massimo di pagine da scaricare
    
    Returns:
        Lista di dict con: progetto_id, prezzo, mq, agenzia, latitudine, longitudine
    """
    import math
    
    base_url = "https://www.immobiliare.it/api-next/search-list/listings/"
    
    headers = {
        'User-Agent': 'Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36',
        'Accept': 'application/json',
        'Referer': 'https://www.immobiliare.it/search-list/',
    }
    
    appartamenti_totali = []
    pagina = 1
    
    print(f"🔍 Inizio scraping Immobiliare.it (raggio {raggio_km} km)...")
    
    while pagina <= max_pagine:
        # Calcola bounding box corretto in base al raggio
        # 1 grado di latitudine ≈ 111 km
        # 1 grado di longitudine ≈ 111 km * cos(latitudine)
        delta_lat = raggio_km / 111.0
        delta_lon = raggio_km / (111.0 * math.cos(math.radians(lat)))
        
        min_lat = lat - delta_lat
        max_lat = lat + delta_lat
        min_lon = lon - delta_lon
        max_lon = lon + delta_lon
        
        # Parametri per questa pagina
        params = {
            'raggio': str(int(raggio_km * 1000)),
            'centro': f'{lat},{lon}',
            'idContratto': '1',
            'idCategoria': '6',
            'idTipologia[0]': '54',  # Appartamenti
            'idTipologia[1]': '85',  # Attici e Mansarde
            '__lang': 'it',
            'minLat': f'{min_lat:.6f}',
            'maxLat': f'{max_lat:.6f}',
            'minLng': f'{min_lon:.6f}',
            'maxLng': f'{max_lon:.6f}',
            'pag': str(pagina),
            'paramsCount': '7',
            'path': '/search-list/',
        }
        
        url = base_url + '?' + '&'.join([f'{k}={v}' for k, v in params.items()])
        
        print(f"📄 Pagina {pagina}...", end=" ")
        
        try:
            response = requests.get(url, headers=headers, timeout=15)
            
            if response.status_code != 200:
                print(f"❌ Errore HTTP {response.status_code}")
                break
            
            data = response.json()
            results = data.get('results', [])
            
            print(f"✓ {len(results)} annunci")
            
            # Info prima pagina
            if pagina == 1:
                total_ads = data.get('totalAds', 0)
                max_pages = data.get('maxPages', 1)
                print(f"   ℹ️  totalAds: {total_ads}, maxPages: {max_pages}")
            
            if len(results) == 0:
                print(f"   ⚠️  Pagina vuota - stop")
                break
            
            # Estrai dati da questa pagina
            
            # DEBUG: Stampa PRIMO result COMPLETO
            if pagina == 1 and len(results) > 0:
                import json
                json_str = json.dumps(results[0], indent=2)
                print("[SCRAPER][DEBUG] === INIZIO JSON COMPLETO ===")
                # Stampa a pezzi per evitare troncamenti
                for i in range(0, len(json_str), 3000):
                    print(json_str[i:i+3000])
                print("[SCRAPER][DEBUG] === FINE JSON COMPLETO ===")
            
            for idx, result in enumerate(results):
                real_estate = result.get('realEstate', {})
                
                # ID del progetto (per raggruppare appartamenti dello stesso annuncio)
                progetto_id = real_estate.get('id', 'N/D')
                
                # COORDINATE GPS - CERCA IN TUTTI I POSTI POSSIBILI
                latitudine = None
                longitudine = None
                
                # Opzione 1: location.latitude/longitude
                location = real_estate.get('location', {})
                if location:
                    latitudine = location.get('latitude') or location.get('lat')
                    longitudine = location.get('longitude') or location.get('lng') or location.get('lon')
                
                # Opzione 2: properties[0].location
                if not latitudine and 'properties' in real_estate:
                    props = real_estate.get('properties', [])
                    if len(props) > 0:
                        prop_loc = props[0].get('location', {})
                        latitudine = prop_loc.get('latitude') or prop_loc.get('lat')
                        longitudine = prop_loc.get('longitude') or prop_loc.get('lng')
                
                # Opzione 3: Direttamente in real_estate
                if not latitudine:
                    latitudine = real_estate.get('latitude') or real_estate.get('lat')
                    longitudine = real_estate.get('longitude') or real_estate.get('lng') or real_estate.get('lon')
                
                # Opzione 4: geometry
                if not latitudine and 'geometry' in real_estate:
                    geom = real_estate.get('geometry', {})
                    if 'coordinates' in geom:
                        coords = geom.get('coordinates', [])
                        if len(coords) >= 2:
                            longitudine = coords[0]  # GeoJSON è lon, lat
                            latitudine = coords[1]
                
                # DEBUG
                if idx == 0:
                    print(f"[SCRAPER][DEBUG] Coordinate trovate: lat={latitudine}, lon={longitudine}")
                
                # Agenzia
                agenzia = "N/D"
                advertiser = real_estate.get('advertiser', {})
                if advertiser:
                    agency = advertiser.get('agency', {})
                    if agency:
                        agenzia = agency.get('displayName', 'N/D')
                
                # Properties array
                properties = real_estate.get('properties', [])
                
                for prop in properties:
                    # Prezzo
                    price_obj = prop.get('price', {})
                    prezzo = price_obj.get('value')
                    
                    # MQ
                    surface = prop.get('surface', '')
                    mq = None
                    if surface:
                        mq_str = surface.replace(' m²', '').strip()
                        try:
                            mq = int(mq_str)
                        except:
                            pass
                    
                    # Salva solo se ha prezzo e mq validi
                    if prezzo and mq:
                        appartamenti_totali.append({
                            'progetto_id': progetto_id,
                            'prezzo': prezzo,
                            'mq': mq,
                            'agenzia': agenzia,
                            'latitudine': latitudine,  # NUOVO
                            'longitudine': longitudine,  # NUOVO
                        })
            
            # Pausa tra pagine
            if pagina < max_pagine:
                time.sleep(1)
            
            pagina += 1
            
        except Exception as e:
            print(f"❌ Errore pagina {pagina}: {e}")
            break
    
    print(f"\n✅ Totale appartamenti estratti (prima rimozione duplicati): {len(appartamenti_totali)}\n")
    
    return appartamenti_totali


def calcola_statistiche(appartamenti: List[Dict]) -> Dict:
    """
    Calcola statistiche sui dati estratti
    Include rimozione duplicati
    
    Returns:
        Dict con statistiche aggregate
    """
    if not appartamenti:
        return None
    
    import pandas as pd
    import numpy as np
    
    df = pd.DataFrame(appartamenti)
    
    # RIMOZIONE DUPLICATI (prezzo + mq + agenzia)
    print(f"🔄 Rimozione duplicati...")
    print(f"   Prima: {len(df)} appartamenti")
    
    df_unique = df.drop_duplicates(subset=['prezzo', 'mq', 'agenzia'], keep='first')
    
    print(f"   Dopo: {len(df_unique)} appartamenti (rimossi {len(df) - len(df_unique)} duplicati)\n")
    
    df = df_unique
    
    # Calcola prezzo/mq
    df['prezzo_mq'] = df['prezzo'] / df['mq']
    
    stats = {
        'n_appartamenti': len(df),
        'n_progetti': df['progetto_id'].nunique(),
        'prezzo': {
            'medio': df['prezzo'].mean(),
            'mediano': df['prezzo'].median(),
            'min': df['prezzo'].min(),
            'max': df['prezzo'].max(),
        },
        'mq': {
            'medio': df['mq'].mean(),
            'mediano': df['mq'].median(),
            'min': df['mq'].min(),
            'max': df['mq'].max(),
        },
        'prezzo_mq': {
            'medio': df['prezzo_mq'].mean(),
            'mediano': df['prezzo_mq'].median(),
            'min': df['prezzo_mq'].min(),
            'max': df['prezzo_mq'].max(),
        },
        'agenzie': df.groupby('agenzia').agg({
            'prezzo': ['count', 'mean'],
            'mq': 'mean',
            'progetto_id': 'nunique'
        }).reset_index().to_dict('records'),
        'dataframe': df  # Per ulteriori analisi
    }
    
    return stats