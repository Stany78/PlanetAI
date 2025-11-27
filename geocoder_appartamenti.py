"""
Geocoder Appartamenti - Aggiunge coordinate GPS agli appartamenti
===================================================================
Geocoda gli indirizzi degli appartamenti trovati su Immobiliare.it
"""

from geopy.geocoders import Nominatim
from typing import List, Dict
import time


# Geocoder
_geolocator = Nominatim(user_agent="planet_ai_geocoder")


def geocoda_appartamento(indirizzo: str, comune: str, timeout: int = 10) -> tuple[float, float]:
    """
    Geocoda un singolo indirizzo.
    
    Args:
        indirizzo: Indirizzo appartamento
        comune: Comune
        timeout: Timeout richiesta
    
    Returns:
        tuple: (lat, lon) o (None, None) se fallisce
    """
    if not indirizzo or indirizzo == "N/D":
        return (None, None)
    
    try:
        full_address = f"{indirizzo}, {comune}, Italia"
        loc = _geolocator.geocode(full_address, timeout=timeout)
        
        if loc:
            return (loc.latitude, loc.longitude)
        else:
            return (None, None)
            
    except Exception as e:
        print(f"[GEOCODER][WARN] Errore geocoding '{indirizzo}': {e}")
        return (None, None)


def geocoda_appartamenti(appartamenti: List[Dict], comune: str, delay: float = 1.0) -> List[Dict]:
    """
    Aggiunge coordinate GPS a una lista di appartamenti.
    
    Args:
        appartamenti: Lista appartamenti da geocodare
        comune: Comune di riferimento
        delay: Pausa tra richieste (per non sovraccaricare Nominatim)
    
    Returns:
        List[Dict]: Lista appartamenti con lat/lon aggiunti
    """
    if not appartamenti:
        return []
    
    print(f"\n[GEOCODER] Geocoding {len(appartamenti)} appartamenti...")
    
    # DEBUG: Mostra quali campi sono disponibili
    if len(appartamenti) > 0:
        print(f"[GEOCODER][DEBUG] Campi disponibili nel primo appartamento: {list(appartamenti[0].keys())}")
    
    geocodati = 0
    falliti = 0
    
    for i, app in enumerate(appartamenti, 1):
        # Prova TUTTI i possibili nomi campo per l'indirizzo
        indirizzo = (
            app.get('indirizzo') or 
            app.get('via') or 
            app.get('address') or 
            app.get('location') or
            app.get('localita') or
            app.get('zona') or
            app.get('title')  # A volte il titolo contiene l'indirizzo
        )
        
        if indirizzo and indirizzo != 'N/D':
            lat, lon = geocoda_appartamento(indirizzo, comune)
            
            if lat and lon:
                app['latitudine'] = lat
                app['longitudine'] = lon
                geocodati += 1
                print(f"[GEOCODER] {i}/{len(appartamenti)}: ✓ {indirizzo[:40]}...")
            else:
                app['latitudine'] = None
                app['longitudine'] = None
                falliti += 1
                print(f"[GEOCODER] {i}/{len(appartamenti)}: ✗ {indirizzo[:40]}...")
        else:
            # Nessun indirizzo disponibile
            app['latitudine'] = None
            app['longitudine'] = None
            falliti += 1
            print(f"[GEOCODER] {i}/{len(appartamenti)}: ✗ Indirizzo non disponibile")
        
        # Pausa per non sovraccaricare Nominatim
        if i < len(appartamenti):
            time.sleep(delay)
    
    print(f"[GEOCODER] Completato: {geocodati} geocodati, {falliti} falliti")
    
    return appartamenti


def filtra_appartamenti_con_coordinate(appartamenti: List[Dict]) -> List[Dict]:
    """
    Filtra solo appartamenti con coordinate valide.
    
    Args:
        appartamenti: Lista appartamenti
    
    Returns:
        List[Dict]: Solo appartamenti con lat/lon
    """
    return [
        app for app in appartamenti
        if app.get('latitudine') is not None and app.get('longitudine') is not None
    ]


if __name__ == "__main__":
    """
    Test del modulo
    """
    print("🧪 Test Geocoder Appartamenti")
    print("="*50)
    
    # Dati test
    appartamenti_test = [
        {
            'prezzo': 350000,
            'mq': 85,
            'indirizzo': 'Via Anzani 10',
            'agenzia': 'Test Immobiliare'
        },
        {
            'prezzo': 280000,
            'mq': 70,
            'indirizzo': 'Via Borgovico 15',
            'agenzia': 'Test Agenzia'
        },
        {
            'prezzo': 400000,
            'mq': 100,
            'indirizzo': 'INDIRIZZO_INESISTENTE_XYZ',
            'agenzia': 'Test'
        }
    ]
    
    # Geocoda
    risultato = geocoda_appartamenti(appartamenti_test, "Como", delay=0.5)
    
    # Mostra risultati
    print("\n📊 Risultati:")
    for app in risultato:
        if app.get('latitudine'):
            print(f"✓ {app['indirizzo']}: {app['latitudine']:.6f}, {app['longitudine']:.6f}")
        else:
            print(f"✗ {app['indirizzo']}: Nessuna coordinata")
    
    # Filtra
    con_coord = filtra_appartamenti_con_coordinate(risultato)
    print(f"\n✅ {len(con_coord)}/{len(risultato)} appartamenti con coordinate")