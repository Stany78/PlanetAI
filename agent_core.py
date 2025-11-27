"""
Agent Core - Funzioni essenziali per Planet AI
"""

from geopy.geocoders import Nominatim

DEBUG_MODE = False

# Geocoder
_geolocator = Nominatim(user_agent="planet_ai_omi_agent")


def geocode_indirizzo(comune: str, indirizzo: str) -> tuple[float, float, dict]:
    """
    Geocoda 'indirizzo, comune, Italia' usando Nominatim.
    
    Returns:
        tuple: (lat, lon, geo_info)
        geo_info = {'success': bool, 'message': str}
    """
    full_address = f"{indirizzo}, {comune}, Italia"
    if DEBUG_MODE:
        print(f"[GEO] Geocoding: {full_address}")

    try:
        loc = _geolocator.geocode(full_address, timeout=15)
        if loc is None:
            print(f"[GEO][WARN] Geocoding fallito")
            return (0, 0, {
                'success': False,
                'message': f"❌ Via non trovata: '{indirizzo}' a {comune}"
            })
        
        # TROVATO
        return (loc.latitude, loc.longitude, {
            'success': True,
            'message': f"✅ Trovato: {loc.address}"
        })
        
    except Exception as e:
        print(f"[GEO][ERROR] Geocoding errore: {e}")
        return (0, 0, {
            'success': False,
            'message': f"❌ Errore connessione: riprova tra qualche secondo"
        })