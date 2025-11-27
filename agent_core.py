"""
Agent Core - Funzioni essenziali per Planet AI
"""

from geopy.geocoders import Nominatim

# Fallback coordinate (Como centro)
FALLBACK_COORDINATE = (45.8081, 9.0852)
DEBUG_MODE = False

# Geocoder
_geolocator = Nominatim(user_agent="planet_ai_omi_agent")


def geocode_indirizzo(comune: str, indirizzo: str) -> tuple[float, float, bool, str]:
    """
    Geocoda 'indirizzo, comune, Italia' usando Nominatim.
    
    Returns:
        tuple: (lat, lon, success, message)
        - success: True se trovato, False se non trovato
        - message: descrive cosa è stato trovato
    """
    full_address = f"{indirizzo}, {comune}, Italia"
    if DEBUG_MODE:
        print(f"[GEO] Geocoding: {full_address}")

    try:
        loc = _geolocator.geocode(full_address, timeout=15)
        if loc is None:
            print(f"[GEO][WARN] Geocoding fallito")
            return (0, 0, False, f"❌ Via non trovata: '{indirizzo}' a {comune}")
        
        # TROVATO - mostra l'indirizzo esatto
        return (loc.latitude, loc.longitude, True, f"✅ Trovato: {loc.address}")
        
    except Exception as e:
        print(f"[GEO][ERROR] Geocoding errore: {e}")
        return (0, 0, False, f"❌ Errore: {str(e)}")