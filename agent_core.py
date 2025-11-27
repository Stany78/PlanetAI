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
    """
    full_address = f"{indirizzo}, {comune}, Italia"
    if DEBUG_MODE:
        print(f"[GEO] Geocoding: {full_address}")

    try:
        loc = _geolocator.geocode(full_address, timeout=15)
        if loc is None:
            print(f"[GEO][WARN] Via non trovata, uso centro {comune}")
            return (FALLBACK_COORDINATE[0], FALLBACK_COORDINATE[1], True, 
                   f"⚠️ Via non trovata. Usando coordinate centro {comune}")
        return (loc.latitude, loc.longitude, True, "✅ Indirizzo trovato")
    except Exception as e:
        print(f"[GEO][ERROR] Geocoding errore: {e}")
        # Usa fallback
        return (FALLBACK_COORDINATE[0], FALLBACK_COORDINATE[1], True,
               f"⚠️ Nominatim non disponibile. Usando coordinate centro Como")