"""
Agent Core - Funzioni essenziali per Planet AI
"""

from geopy.geocoders import Nominatim

# Fallback coordinate (Como centro)
FALLBACK_COORDINATE = (45.8081, 9.0852)
DEBUG_MODE = False

# Geocoder
_geolocator = Nominatim(user_agent="planet_ai_omi_agent")


def geocode_indirizzo(comune: str, indirizzo: str) -> tuple[float, float, bool]:
    """
    Geocoda 'indirizzo, comune, Italia' usando Nominatim.
    
    Returns:
        tuple: (lat, lon, success)
        - success è sempre True (usa fallback se necessario)
    """
    full_address = f"{indirizzo}, {comune}, Italia"
    if DEBUG_MODE:
        print(f"[GEO] Geocoding: {full_address}")

    try:
        loc = _geolocator.geocode(full_address, timeout=15)
        if loc is None:
            print(f"[GEO][WARN] Via non trovata, uso centro {comune}")
            return (FALLBACK_COORDINATE[0], FALLBACK_COORDINATE[1], True)
        return (loc.latitude, loc.longitude, True)
    except Exception as e:
        print(f"[GEO][ERROR] Geocoding errore: {e}. Uso fallback.")
        return (FALLBACK_COORDINATE[0], FALLBACK_COORDINATE[1], True)