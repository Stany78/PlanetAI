"""
Agent Core - Funzioni essenziali per Planet AI
"""

from geopy.geocoders import Nominatim

# Fallback coordinate (Como centro)
FALLBACK_COORDINATE = (45.8081, 9.0852)
DEBUG_MODE = False

# Geocoder
_geolocator = Nominatim(user_agent="planet_ai_omi_agent")


def geocode_indirizzo(comune: str, indirizzo: str) -> tuple[float, float]:
    """
    Geocoda 'indirizzo, comune, Italia' usando Nominatim.
    Se fallisce, usa FALLBACK_COORDINATE.
    """
    full_address = f"{indirizzo}, {comune}, Italia"
    if DEBUG_MODE:
        print(f"[GEO] Geocoding: {full_address}")

    try:
        loc = _geolocator.geocode(full_address)
        if loc is None:
            print("[GEO][WARN] Geocoding fallito, uso FALLBACK_COORDINATE.")
            return FALLBACK_COORDINATE
        return (loc.latitude, loc.longitude)
    except Exception as e:
        print(f"[GEO][ERROR] Geocoding errore: {e}. Uso FALLBACK_COORDINATE.")
        return FALLBACK_COORDINATE