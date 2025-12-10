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
    # Validazione input
    if not comune or not comune.strip():
        return (0, 0, {
            'success': False,
            'message': "❌ Comune non specificato"
        })
    
    if not indirizzo or not indirizzo.strip():
        return (0, 0, {
            'success': False,
            'message': "❌ Via/Indirizzo non specificato"
        })
    
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
        error_msg = str(e)
        print(f"[GEO][ERROR] Geocoding errore: {error_msg}")
        
        # Messaggio più specifico
        if "timed out" in error_msg.lower() or "timeout" in error_msg.lower():
            user_message = "❌ Timeout connessione al servizio di geocoding. Riprova tra qualche secondo."
        elif "not found" in error_msg.lower():
            user_message = f"❌ Indirizzo non trovato: '{indirizzo}, {comune}'\n\n💡 Suggerimenti:\n• Verifica l'ortografia\n• Aggiungi numero civico (es: 'Via Anzani 10')\n• Prova una via principale vicina"
        else:
            user_message = f"❌ Errore geocoding: {error_msg}\n\nRiprova tra qualche secondo."
        
        return (0, 0, {
            'success': False,
            'message': user_message
        })
