"""
Agent Core - Geocoding con Google Geocoding API
================================================
Converte 'Via, Comune' in coordinate GPS precise.
Richiede GOOGLE_GEOCODING_API_KEY nei Secrets di Streamlit.
"""

import os
import requests


def get_google_api_key():
    """
    Recupera la API key di Google Geocoding.
    Cerca in: variabile d'ambiente, Streamlit secrets.
    """
    # 1. Variabile d'ambiente
    api_key = os.getenv('GOOGLE_GEOCODING_API_KEY')
    if api_key:
        return api_key

    # 2. Streamlit secrets
    try:
        import streamlit as st
        if 'GOOGLE_GEOCODING_API_KEY' in st.secrets:
            return st.secrets['GOOGLE_GEOCODING_API_KEY']
    except Exception:
        pass

    return None


def geocode_indirizzo(comune: str, indirizzo: str):
    """
    Geocoda 'indirizzo, comune, Italia' usando Google Geocoding API.

    Returns:
        tuple: (lat, lon, geo_info)
        geo_info dict:
        - 'success': bool
        - 'type': 'street' / 'city' / 'failed'
        - 'display_name': indirizzo completo trovato
        - 'message': messaggio per l'utente
    """
    # Validazione input
    if not comune or not comune.strip():
        return (0, 0, {
            'success': False,
            'type': 'failed',
            'display_name': '',
            'message': '❌ Comune non specificato'
        })

    if not indirizzo or not indirizzo.strip():
        return (0, 0, {
            'success': False,
            'type': 'failed',
            'display_name': '',
            'message': '❌ Via/Indirizzo non specificato'
        })

    api_key = get_google_api_key()
    if not api_key:
        return (0, 0, {
            'success': False,
            'type': 'failed',
            'display_name': '',
            'message': '❌ GOOGLE_GEOCODING_API_KEY non configurata nei Secrets'
        })

    full_address = f"{indirizzo}, {comune}, Italia"

    try:
        url = "https://maps.googleapis.com/maps/api/geocode/json"
        params = {
            'address': full_address,
            'key': api_key,
            'language': 'it',
            'region': 'it',
        }

        response = requests.get(url, params=params, timeout=15)
        data = response.json()

        status = data.get('status', '')

        if status == 'ZERO_RESULTS':
            return (0, 0, {
                'success': False,
                'type': 'failed',
                'display_name': '',
                'message': f'❌ Indirizzo non trovato: "{indirizzo}, {comune}"'
            })

        if status != 'OK':
            error_msg = data.get('error_message', status)
            return (0, 0, {
                'success': False,
                'type': 'failed',
                'display_name': '',
                'message': f'❌ Errore Google Geocoding: {error_msg}'
            })

        result = data['results'][0]
        location = result['geometry']['location']
        lat = location['lat']
        lon = location['lng']
        formatted = result.get('formatted_address', full_address)

        # Verifica tipo risultato: via specifica o solo comune?
        types = result.get('types', [])
        is_street = any(t in types for t in [
            'street_address', 'route', 'premise', 'subpremise'
        ])

        # Verifica che il comune trovato corrisponda
        comune_trovato = False
        for comp in result.get('address_components', []):
            if 'locality' in comp.get('types', []) or 'administrative_area_level_3' in comp.get('types', []):
                if comune.strip().lower() in comp.get('long_name', '').lower():
                    comune_trovato = True
                    break

        if is_street:
            return (lat, lon, {
                'success': True,
                'type': 'street',
                'display_name': formatted,
                'message': f'✅ Indirizzo trovato: {formatted}'
            })
        elif comune_trovato:
            # Trovato solo il comune, non la via specifica
            return (0, 0, {
                'success': False,
                'type': 'city',
                'display_name': formatted,
                'message': (
                    f'⚠️ Via non trovata con precisione. Trovato: {formatted}\n'
                    f'💡 Prova con il numero civico (es: "{indirizzo} 10")'
                )
            })
        else:
            # Risultato ambiguo
            return (0, 0, {
                'success': False,
                'type': 'failed',
                'display_name': formatted,
                'message': (
                    f'⚠️ Risultato ambiguo: {formatted}\n'
                    f'💡 Verifica ortografia di via e comune'
                )
            })

    except requests.exceptions.Timeout:
        return (0, 0, {
            'success': False,
            'type': 'failed',
            'display_name': '',
            'message': '❌ Timeout connessione a Google. Riprova tra qualche secondo.'
        })
    except Exception as e:
        return (0, 0, {
            'success': False,
            'type': 'failed',
            'display_name': '',
            'message': f'❌ Errore durante la ricerca: {str(e)}'
        })