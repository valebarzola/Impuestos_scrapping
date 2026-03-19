from pathlib import Path

CACHE_DIR = Path(__file__).parent / "cache"
CACHE_FILE = CACHE_DIR / "cotizaciones.json"
CACHE_DIR.mkdir(exist_ok=True)

AFIP_URL = "https://serviciosweb.afip.gob.ar/aduana/cotizacionesMaria/formulario.asp"

MONEDAS_SOPORTADAS = {
    "DOL": "DOL - DOLAR ESTADOUNIDENSE",
    "BRL": "BRL - REAL BRASILEÑO",
    "EUR": "EUR - EURO",
    "UYU": "UYU - PESO URUGUAYO",
    "CLP": "CLP - PESO CHILENO"
}
