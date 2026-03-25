import logging
import requests
from bs4 import BeautifulSoup
from datetime import datetime
from typing import Dict
from config import AFIP_URL

logger = logging.getLogger(__name__)


def validar_formato_fecha(fecha_str: str) -> bool:
    """Valida que la fecha tenga formato DD/MM/YYYY"""
    try:
        datetime.strptime(fecha_str, "%d/%m/%Y")
        return True
    except ValueError:
        return False


def obtener_cotizaciones_de_afip(fecha: str) -> Dict[str, Dict]:
    """
    Scrape cotizaciones de ARCA para una fecha específica.

    Args:
        fecha: formato DD/MM/YYYY

    Returns:
        Dict con estructura: {moneda: {tipo_comprador, tipo_vendedor, fecha_oficial}}

    Raises:
        ValueError: si el formato de fecha es inválido
        LookupError: si no hay cotizaciones para esa fecha
        ConnectionError: si ARCA no está disponible
    """
    if not validar_formato_fecha(fecha):
        raise ValueError(f"Formato de fecha inválido. Use DD/MM/YYYY. Recibido: {fecha}")

    dia, mes, anio = fecha.split('/')
    payload = {
        "dia": dia,
        "mes": mes,
        "anio": anio,
        "consultarConstancia.x": "57",
        "consultarConstancia.y": "27"
    }

    try:
        logger.info(f"Consultando ARCA para fecha: {fecha}")
        response = requests.post(AFIP_URL, data=payload, timeout=10)
        response.raise_for_status()

        soup = BeautifulSoup(response.content, 'html.parser')
        filas = soup.find_all('tr')
        cotizaciones = {}

        for fila in filas:
            celdas = fila.find_all('td')
            for i in range(0, len(celdas) - 3, 5):
                try:
                    codigo_moneda = celdas[i].get_text().split()[0]
                    fecha_celda = celdas[i + 1].get_text(strip=True)
                    tipo_vendedor_str = celdas[i + 2].get_text(strip=True)
                    tipo_comprador_str = celdas[i + 3].get_text(strip=True)

                    if not codigo_moneda or not fecha_celda:
                        continue

                    cotizaciones[codigo_moneda] = {
                        "tipo_comprador": float(tipo_comprador_str.replace(',', '.')),
                        "tipo_vendedor": float(tipo_vendedor_str.replace(',', '.')),
                        "fecha_oficial": fecha_celda
                    }
                except (ValueError, IndexError):
                    continue

        if not cotizaciones:
            raise LookupError(f"No hay cotizaciones disponibles para {fecha}")

        logger.info(f"ARCA: {len(cotizaciones)} monedas obtenidas para {fecha}")
        return cotizaciones

    except requests.exceptions.RequestException as e:
        raise ConnectionError(f"Servicio de ARCA no disponible: {e}")


def obtener_dolar_comprador(fecha: str) -> float:
    """
    Devuelve únicamente el tipo de cambio comprador del dólar para una fecha.

    Args:
        fecha: formato DD/MM/YYYY

    Returns:
        float con el tipo de cambio comprador del DOL

    Raises:
        LookupError: si no hay cotización del DOL para esa fecha
    """
    cotizaciones = obtener_cotizaciones_de_afip(fecha)
    if "DOL" not in cotizaciones:
        raise LookupError(f"No se encontró cotización del DOL para {fecha}")
    return cotizaciones["DOL"]["tipo_comprador"]
