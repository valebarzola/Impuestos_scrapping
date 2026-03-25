import logging
from datetime import datetime, timedelta
from typing import Dict, Tuple

from cache.cache_config import Cache
from scrapper.scrapping import obtener_cotizaciones_de_afip


class logica_service:

    def __init__(self):
        self.cache = Cache()
        self.logger = logging.getLogger(__name__)

    def restar_un_dia(self, fecha_str: str) -> str:
        fecha = datetime.strptime(fecha_str, "%d/%m/%Y")
        return (fecha - timedelta(days=1)).strftime("%d/%m/%Y")

    def buscar_ultima_cotizacion_disponible(
        self,
        fecha_solicitada: str,
        moneda: str,
        max_retroceso: int = 7
    ) -> Tuple[str, Dict]:
        """
        Busca cotización del día anterior a la fecha indicada.
        Si no existe, retrocede hasta max_retroceso días.

        Raises:
            LookupError: si no encuentra cotización en el rango
            ConnectionError: si ARCA no está disponible
        """
        fecha_buscada = self.restar_un_dia(fecha_solicitada)

        for _ in range(max_retroceso):
            cotizacion = self.cache.obtener_del_cache(fecha_buscada, moneda)
            if cotizacion:
                return fecha_buscada, cotizacion

            try:
                cotizaciones = obtener_cotizaciones_de_afip(fecha_buscada)
                if moneda in cotizaciones:
                    datos = cotizaciones[moneda]
                    self.cache.guardar_en_cache(fecha_buscada, moneda, datos)
                    return fecha_buscada, datos
            except LookupError:
                pass  # Fecha sin datos, seguir retrocediendo
            except (ValueError, ConnectionError):
                raise  # Errores no recuperables, propagar

            fecha_buscada = self.restar_un_dia(fecha_buscada)

        raise LookupError(
            f"No hay cotización para {moneda} en los últimos {max_retroceso} días anteriores a {fecha_solicitada}"
        )
