import json
import logging
from typing import Optional, Dict
from config import CACHE_FILE


class Cache:

    def __init__(self):
        self.logger = logging.getLogger(__name__)

    def cargar_cache(self) -> Dict:
        """Carga el archivo de caché desde disco"""
        if CACHE_FILE.exists():
            try:
                with open(CACHE_FILE, 'r', encoding='utf-8') as f:
                    return json.load(f)
            except Exception as e:
                self.logger.warning(f"Error al cargar caché: {e}")
        return {}

    def guardar_cache(self, cache: Dict) -> None:
        """Guarda el caché en disco"""
        try:
            with open(CACHE_FILE, 'w', encoding='utf-8') as f:
                json.dump(cache, f, indent=2, ensure_ascii=False)
            self.logger.info("Caché guardado en disco")
        except Exception as e:
            self.logger.error(f"Error al guardar caché: {e}")

    def obtener_del_cache(self, fecha: str, moneda: str) -> Optional[Dict]:
        """Obtiene una cotización del caché si existe"""
        cache = self.cargar_cache()
        if fecha in cache and moneda in cache[fecha]:
            self.logger.info(f"Cotización en CACHÉ: {moneda} en {fecha}")
            return cache[fecha][moneda]
        return None

    def guardar_en_cache(self, fecha: str, moneda: str, cotizacion: Dict) -> None:
        """Guarda una cotización en el caché"""
        cache = self.cargar_cache()
        if fecha not in cache:
            cache[fecha] = {}
        cache[fecha][moneda] = cotizacion
        self.guardar_cache(cache)
