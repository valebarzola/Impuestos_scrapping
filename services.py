import json
from datetime import datetime, timedelta
from typing import List, Optional, Dict, Tuple
from bs4 import BeautifulSoup
from fastapi import FastAPI, HTTPException, Query
from fastapi.responses import JSONResponse
from fastapi.middleware.cors import CORSMiddleware
import config
from cache.cache_config import cache
import logger
from logger import config_logger



class logica_service:

    def __init__(self):
        self.cache = cache()  
        config_logger()



    def restar_dias_habiles(fecha_str: str, dias: int = 1, max_retroceso: int = 7) -> str:
        """
        Resta días hábiles a una fecha, con máximo retroceso de 7 días
        
        Args:
            fecha_str: formato DD/MM/YYYY
            dias: cantidad de días a restar
            max_retroceso: máximo de días a retroceder
            
        Returns:
            Fecha en formato DD/MM/YYYY
        """
        fecha = datetime.strptime(fecha_str, "%d/%m/%Y")
        fecha = fecha - timedelta(days=dias)
        return fecha.strftime("%d/%m/%Y")

    def buscar_ultima_cotizacion_disponible(
        self,
        fecha_solicitada: str,
        moneda: str,
        max_retroceso: int = 7
    ) -> Tuple[str, Dict]:
        """
        Busca la cotización del día anterior o retrocede hasta encontrar una válida
        
        Args:
            fecha_solicitada: formato DD/MM/YYYY
            moneda: código de moneda (ej: "DOL")
            max_retroceso: máximo de días a retroceder
            
        Returns:
            Tupla (fecha_encontrada, datos_cotizacion)
            
        Raises:
            HTTPException si no encuentra cotización en el rango
        """

         
        fecha_buscada = self.restar_dias_habiles(fecha_solicitada, 1)
        
        for intento in range(max_retroceso):
            # Intentar obtener del caché
            cotizacion_cache = self.cache.obtener_del_cache(fecha_buscada, moneda)
            if cotizacion_cache:
                return fecha_buscada, cotizacion_cache
            
            # Si no está en caché, consultar AFIP
            try:
                cotizaciones_afip = self.cache.obtener_cotizaciones_de_afip(fecha_buscada)
                if moneda in cotizaciones_afip:
                    datos = cotizaciones_afip[moneda]
                    self.cache.guardar_en_cache(fecha_buscada, moneda, datos)
                    return fecha_buscada, datos
            except HTTPException as e:
                # Si es error 503 (servicio no disponible), relanzar inmediatamente
                if e.status_code == 503:
                    raise
                # Si es otro error (404, 400), continuar retrocediendo
                pass
            
            # Retroceder un día más
            fecha_buscada = self.cache.restar_dias_habiles(fecha_buscada, 1)
        
        # No encontró en ninguno de los últimos 7 días
        logger.error(
            f"No hay cotización para {moneda} en los últimos {max_retroceso} días "
            f"anteriores a {fecha_solicitada}"
        )
        raise HTTPException(
            status_code=404,
            detail=f"No hay cotización disponible para {moneda} en los últimos {max_retroceso} días"
        )
