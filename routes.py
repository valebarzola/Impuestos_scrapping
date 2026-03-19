import logging
from datetime import datetime
from typing import List

from fastapi import APIRouter, HTTPException, Query

from config import MONEDAS_SOPORTADAS, CACHE_FILE
from models import Cotizacion, CotizacionLote, EstadoCache
from scrapper.scrapping import validar_formato_fecha
from services import buscar_ultima_cotizacion_disponible, obtener_del_cache, cargar_cache

logger = logging.getLogger(__name__)
router = APIRouter()


@router.get("/cotizacion", response_model=Cotizacion, tags=["Cotizaciones"])
async def obtener_cotizacion(
    fecha: str = Query(..., description="Fecha de oficialización (DD/MM/YYYY)"),
    moneda: str = Query(..., description="Código de moneda (ej: DOL, EUR, BRL)")
) -> Cotizacion:
    if not validar_formato_fecha(fecha):
        raise HTTPException(status_code=400, detail=f"Formato de fecha inválido. Use DD/MM/YYYY. Recibido: {fecha}")

    moneda = moneda.upper()
    if moneda not in MONEDAS_SOPORTADAS:
        raise HTTPException(status_code=400, detail=f"Moneda no soportada: {moneda}. Soportadas: {', '.join(MONEDAS_SOPORTADAS.keys())}")

    try:
        fecha_cotizacion, datos = buscar_ultima_cotizacion_disponible(fecha, moneda)
    except LookupError as e:
        raise HTTPException(status_code=404, detail=str(e))
    except ConnectionError as e:
        raise HTTPException(status_code=503, detail=str(e))

    fuente = "cache" if obtener_del_cache(fecha_cotizacion, moneda) else "afip"

    return Cotizacion(
        fecha_solicitada=fecha,
        fecha_cotizacion=fecha_cotizacion,
        moneda=moneda,
        tipo_cambio_comprador=datos["tipo_comprador"],
        tipo_cambio_vendedor=datos["tipo_vendedor"],
        fuente=fuente
    )


@router.post("/cotizacion/lote", response_model=List[Cotizacion], tags=["Cotizaciones"])
async def obtener_cotizacion_lote(
    moneda: str = Query(..., description="Código de moneda"),
    body: CotizacionLote = None
) -> List[Cotizacion]:
    if not body:
        raise HTTPException(status_code=400, detail="Body JSON requerido con campo 'fechas'")

    moneda = moneda.upper()
    if moneda not in MONEDAS_SOPORTADAS:
        raise HTTPException(status_code=400, detail=f"Moneda no soportada: {moneda}")

    resultados = []
    for fecha in body.fechas:
        try:
            resultado = await obtener_cotizacion(fecha=fecha, moneda=moneda)
            resultados.append(resultado)
        except HTTPException as e:
            logger.warning(f"Error en {fecha}: {e.detail}")

    if not resultados:
        raise HTTPException(status_code=404, detail="No se encontraron cotizaciones para ninguna de las fechas")

    return resultados


@router.get("/cache/estado", response_model=EstadoCache, tags=["Administración"])
async def estado_cache() -> EstadoCache:
    cache = cargar_cache()
    fechas = set(cache.keys())
    monedas = set(m for v in cache.values() for m in v.keys())

    return EstadoCache(
        total_fechas=len(fechas),
        ultima_fecha_consultada=max(fechas) if fechas else None,
        monedas_en_cache=sorted(monedas),
        ruta_cache=str(CACHE_FILE)
    )


@router.delete("/cache/limpiar", tags=["Administración"])
async def limpiar_cache():
    try:
        if CACHE_FILE.exists():
            CACHE_FILE.unlink()
        return {"mensaje": "Caché limpiado correctamente"}
    except Exception as e:
        raise HTTPException(status_code=500, detail=f"Error al limpiar caché: {e}")


@router.get("/salud", tags=["Health Check"])
async def health_check():
    return {"estado": "ok", "timestamp": datetime.now().isoformat(), "version": "1.0.0"}
