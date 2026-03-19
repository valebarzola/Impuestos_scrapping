from pydantic import BaseModel


class Cotizacion(BaseModel):
    fecha_solicitada: str
    fecha_cotizacion: str
    moneda: str
    tipo_cambio_comprador: float
    tipo_cambio_vendedor: float
    fuente: str
