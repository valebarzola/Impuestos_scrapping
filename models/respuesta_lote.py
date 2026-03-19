from pydantic import BaseModel
from typing import List
from models.cotizacion import Cotizacion


class RespuestaLote(BaseModel):
    moneda: str
    cotizaciones: List[Cotizacion]
