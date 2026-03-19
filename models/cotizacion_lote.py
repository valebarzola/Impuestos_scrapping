from pydantic import BaseModel
from typing import List


class CotizacionLote(BaseModel):
    fechas: List[str]
