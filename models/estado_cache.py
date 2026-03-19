from pydantic import BaseModel
from typing import List, Optional


class EstadoCache(BaseModel):
    total_fechas: int
    ultima_fecha_consultada: Optional[str]
    monedas_en_cache: List[str]
    ruta_cache: str
