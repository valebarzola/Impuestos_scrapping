# API Cotizaciones ARCA

API REST para obtener tipos de cambio del sitio público de ARCA (ex-AFIP).
Pensada para ser consumida desde macros de Excel via HTTP.

---

## Archivos principales

| Archivo | Descripción |
|---------|-------------|
| `main.py` | Entry point. Inicializa el logger, registra las rutas y levanta el servidor con uvicorn. |
| `app.py` | Instancia de FastAPI con configuración de CORS (permite acceso desde cualquier origen). |
| `routes.py` | Define todos los endpoints de la API. |
| `services.py` | Lógica de negocio: manejo del caché en disco y búsqueda de la última cotización disponible. |
| `config.py` | Constantes globales: URL de ARCA, monedas soportadas, rutas del caché. |
| `logger.py` | Configuración del sistema de logging. |
| `scrapper/scrapping.py` | Realiza el scraping del sitio de ARCA con BeautifulSoup. Expone dos funciones: una que devuelve todas las monedas y otra que devuelve solo el dólar comprador. |
| `models/` | Modelos Pydantic usados como esquemas de request/response. |
| `macro_excel.vba` | Macro VBA lista para instalar en Excel. Consulta la API automáticamente al ingresar una fecha. |
| `cache/cotizaciones.json` | Caché local de cotizaciones ya consultadas. Se genera automáticamente. |

---

## Levantar la API

```bash
python main.py
```

La API queda disponible en `http://localhost:8000`.

---

## Endpoints

### Cotización completa (todas las monedas soportadas)
```bash
curl "http://localhost:8000/cotizacion?fecha=22/11/2024&moneda=DOL"
```

Monedas soportadas: `DOL`, `EUR`, `BRL`, `UYU`, `CLP`

Respuesta:
```json
{
  "fecha_solicitada": "22/11/2024",
  "fecha_cotizacion": "21/11/2024",
  "moneda": "DOL",
  "tipo_cambio_comprador": 1012.5,
  "tipo_cambio_vendedor": 1015.0,
  "fuente": "afip"
}
```

---

### Dólar comprador (solo el valor)
```bash
curl "http://localhost:8000/dolar/comprador?fecha=22/11/2024"
```

Respuesta:
```json
{

  "tipo_cambio_comprador": 1012.5
}
```

---

### Cotización en lote (múltiples fechas)
```bash
curl -X POST "http://localhost:8000/cotizacion/lote?moneda=DOL" \
  -H "Content-Type: application/json" \
  -d '{"fechas": ["20/11/2024", "21/11/2024", "22/11/2024"]}'
```

---

### Estado del caché
```bash
curl "http://localhost:8000/cache/estado"
```

---

### Limpiar caché
```bash
curl -X DELETE "http://localhost:8000/cache/limpiar"
```

---

### Health check
```bash
curl "http://localhost:8000/salud"
```

---

## Regla de negocio

Al consultar una fecha X, la API devuelve la cotización del día X-1.
Si ese día no tiene datos (fin de semana, feriado), retrocede hasta 7 días para encontrar la última cotización válida.

---

## Documentación interactiva

Con la API corriendo, acceder a:
- Swagger UI: `http://localhost:8000/docs`
- ReDoc: `http://localhost:8000/redoc`
