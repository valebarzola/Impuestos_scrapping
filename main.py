from logger import config_logger
from app import app
import routes
import uvicorn 

config_logger()
app.include_router(routes.router)

if __name__ == "__main__":
    
    uvicorn.run(app, host="0.0.0.0", port=8000, log_level="info")
