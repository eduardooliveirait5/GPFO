from fastapi import FastAPI

from app.database import init_db
from app.api.routes import router

app = FastAPI(
    title="Gestor de Presenças - Formação Online",
    version="1.0"
)

from app.exceptions import register_exception_handlers
register_exception_handlers(app)

@app.on_event("startup")
def startup():

    init_db()

app.include_router(router)