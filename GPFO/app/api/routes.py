# ═══════════════════════════════════════════════════════════
# Routes - versão FastAPI
# ═══════════════════════════════════════════════════════════

from fastapi import (
    APIRouter,
    UploadFile,
    File,
    Form,
    HTTPException
)

from tempfile import NamedTemporaryFile

from pathlib import Path

from fastapi.responses import FileResponse

from app.services.teams import importar_excel

from app.services.exportacao import exportar_formacao_excel

from app.services.formacoes import (
    listar_formacoes,
    obter_formacao,
    listar_sessoes,
    apagar_sessao,
    apagar_formacao
)

from app.schemas.formacao import Formacao
from app.schemas.sessao import Sessao

EXPORT_DIR = Path("exports")
EXPORT_DIR.mkdir(exist_ok=True)

router = APIRouter()

@router.get("/")
def inicio():

    return {
        "programa": "GPFO",
        "versao": "2.0",
        "status": "Online"
    }

@router.get(
    "/formacoes",
    response_model=list[Formacao]
)
def get_formacoes():

    return listar_formacoes()

@router.get(
    "/formacoes/{codigo}",
    response_model=Formacao
)
def get_formacao(codigo: str):

    formacao = obter_formacao(codigo)

    if formacao is None:

        raise HTTPException(
            status_code=404,
            detail="Formação não encontrada."
        )

    return formacao

@router.get(
    "/formacoes/{codigo}/sessoes",
    response_model=list[Sessao]
)
def get_sessoes(codigo: str):

    return listar_sessoes(codigo)

@router.delete("/sessao/{sessao_id}")
def delete_sessao(sessao_id: int):

    apagar_sessao(sessao_id)

    return {

        "status": "ok",

        "mensagem": "Sessão eliminada."

    }

@router.delete("/formacao/{codigo}")
def delete_formacao(codigo: str):

    apagar_formacao(codigo)

    return {

        "status": "ok",

        "mensagem": "Formação eliminada."

    }

@router.post("/importar")
async def importar(
    arquivo: UploadFile = File(...),
    codigo: str = Form(...),
    nome: str = Form(""),
    hora_inicio: str = Form(""),
    hora_fim: str = Form("")
):

    sessao = await importar_excel(
        arquivo,
        codigo,
        nome,
        hora_inicio,
        hora_fim
    )

    return {

        "status": "ok",

        "sessao": sessao

    }

@router.get("/formacoes/{codigo}/exportar")
def exportar(codigo: str):

    caminho = EXPORT_DIR / f"{codigo}.xlsx"

    exportar_formacao_excel(
        codigo,
        str(caminho)
    )

    return FileResponse(
        path=str(caminho),
        filename=f"{codigo}.xlsx",
        media_type="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
    )