"""routers/manual.py — Manual de uso de regras e ajustes (TASK-030). Somente leitura; exige login (qualquer perfil)."""
from fastapi import APIRouter, HTTPException, Request
from fastapi.responses import PlainTextResponse

import config
from middleware.auth_middleware import get_current_user_from_state

router = APIRouter(prefix="/api/manual", tags=["manual"])


@router.get("")
def manual(request: Request):
    get_current_user_from_state(request)
    try:
        with open(config.MANUAL_PATH, "r", encoding="utf-8") as f:
            texto = f.read()
    except OSError:
        raise HTTPException(status_code=404, detail="Manual não encontrado.")
    return PlainTextResponse(texto, media_type="text/markdown; charset=utf-8")
