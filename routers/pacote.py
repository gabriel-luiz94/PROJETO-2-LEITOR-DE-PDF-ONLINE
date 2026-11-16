"""
routers/pacote.py — Pacote de projeto (TASK-177): exportar tudo o que um projeto tem num `.zip` e importar num projeto (com prévia).

  GET  /api/pacote/partes                          as partes que existem e seus rótulos
  GET  /api/pacote/exportar?projeto=&partes=a,b    baixa o `.zip` (qualquer usuário autenticado)
  POST /api/pacote/previa   (multipart: file, projeto)             o que mudaria, por parte — NÃO grava nada
  POST /api/pacote/importar (multipart: file, projeto, partes, confirmar=true)   (admin) grava as partes escolhidas
"""
from typing import Optional

from fastapi import APIRouter, File, Form, HTTPException, Request, UploadFile
from fastapi.responses import Response

from middleware.auth_middleware import get_current_user_from_state
from services import pacote_projeto as pp

router = APIRouter(prefix="/api/pacote", tags=["pacote"])


def _admin(request: Request) -> dict:
    user = get_current_user_from_state(request)
    if user.get("role") != "admin":
        raise HTTPException(status_code=403, detail="Acesso negado. Apenas administradores podem importar um pacote de projeto.")
    return user


def _partes(texto) -> list:
    return [p.strip() for p in str(texto or "").split(",") if p.strip()]


@router.get("/partes")
def listar_partes(request: Request):
    get_current_user_from_state(request)
    return {"partes": [{"parte": p, "rotulo": pp.ROTULOS[p]} for p in pp.PARTES], "versao_formato": pp.VERSAO_FORMATO}


@router.get("/embutidos")
def embutidos(request: Request):
    """Pacotes prontos que vêm com o programa (ponto de partida): o de exemplo para quem cria uma base nova, o de Pernambuco…"""
    get_current_user_from_state(request)
    return {"pacotes": [{"id": m["id"], "nome": m.get("nome", m["id"]), "descricao": m.get("descricao", ""), "partes": [p for p in pp.PARTES if p in m["partes"]]} for m in pp.pacotes_embutidos().values()]}


@router.get("/embutidos/{pacote_id}.zip")
def embutido_zip(request: Request, pacote_id: str):
    get_current_user_from_state(request)
    try:
        conteudo = pp.empacotar_embutido(pacote_id)
    except pp.ErroPacote as e:
        raise HTTPException(status_code=404, detail=str(e))
    return Response(content=conteudo, media_type="application/zip", headers={"Content-Disposition": f'attachment; filename="pacote_{pacote_id}.zip"'})


@router.get("/exportar")
def exportar(request: Request, projeto: str, partes: str = "", descricao: str = ""):
    get_current_user_from_state(request)
    try:
        conteudo = pp.exportar(projeto, _partes(partes) or None, descricao)
    except pp.ErroPacote as e:
        raise HTTPException(status_code=400, detail=str(e))
    proj = pp.projeto_do_cadastro(projeto)
    return Response(content=conteudo, media_type="application/zip", headers={"Content-Disposition": f'attachment; filename="pacote_{(proj or {}).get("codigo", "projeto")}.zip"'})


@router.post("/previa")
async def previa(request: Request, file: UploadFile = File(...), projeto: str = Form(...)):
    get_current_user_from_state(request)
    try:
        return pp.previa(projeto, pp.ler_pacote(await file.read()))
    except pp.ErroPacote as e:
        raise HTTPException(status_code=400, detail=str(e))


@router.post("/importar")
async def importar(request: Request, file: UploadFile = File(...), projeto: str = Form(...), partes: Optional[str] = Form(""), confirmar: Optional[str] = Form("")):
    user = _admin(request)
    if str(confirmar).lower() not in ("1", "true", "sim"):
        raise HTTPException(status_code=400, detail="Confirme a importação depois de ver a prévia (nada foi gravado).")
    try:
        return pp.aplicar(projeto, pp.ler_pacote(await file.read()), _partes(partes) or None, user.get("user_id"), user.get("email"))
    except pp.ErroPacote as e:
        raise HTTPException(status_code=400, detail=str(e))
