"""
routers/rede.py — Estado da rede e da cópia local do banco (TASK-121). Só tem efeito no modo desktop; na versão online (servidor) responde `ativa: false`.
"""
from fastapi import APIRouter

from config import APP_MODE

router = APIRouter(prefix="/api/rede", tags=["rede"])


@router.get("/estado")
def estado_da_rede():
    """{modo, ativa, online, caminho ("proxy"|"direto"|"nenhum"), proxy, copia: {ultima, em_andamento, erro, pendencias}}"""
    if APP_MODE != "desktop":
        return {"modo": APP_MODE, "ativa": False, "online": True}
    from services import copia_local, rede_adaptativa
    e = rede_adaptativa.estado()
    return {"modo": APP_MODE, **{k: e[k] for k in ("ativa", "online", "caminho", "proxy", "desde", "ultima_checagem")}, "copia": copia_local.estado()}


@router.post("/copiar")
def copiar_agora():
    """Dispara a cópia/mescla do banco local agora (em segundo plano). Só desktop."""
    if APP_MODE != "desktop":
        return {"status": "ignorado", "motivo": "versão online"}
    from services import copia_local
    return {"status": "iniciada" if copia_local.iniciar_em_segundo_plano("pedido do usuário") else "ignorado"}
