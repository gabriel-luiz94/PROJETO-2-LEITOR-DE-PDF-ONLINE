"""
routers/autonomo.py — Controle do modo autônomo (TASK-031, fase C). Só administrador.

Liga/desliga a vigia da pasta, edita a configuração, lista o histórico e responde às pendências ("Sim" / "Sim para todos" / rejeitar),
reverte e reprocessa. A tela de controle é a fase D; este router é a API em que ela se apoia. Não registrar o conteúdo das planilhas em log.
"""
from typing import List, Optional

from fastapi import APIRouter, Depends, HTTPException, Request
from pydantic import BaseModel

from middleware.auth_middleware import require_role
from services.autonomo import config_autonomo, execucoes, pasta, pipeline

router = APIRouter(prefix="/api/autonomo", tags=["autonomo"], dependencies=[Depends(require_role("admin"))])


class ConfigRequest(BaseModel):
    pasta_base: Optional[str] = None
    pasta_entrada: Optional[str] = None
    pasta_processados: Optional[str] = None
    pasta_erros: Optional[str] = None
    pasta_saida: Optional[str] = None
    user_id: Optional[str] = None
    intervalo_s: Optional[int] = None
    estabilizacao_s: Optional[int] = None


class DecisaoRequest(BaseModel):
    indices: Optional[List[int]] = None        # None = "Sim para todos" (todas as pendências DESTE arquivo)


def _publica(ex: dict) -> dict:
    """A execução sem os blocos grandes: as pendências aparecem como {indice, descricao}."""
    d = {k: v for k, v in ex.items() if k not in ("originais", "diff")}
    diff, dec = ex.get("diff") or {}, ex.get("decisoes") or {}
    abertas = [i for i in dec.get("pendentes", []) if i not in dec.get("confirmadas", []) and i not in dec.get("rejeitadas", [])]
    d["pendencias"] = [{"indice": i, "descricao": (diff.get("frases") or [])[i]} for i in abertas] if diff else []
    return d


def _acao(funcao, *args, **kw):
    try:
        with pasta.TRAVA:
            return funcao(*args, **kw)
    except pipeline.ErroPipeline as e:
        raise HTTPException(status_code=400, detail=str(e))


@router.get("/config")
def ler_config(request: Request):
    cfg = config_autonomo.carregar()
    user = getattr(request.state, "user", None) or {}
    return {"config": cfg, "pastas": config_autonomo.pastas(cfg), "usuario_atual": {"user_id": user.get("user_id"), "email": user.get("email")}}


@router.put("/config")
def gravar_config(req: ConfigRequest):
    cfg = {**config_autonomo.carregar(), **req.model_dump(exclude_none=True)}
    try:
        novo = config_autonomo.salvar(cfg)
    except config_autonomo.ErroConfig as e:
        raise HTTPException(status_code=400, detail=str(e))
    return {"config": novo, "pastas": config_autonomo.pastas(novo)}


@router.post("/ligar")
def ligar():
    cfg = config_autonomo.carregar()
    if not cfg.get("user_id"):
        raise HTTPException(status_code=400, detail="Escolha o usuário dono das obras antes de ligar o modo autônomo.")
    cfg = config_autonomo.salvar({**cfg, "ligado": True})
    config_autonomo.garantir_pastas(cfg)
    pasta.obter_vigia().iniciar()
    return {"status": pasta.obter_vigia().status(), "pastas": config_autonomo.pastas(cfg)}


@router.post("/desligar")
def desligar():
    config_autonomo.salvar({**config_autonomo.carregar(), "ligado": False})
    pasta.obter_vigia().parar()
    return {"status": pasta.obter_vigia().status()}


@router.get("/status")
def status():
    cfg = config_autonomo.carregar()
    pendentes = len(execucoes.listar("aguardando_confirmacao"))
    return {"status": pasta.obter_vigia().status(), "configurado": bool(cfg.get("user_id")), "aguardando_confirmacao": pendentes,
            "pastas": config_autonomo.pastas(cfg)}


@router.post("/varrer")
def varrer_agora():
    """Uma varredura imediata (a vigia continua como estava)."""
    cfg = config_autonomo.carregar()
    return {"resultados": pasta.obter_vigia().varrer(cfg)}


@router.get("/execucoes")
def listar_execucoes(status: Optional[str] = None, projeto: Optional[str] = None, limite: int = 100):
    return {"execucoes": execucoes.listar(status, projeto, max(1, min(limite, 500)))}


@router.get("/execucoes/{exec_id}")
def ver_execucao(exec_id: str):
    ex = execucoes.buscar(exec_id)
    if not ex:
        raise HTTPException(status_code=404, detail="Execução não encontrada.")
    return _publica(ex)


@router.post("/execucoes/{exec_id}/confirmar")
def confirmar(exec_id: str, req: DecisaoRequest):
    return _publica(_acao(pipeline.confirmar, exec_id, req.indices))


@router.post("/execucoes/{exec_id}/rejeitar")
def rejeitar(exec_id: str, req: DecisaoRequest):
    return _publica(_acao(pipeline.rejeitar, exec_id, req.indices))


@router.post("/execucoes/{exec_id}/reverter")
def reverter(exec_id: str):
    return _publica(_acao(pipeline.reverter, exec_id))


@router.post("/execucoes/{exec_id}/reprocessar")
def reprocessar(exec_id: str):
    return _publica(_acao(pasta.reprocessar, exec_id))
