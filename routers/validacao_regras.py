"""
routers/validacao_regras.py — Regras de domínio da validação das planilhas (TASK-013, ADR-004 camada 2).

Regras são dados editáveis pelo admin (services/regras_dominio.py define o schema e o motor). Mesmo padrão de
routers/regras_leitor.py: nuvem primeiro com fallback local, validação antes de salvar, histórico com reversão,
escrita só para admin e erro de sincronização visível. "DEFAULT" vale para o projeto sem versão própria.

Não registrar o conteúdo das planilhas em log.
"""
import json

from fastapi import APIRouter, Depends, HTTPException, Request
from pydantic import BaseModel

from config import REGRAS_DOMINIO_SEED_PATH, logger
from database import get_connection
from middleware.auth_middleware import require_role
from services.regras_dominio import avaliar, validar_regras
from services.supabase_client import get_supabase
from services.validacao_planilhas import resumir

router = APIRouter(prefix="/api/validacao/regras", tags=["validacao-regras"])

DEFAULT = "DEFAULT"


class SalvarRegrasPayload(BaseModel):
    projeto_codigo: str = DEFAULT
    regras: list


class ProjetoPayload(BaseModel):
    projeto_codigo: str = DEFAULT


class ReverterPayload(BaseModel):
    projeto_codigo: str = DEFAULT
    historico_id: int


class TestarPayload(BaseModel):
    regras: list          # rascunho em edição — não é salvo
    cabos: list = []
    outros: list = []


def _extrair_email(request: Request) -> str | None:
    user = getattr(request.state, "user", None)
    return user.get("email") if user else None


# ═══════════════════════════════════════════════════════════════════════════
# LEITURA
# ═══════════════════════════════════════════════════════════════════════════
def _buscar_um(projeto_codigo: str):
    supabase = get_supabase()
    if supabase:
        try:
            res = supabase.table("regras_dominio").select("regras_json").eq("projeto_codigo", projeto_codigo).execute()
            if res.data:
                return json.loads(res.data[0]["regras_json"])
        except Exception as e:
            logger.warning(f"Falha ao buscar regras_dominio no Supabase: {e}")

    conn = get_connection()
    row = conn.execute("SELECT regras_json FROM regras_dominio WHERE projeto_codigo = ?", (projeto_codigo,)).fetchone()
    conn.close()
    return json.loads(row[0]) if row else None


def buscar_regras(projeto_codigo: str = DEFAULT):
    """(regras, projeto_de_origem): versão do projeto ou, na falta, a DEFAULT. ([], DEFAULT) se não houver.

    Função pública: usada também por POST /api/validacao/planilhas.
    """
    for codigo in dict.fromkeys([projeto_codigo or DEFAULT, DEFAULT]):
        regras = _buscar_um(codigo)
        if regras is not None:
            return regras, codigo
    return [], DEFAULT


# ═══════════════════════════════════════════════════════════════════════════
# ESCRITA
# ═══════════════════════════════════════════════════════════════════════════
def _salvar(projeto_codigo: str, regras: list, usuario_email: str | None):
    erros = validar_regras(regras)
    if erros:
        raise HTTPException(status_code=400, detail={"erros": erros})
    valor = json.dumps(regras, ensure_ascii=False)

    conn = get_connection()
    cursor = conn.cursor()
    cursor.execute("SELECT regras_json FROM regras_dominio WHERE projeto_codigo = ?", (projeto_codigo,))
    atual = cursor.fetchone()
    if atual:  # preserva a versão atual ANTES de sobrescrever
        cursor.execute(
            "INSERT INTO regras_dominio_historico (projeto_codigo, regras_json, criado_por) VALUES (?, ?, ?)",
            (projeto_codigo, atual[0], usuario_email)
        )
    cursor.execute("INSERT OR REPLACE INTO regras_dominio (projeto_codigo, regras_json) VALUES (?, ?)", (projeto_codigo, valor))
    conn.commit()
    conn.close()

    supabase = get_supabase()
    if supabase:
        try:
            if atual:
                supabase.table("regras_dominio_historico").insert({
                    "projeto_codigo": projeto_codigo, "regras_json": atual[0], "criado_por": usuario_email,
                }).execute()
            supabase.table("regras_dominio").upsert({"projeto_codigo": projeto_codigo, "regras_json": valor}).execute()
        except Exception as e:
            logger.warning(f"Erro ao salvar regras_dominio no Supabase: {e}")
            raise HTTPException(
                status_code=500,
                detail=f"Regras salvas localmente, mas falharam ao sincronizar com a nuvem (Supabase): {e}"
            )
    return {"status": "success"}


def _listar_historico(projeto_codigo: str):
    supabase = get_supabase()
    if supabase:
        try:
            res = (supabase.table("regras_dominio_historico").select("id, regras_json, criado_em, criado_por")
                   .eq("projeto_codigo", projeto_codigo).order("criado_em", desc=True).execute())
            if res.data:
                return [{"id": r["id"], "criado_em": r["criado_em"], "criado_por": r.get("criado_por"),
                         "regras": json.loads(r["regras_json"])} for r in res.data]
        except Exception as e:
            logger.warning(f"Falha ao buscar histórico de regras_dominio no Supabase: {e}")

    conn = get_connection()
    rows = conn.execute(
        "SELECT id, regras_json, criado_em, criado_por FROM regras_dominio_historico "
        "WHERE projeto_codigo = ? ORDER BY criado_em DESC, id DESC", (projeto_codigo,)
    ).fetchall()
    conn.close()
    return [{"id": r[0], "criado_em": r[2], "criado_por": r[3], "regras": json.loads(r[1])} for r in rows]


# ═══════════════════════════════════════════════════════════════════════════
# ROTAS
# ═══════════════════════════════════════════════════════════════════════════
@router.get("")
def obter(projeto_codigo: str = DEFAULT):
    regras, origem = buscar_regras(projeto_codigo)
    return {"regras": regras, "versao_de": origem, "personalizado": origem != DEFAULT}


@router.post("", dependencies=[Depends(require_role("admin"))])
def salvar(payload: SalvarRegrasPayload, request: Request):
    return _salvar(payload.projeto_codigo, payload.regras, _extrair_email(request))


@router.get("/historico", dependencies=[Depends(require_role("admin"))])
def historico(projeto_codigo: str = DEFAULT):
    return {"historico": _listar_historico(projeto_codigo)}


@router.post("/reverter", dependencies=[Depends(require_role("admin"))])
def reverter(payload: ReverterPayload, request: Request):
    versao = next((h for h in _listar_historico(payload.projeto_codigo) if h["id"] == payload.historico_id), None)
    if versao is None:
        raise HTTPException(status_code=404, detail="Versão de histórico não encontrada para este projeto.")
    return _salvar(payload.projeto_codigo, versao["regras"], _extrair_email(request))


@router.post("/restaurar-semente", dependencies=[Depends(require_role("admin"))])
def restaurar_semente(payload: ProjetoPayload, request: Request):
    try:
        with open(REGRAS_DOMINIO_SEED_PATH, "r", encoding="utf-8") as f:
            regras = json.load(f)
    except (FileNotFoundError, json.JSONDecodeError):
        raise HTTPException(status_code=404, detail="Semente das regras de domínio não encontrada.")
    return _salvar(payload.projeto_codigo, regras, _extrair_email(request))


@router.post("/adicionar-novas", dependencies=[Depends(require_role("admin"))])
def adicionar_novas(payload: ProjetoPayload, request: Request):
    """Acrescenta à versão do projeto as regras da semente que ainda não existem nela (sempre DESLIGADAS).

    Não altera, religa nem remove nenhuma regra existente. Existe porque a semente só é carregada quando ainda não
    há versão DEFAULT: instalações com regras já salvas (local ou no Supabase) não recebem as regras novas sozinhas.
    """
    try:
        with open(REGRAS_DOMINIO_SEED_PATH, "r", encoding="utf-8") as f:
            semente = json.load(f)
    except (FileNotFoundError, json.JSONDecodeError):
        raise HTTPException(status_code=404, detail="Semente das regras de domínio não encontrada.")
    atuais, _ = buscar_regras(payload.projeto_codigo)
    existentes = {r.get("id") for r in atuais}
    novas = [{**r, "ativa": False} for r in semente if r.get("id") not in existentes]
    if novas:
        _salvar(payload.projeto_codigo, atuais + novas, _extrair_email(request))
    return {"adicionadas": [r["id"] for r in novas]}


@router.post("/testar", dependencies=[Depends(require_role("admin"))])
def testar(payload: TestarPayload):
    """Roda um rascunho de regras (mesmo com ativa=false não são avaliadas) contra linhas de exemplo, sem salvar."""
    erros = validar_regras(payload.regras)
    if erros:
        raise HTTPException(status_code=400, detail={"erros": erros})
    achados = avaliar(payload.regras, payload.cabos, payload.outros)
    return {"achados": achados, "resumo": resumir(achados)}
