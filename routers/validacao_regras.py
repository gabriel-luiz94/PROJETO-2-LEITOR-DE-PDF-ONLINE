"""
routers/validacao_regras.py — Regras de domínio da validação das planilhas (TASK-013/016, ADR-004 camada 2).

Regras são dados editáveis pelo admin (services/regras_dominio.py define a linguagem v2 e o motor). Mesmo padrão de
routers/regras_leitor.py: nuvem primeiro com fallback local, validação antes de salvar, histórico com reversão,
escrita só para admin e erro de sincronização visível. "DEFAULT" vale para o projeto sem versão própria.

Formato guardado em `regras_json`: container v2 {"versao": 2, "grupos": {...}, "regras": [...]}; o formato antigo
(lista de regras v1, TASK-013) continua sendo lido e é convertido na leitura.

Não registrar o conteúdo das planilhas em log.
"""
import json

from fastapi import APIRouter, Depends, HTTPException, Request
from pydantic import BaseModel

from config import REGRAS_DOMINIO_SEED_PATH, logger
from database import get_connection
from middleware.auth_middleware import require_role
from services.regras_dominio import (avaliar, descrever, explicar, normalizar_regra, validar_conjunto)
from services.supabase_client import get_supabase
from services.validacao_planilhas import resumir

router = APIRouter(prefix="/api/validacao/regras", tags=["validacao-regras"])

DEFAULT = "DEFAULT"


class SalvarRegrasPayload(BaseModel):
    projeto_codigo: str = DEFAULT
    regras: list
    grupos: dict | None = None


class ProjetoPayload(BaseModel):
    projeto_codigo: str = DEFAULT


class ReverterPayload(BaseModel):
    projeto_codigo: str = DEFAULT
    historico_id: int


class TestarPayload(BaseModel):
    regras: list          # rascunho em edição — não é salvo
    grupos: dict | None = None
    cabos: list = []
    outros: list = []
    explicar: bool = False


class DescreverPayload(BaseModel):
    regra: dict
    grupos: dict | None = None


def _extrair_email(request: Request) -> str | None:
    user = getattr(request.state, "user", None)
    return user.get("email") if user else None


def _container(bruto) -> dict:
    """Bruto guardado (lista v1 ou container v2) → {"grupos": {...}, "regras": [regras v2]}."""
    if isinstance(bruto, list):
        return {"grupos": {}, "regras": [normalizar_regra(r) for r in bruto]}
    if isinstance(bruto, dict):
        return {"grupos": dict(bruto.get("grupos") or {}), "regras": [normalizar_regra(r) for r in bruto.get("regras") or []]}
    return {"grupos": {}, "regras": []}


def _limpa(regra: dict) -> dict:
    return {k: v for k, v in regra.items() if k not in ("origem", "oculta", "frase")}


def _com_frase(regras: list, grupos: dict) -> list:
    saida = []
    for r in regras:
        try:
            frase = descrever(r, grupos)
        except Exception:  # regra malformada gravada por outra versão: não derruba a listagem
            frase = ""
        saida.append({**r, "frase": frase})
    return saida


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


def regras_efetivas(projeto_codigo: str = DEFAULT):
    """(regras v2, grupos, projeto_de_origem): versão do projeto ou, na falta, a DEFAULT. Usada por /api/validacao/planilhas."""
    for codigo in dict.fromkeys([projeto_codigo or DEFAULT, DEFAULT]):
        bruto = _buscar_um(codigo)
        if bruto is not None:
            c = _container(bruto)
            return c["regras"], c["grupos"], codigo
    return [], {}, DEFAULT


# ═══════════════════════════════════════════════════════════════════════════
# ESCRITA
# ═══════════════════════════════════════════════════════════════════════════
def _gravar(projeto_codigo: str, objeto, usuario_email: str | None):
    valor = json.dumps(objeto, ensure_ascii=False)
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


def _salvar(projeto_codigo: str, regras: list, grupos: dict | None, usuario_email: str | None):
    if grupos is None:
        bruto = _buscar_um(projeto_codigo)
        grupos = _container(bruto if bruto is not None else _buscar_um(DEFAULT))["grupos"]
    erros = validar_conjunto(regras, grupos)
    if erros:
        raise HTTPException(status_code=400, detail={"erros": erros})
    return _gravar(projeto_codigo, {"versao": 2, "grupos": grupos,
                                    "regras": [_limpa(normalizar_regra(r)) for r in regras]}, usuario_email)


def _listar_historico(projeto_codigo: str):
    supabase = get_supabase()
    if supabase:
        try:
            res = (supabase.table("regras_dominio_historico").select("id, regras_json, criado_em, criado_por")
                   .eq("projeto_codigo", projeto_codigo).order("criado_em", desc=True).execute())
            if res.data:
                return [{"id": r["id"], "criado_em": r["criado_em"], "criado_por": r.get("criado_por"),
                         "bruto": json.loads(r["regras_json"])} for r in res.data]
        except Exception as e:
            logger.warning(f"Falha ao buscar histórico de regras_dominio no Supabase: {e}")

    conn = get_connection()
    rows = conn.execute(
        "SELECT id, regras_json, criado_em, criado_por FROM regras_dominio_historico "
        "WHERE projeto_codigo = ? ORDER BY criado_em DESC, id DESC", (projeto_codigo,)
    ).fetchall()
    conn.close()
    return [{"id": r[0], "criado_em": r[2], "criado_por": r[3], "bruto": json.loads(r[1])} for r in rows]


def _ler_semente() -> dict:
    try:
        with open(REGRAS_DOMINIO_SEED_PATH, "r", encoding="utf-8") as f:
            return _container(json.load(f))
    except (FileNotFoundError, json.JSONDecodeError):
        raise HTTPException(status_code=404, detail="Semente das regras de domínio não encontrada.")


# ═══════════════════════════════════════════════════════════════════════════
# ROTAS
# ═══════════════════════════════════════════════════════════════════════════
@router.get("")
def obter(projeto_codigo: str = DEFAULT):
    regras, grupos, origem = regras_efetivas(projeto_codigo)
    return {"regras": _com_frase(regras, grupos), "grupos": grupos, "versao_de": origem, "personalizado": origem != DEFAULT}


@router.post("", dependencies=[Depends(require_role("admin"))])
def salvar(payload: SalvarRegrasPayload, request: Request):
    return _salvar(payload.projeto_codigo, payload.regras, payload.grupos, _extrair_email(request))


@router.get("/historico", dependencies=[Depends(require_role("admin"))])
def historico(projeto_codigo: str = DEFAULT):
    itens = []
    for h in _listar_historico(projeto_codigo):
        c = _container(h["bruto"])
        itens.append({"id": h["id"], "criado_em": h["criado_em"], "criado_por": h["criado_por"],
                      "regras": c["regras"], "grupos": c["grupos"]})
    return {"historico": itens}


@router.post("/reverter", dependencies=[Depends(require_role("admin"))])
def reverter(payload: ReverterPayload, request: Request):
    versao = next((h for h in _listar_historico(payload.projeto_codigo) if h["id"] == payload.historico_id), None)
    if versao is None:
        raise HTTPException(status_code=404, detail="Versão de histórico não encontrada para este projeto.")
    c = _container(versao["bruto"])
    return _salvar(payload.projeto_codigo, c["regras"], c["grupos"], _extrair_email(request))


@router.post("/restaurar-semente", dependencies=[Depends(require_role("admin"))])
def restaurar_semente(payload: ProjetoPayload, request: Request):
    semente = _ler_semente()
    return _salvar(payload.projeto_codigo, semente["regras"], semente["grupos"], _extrair_email(request))


@router.post("/adicionar-novas", dependencies=[Depends(require_role("admin"))])
def adicionar_novas(payload: ProjetoPayload, request: Request):
    """Acrescenta à versão do projeto as regras (e grupos) da semente que ainda não existem nela — regras sempre
    DESLIGADAS, nada existente é alterado."""
    semente = _ler_semente()
    regras, grupos, _ = regras_efetivas(payload.projeto_codigo)
    existentes = {r.get("id") for r in regras}
    novas = [{**r, "ativa": False} for r in semente["regras"] if r.get("id") not in existentes]
    grupos_novos = {n: v for n, v in semente["grupos"].items() if n not in grupos}
    if novas or grupos_novos:
        _salvar(payload.projeto_codigo, regras + novas, {**grupos, **grupos_novos}, _extrair_email(request))
    return {"adicionadas": [r["id"] for r in novas], "grupos_adicionados": sorted(grupos_novos)}


@router.post("/descrever", dependencies=[Depends(require_role("admin"))])
def descrever_regra(payload: DescreverPayload):
    """Frase em português de UMA regra (rascunho do editor). Erros de schema voltam em `erros`, sem frase."""
    erros = validar_conjunto([payload.regra], payload.grupos)
    if erros:
        return {"frase": "", "erros": erros}
    return {"frase": descrever(payload.regra, payload.grupos or {}), "erros": []}


@router.post("/testar", dependencies=[Depends(require_role("admin"))])
def testar(payload: TestarPayload):
    """Roda um rascunho de regras (as desligadas não são avaliadas) contra linhas de exemplo, sem salvar."""
    erros = validar_conjunto(payload.regras, payload.grupos)
    if erros:
        raise HTTPException(status_code=400, detail={"erros": erros})
    achados = avaliar(payload.regras, payload.cabos, payload.outros, payload.grupos)
    resposta = {"achados": achados, "resumo": resumir(achados)}
    if payload.explicar:
        resposta["explicacoes"] = explicar(payload.regras, payload.cabos, payload.outros, payload.grupos)
    return resposta
