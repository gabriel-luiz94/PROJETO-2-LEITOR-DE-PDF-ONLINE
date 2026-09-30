"""
routers/validacao_regras.py — Regras de domínio da validação das planilhas (TASK-013/016, ADR-004 camada 2).

Regras são dados editáveis pelo admin (services/regras_dominio.py define a linguagem v2 e o motor). Mesmo padrão de
routers/regras_leitor.py: nuvem primeiro com fallback local, validação antes de salvar, histórico com reversão,
escrita só para admin e erro de sincronização visível. "DEFAULT" vale para o projeto sem versão própria.

Formato guardado em `regras_json` (TASK-018, camadas): o DEFAULT guarda o container v2 completo
{"versao": 2, "grupos": {...}, "regras": [...]}; cada projeto guarda só o *overlay* (services/regras_camadas.py) com o que
difere do padrão. Os formatos antigos (lista v1 da TASK-013; container v2 completo num projeto) continuam sendo lidos e
viram overlay equivalente na leitura — sem migração em lote e sem SQL novo.

Não registrar o conteúdo das planilhas em log.
"""
import json

from fastapi import APIRouter, Depends, HTTPException, Request
from pydantic import BaseModel

from config import REGRAS_DOMINIO_SEED_PATH, logger
from database import get_connection
from middleware.auth_middleware import require_role
from services.regras_camadas import (efetivo, overlay_de_bruto,
                                     overlay_de_efetivo, overlay_vazio, overlay_vazio_de)
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


def _padrao() -> dict:
    bruto = _buscar_um(DEFAULT)
    return _container(bruto) if bruto is not None else {"grupos": {}, "regras": []}


def _overlay_do_projeto(projeto_codigo: str, padrao: dict):
    """Overlay do projeto (convertido se estiver no formato antigo) ou None se ele não tem versão própria."""
    bruto = _buscar_um(projeto_codigo)
    return None if bruto is None else overlay_de_bruto(padrao, bruto)


def resolver(projeto_codigo: str = DEFAULT):
    """(regras, grupos, grupos_origem, avisos, overlay|None) — efetivo do projeto = padrão + overlay."""
    padrao = _padrao()
    overlay = None if (projeto_codigo or DEFAULT) == DEFAULT else _overlay_do_projeto(projeto_codigo, padrao)
    regras, grupos, grupos_origem, avisos = efetivo(padrao, overlay or overlay_vazio())
    return regras, grupos, grupos_origem, avisos, overlay


def regras_efetivas(projeto_codigo: str = DEFAULT):
    """(regras v2, grupos, projeto_de_origem): padrão + ajustes do projeto. Usada por /api/validacao/planilhas."""
    regras, grupos, _, _, overlay = resolver(projeto_codigo)
    return regras, grupos, (projeto_codigo if overlay is not None else DEFAULT)


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


def _erros_de_edicao(regras: list, grupos: dict):
    erros = validar_conjunto(regras, grupos)
    if erros:
        raise HTTPException(status_code=400, detail={"erros": erros})


def _salvar(projeto_codigo: str, regras: list, grupos: dict | None, usuario_email: str | None):
    """DEFAULT: grava o container completo. Projeto: `regras` é o efetivo editado; grava só a diferença para o padrão."""
    if projeto_codigo == DEFAULT:
        if grupos is None:
            grupos = _padrao()["grupos"]
        _erros_de_edicao(regras, grupos)
        return _gravar(DEFAULT, {"versao": 2, "grupos": grupos,
                                 "regras": [_limpa(normalizar_regra(r)) for r in regras]}, usuario_email)
    padrao = _padrao()
    if grupos is None:
        grupos = resolver(projeto_codigo)[1]
    _erros_de_edicao(regras, grupos)
    return _salvar_overlay(projeto_codigo, overlay_de_efetivo(padrao, regras, grupos), usuario_email)


def _salvar_overlay(projeto_codigo: str, overlay: dict, usuario_email: str | None):
    padrao = _padrao()
    regras, grupos, _, _ = efetivo(padrao, overlay)
    _erros_de_edicao(regras, grupos)
    return _gravar(projeto_codigo, overlay, usuario_email)


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
    regras, grupos, grupos_origem, avisos, overlay = resolver(projeto_codigo)
    personalizado = overlay is not None and not overlay_vazio_de(overlay)
    return {"regras": _com_frase(regras, grupos), "grupos": grupos, "grupos_origem": grupos_origem, "avisos": avisos,
            "versao_de": projeto_codigo if overlay is not None else DEFAULT, "personalizado": personalizado}


@router.post("", dependencies=[Depends(require_role("admin"))])
def salvar(payload: SalvarRegrasPayload, request: Request):
    return _salvar(payload.projeto_codigo, payload.regras, payload.grupos, _extrair_email(request))


@router.get("/historico", dependencies=[Depends(require_role("admin"))])
def historico(projeto_codigo: str = DEFAULT):
    padrao = _padrao()
    itens = []
    for h in _listar_historico(projeto_codigo):
        if projeto_codigo == DEFAULT:
            c = _container(h["bruto"])
            regras, grupos = c["regras"], c["grupos"]
        else:
            regras, grupos, _, _ = efetivo(padrao, overlay_de_bruto(padrao, h["bruto"]))
        itens.append({"id": h["id"], "criado_em": h["criado_em"], "criado_por": h["criado_por"],
                      "regras": regras, "grupos": grupos})
    return {"historico": itens}


@router.post("/reverter", dependencies=[Depends(require_role("admin"))])
def reverter(payload: ReverterPayload, request: Request):
    versao = next((h for h in _listar_historico(payload.projeto_codigo) if h["id"] == payload.historico_id), None)
    if versao is None:
        raise HTTPException(status_code=404, detail="Versão de histórico não encontrada para este projeto.")
    if payload.projeto_codigo == DEFAULT:
        c = _container(versao["bruto"])
        return _salvar(DEFAULT, c["regras"], c["grupos"], _extrair_email(request))
    return _salvar_overlay(payload.projeto_codigo, overlay_de_bruto(_padrao(), versao["bruto"]), _extrair_email(request))


@router.post("/restaurar-semente", dependencies=[Depends(require_role("admin"))])
def restaurar_semente(payload: ProjetoPayload, request: Request):
    """DEFAULT: volta ao conteúdo da semente. Projeto: descarta os ajustes do projeto (volta a ser o padrão puro)."""
    if payload.projeto_codigo != DEFAULT:
        return _gravar(payload.projeto_codigo, overlay_vazio(), _extrair_email(request))
    semente = _ler_semente()
    return _salvar(DEFAULT, semente["regras"], semente["grupos"], _extrair_email(request))


@router.post("/adicionar-novas", dependencies=[Depends(require_role("admin"))])
def adicionar_novas(payload: ProjetoPayload, request: Request):
    """Só para o DEFAULT: acrescenta as regras (e grupos) da semente que ainda não existem — regras sempre DESLIGADAS,
    nada existente é alterado. Projetos herdam o padrão sozinhos (TASK-018): nada a fazer neles."""
    if payload.projeto_codigo != DEFAULT:
        return {"adicionadas": [], "grupos_adicionados": [],
                "mensagem": "Projetos acompanham o padrão automaticamente; use este botão no DEFAULT."}
    semente = _ler_semente()
    padrao = _padrao()
    existentes = {r.get("id") for r in padrao["regras"]}
    novas = [{**r, "ativa": False} for r in semente["regras"] if r.get("id") not in existentes]
    grupos_novos = {n: v for n, v in semente["grupos"].items() if n not in padrao["grupos"]}
    if novas or grupos_novos:
        _salvar(DEFAULT, padrao["regras"] + novas, {**padrao["grupos"], **grupos_novos}, _extrair_email(request))
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
