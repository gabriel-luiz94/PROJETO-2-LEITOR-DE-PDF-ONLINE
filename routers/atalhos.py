"""
routers/atalhos.py — Atalhos e favoritos dos modais geradores (Cabos e Postes/Estruturas), escopo "Projeto" (TASK-068).

Um JSON por (projeto, modal): {"itens": [{id, grupo, valor, comp?, autor}], "favoritos": [{id, nome?, linha, autor}]}.
Nuvem (Supabase) primeiro, fallback SQLite, erro de sincronização visível — mesmo padrão de routers/regras_vinculacao.py.
Quem pode o quê: qualquer usuário autenticado cria e altera os SEUS itens; só administrador exclui itens do projeto ou
altera os de outra pessoa. Atalhos "do computador" não passam por aqui (localStorage).
"""
import json
import re

from fastapi import APIRouter, HTTPException, Request
from pydantic import BaseModel

from config import logger
from database import get_connection
from services.supabase_client import get_supabase

router = APIRouter(prefix="/api/atalhos", tags=["atalhos"])

GRUPOS = {"cabos": {"ativos", "fases", "comps"}, "postes": {"postes", "mt", "bt"}}
OPERACOES = {"I", "*I", "R", "*R", "M", "*M"}
FASES = {"ABC", "A", "B", "C", "AC", "AB", "BC"}
LIMITE_ITENS = 200
LIMITE_FAVORITOS = 100
_RE_ID = re.compile(r"^[A-Za-z0-9_-]{1,40}$")
_PROIBIDO = re.compile(r"[<>\x00-\x1f]")


class AtalhosPayload(BaseModel):
    projeto_codigo: str
    dados: dict


def _texto(v, maximo, rotulo, erros, obrigatorio=True):
    if v is None or v == "":
        if obrigatorio:
            erros.append(f"{rotulo}: obrigatório.")
        return
    if not isinstance(v, str) or len(v) > maximo or _PROIBIDO.search(v):
        erros.append(f"{rotulo}: texto inválido (até {maximo} caracteres, sem < >).")


def _numero(v, rotulo, erros):
    if v is None or v == "":
        return
    if isinstance(v, bool) or not isinstance(v, (int, float)) or not (0 < v <= 10000):
        erros.append(f"{rotulo}: número inválido (0 a 10000).")


def validar(modal: str, dados) -> list:
    erros = []
    if modal not in GRUPOS:
        return [f"Modal desconhecido: {modal!r}."]
    if not isinstance(dados, dict):
        return ["O conteúdo precisa ser um objeto."]
    itens, favs = dados.get("itens", []), dados.get("favoritos", [])
    if not isinstance(itens, list) or not isinstance(favs, list):
        return ["'itens' e 'favoritos' precisam ser listas."]
    if len(itens) > LIMITE_ITENS:
        erros.append(f"Atalhos demais (máx. {LIMITE_ITENS}).")
    if len(favs) > LIMITE_FAVORITOS:
        erros.append(f"Favoritos demais (máx. {LIMITE_FAVORITOS}).")
    ids = set()
    for i, it in enumerate(itens[:LIMITE_ITENS], 1):
        r = f"Atalho #{i}"
        if not isinstance(it, dict):
            erros.append(f"{r}: precisa ser um objeto.")
            continue
        if not isinstance(it.get("id"), str) or not _RE_ID.match(it["id"]) or it["id"] in ids:
            erros.append(f"{r}: 'id' inválido ou repetido.")
        ids.add(it.get("id"))
        if it.get("grupo") not in GRUPOS[modal]:
            erros.append(f"{r}: grupo inválido para o modal {modal}.")
        _texto(it.get("valor"), 40, f"{r} valor", erros)
        _numero(it.get("comp"), f"{r} comp", erros)
    for i, f in enumerate(favs[:LIMITE_FAVORITOS], 1):
        r = f"Favorito #{i}"
        if not isinstance(f, dict) or not isinstance(f.get("linha"), dict):
            erros.append(f"{r}: precisa ter uma 'linha'.")
            continue
        if not isinstance(f.get("id"), str) or not _RE_ID.match(f["id"]) or f["id"] in ids:
            erros.append(f"{r}: 'id' inválido ou repetido.")
        ids.add(f.get("id"))
        _texto(f.get("nome"), 60, f"{r} nome", erros, obrigatorio=False)
        l = f["linha"]
        if l.get("op") not in OPERACOES:
            erros.append(f"{r}: operação inválida.")
        _texto(l.get("ativo"), 120, f"{r} ativo", erros, obrigatorio=False)
        if modal == "cabos":
            if l.get("fase") not in FASES:
                erros.append(f"{r}: fase inválida.")
            _numero(l.get("comp"), f"{r} comp", erros)
    return erros


def _ler(projeto: str, modal: str):
    supabase = get_supabase()
    if supabase:
        try:
            res = supabase.table("atalhos_modais").select("dados_json").eq("projeto_codigo", projeto).eq("modal", modal).execute()
            if res.data:
                return json.loads(res.data[0]["dados_json"])
        except Exception as e:
            logger.warning(f"Falha ao buscar atalhos_modais no Supabase: {e}")
    conn = get_connection()
    try:
        row = conn.execute("SELECT dados_json FROM atalhos_modais WHERE projeto_codigo = ? AND modal = ?", (projeto, modal)).fetchone()
    finally:
        conn.close()
    if row:
        try:
            return json.loads(row[0])
        except Exception:
            pass
    return {"itens": [], "favoritos": []}


def _regras_de_permissao(user, antigo: dict, novo: dict) -> list:
    """Não-admin: não exclui nada do projeto e só altera o que é seu."""
    if (user or {}).get("role") == "admin":
        return []
    eu = (user or {}).get("email") or (user or {}).get("user_id") or ""
    erros = []
    for chave in ("itens", "favoritos"):
        velhos = {x.get("id"): x for x in antigo.get(chave, []) if isinstance(x, dict)}
        novos = {x.get("id"): x for x in novo.get(chave, []) if isinstance(x, dict)}
        for id_ in velhos:
            if id_ not in novos:
                erros.append("Só administrador exclui atalhos/favoritos do projeto.")
                break
        for id_, x in novos.items():
            ant = velhos.get(id_)
            if ant is not None and ant.get("autor") != eu and {k: v for k, v in ant.items() if k != "autor"} != {k: v for k, v in x.items() if k != "autor"}:
                erros.append("Você só altera os atalhos/favoritos que criou.")
                break
    return erros


@router.get("/{modal}")
def obter(modal: str, projeto_codigo: str = "DEFAULT"):
    if modal not in GRUPOS:
        raise HTTPException(status_code=404, detail="Modal desconhecido.")
    return {"dados": _ler(projeto_codigo, modal)}


@router.put("/{modal}")
def gravar(modal: str, payload: AtalhosPayload, request: Request):
    erros = validar(modal, payload.dados)
    if erros:
        raise HTTPException(status_code=400, detail={"erros": erros[:20]})
    user = getattr(request.state, "user", None)
    antigo = _ler(payload.projeto_codigo, modal)
    negados = _regras_de_permissao(user, antigo, payload.dados)
    if negados:
        raise HTTPException(status_code=403, detail=negados[0])
    eu = (user or {}).get("email") or (user or {}).get("user_id") or ""
    anteriores = {x.get("id"): x for x in antigo.get("itens", []) + antigo.get("favoritos", []) if isinstance(x, dict)}
    limpo = {"itens": [], "favoritos": []}
    for chave in ("itens", "favoritos"):
        for x in payload.dados.get(chave, []):
            x = dict(x)
            x["autor"] = (anteriores.get(x["id"]) or {}).get("autor") or eu     # autor nunca vem do cliente
            limpo[chave].append(x)
    valor = json.dumps(limpo, ensure_ascii=False)
    conn = get_connection()
    try:
        conn.execute("INSERT OR REPLACE INTO atalhos_modais (projeto_codigo, modal, dados_json) VALUES (?, ?, ?)",
                     (payload.projeto_codigo, modal, valor))
        conn.commit()
    finally:
        conn.close()
    supabase = get_supabase()
    if supabase:
        try:
            supabase.table("atalhos_modais").upsert({"projeto_codigo": payload.projeto_codigo, "modal": modal, "dados_json": valor},
                                                    on_conflict="projeto_codigo,modal").execute()
        except Exception as e:
            logger.warning(f"Erro ao salvar atalhos_modais no Supabase: {e}")
            raise HTTPException(status_code=500, detail=f"Atalhos salvos localmente, mas falharam ao sincronizar com a nuvem (Supabase): {e}")
    return {"status": "success", "dados": limpo}
