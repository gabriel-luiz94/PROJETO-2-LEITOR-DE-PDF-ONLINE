"""
routers/modais_montaveis.py — Modais montáveis por projeto (TASK-077, etapa 1).

Cada projeto tem uma lista de modais definidos em JSON declarativo (interpretados por static/modal_montavel_engine.js).
Leitura: qualquer usuário autenticado. Gravação: só administrador. Nuvem (Supabase) primeiro, fallback SQLite;
guarda na tabela já existente `atalhos_modais` com modal = "montaveis" (sem SQL novo).

Formato: {"versao": 1, "modais": [{"id", "nome", "tipo", "blocos": [...]}]}
Blocos: texto, campo, grade, interpretar, tabela_interna, saida_texto, destino, acao_adicionar (ver validar_modal).
Nada aqui executa código: regex são compiladas só para validar, expressões passam por uma lista branca aritmética.
"""
import ast
import json
import re

from fastapi import APIRouter, Depends, HTTPException
from pydantic import BaseModel

from config import logger
from database import get_connection
from middleware.auth_middleware import require_role
from services.supabase_client import get_supabase

router = APIRouter(prefix="/api/modais-montaveis", tags=["modais-montaveis"])

CHAVE = "montaveis"
TIPOS_MODAL = {"montagem", "colagem", "interpretativo", "processamento", "gerador_texto"}
TIPOS_BLOCO = {"texto", "campo", "grade", "interpretar", "tabela_interna", "saida_texto", "destino", "acao_adicionar"}
ENTRADAS = {"texto", "numero", "select", "radio"}
OPERACOES_CONDICAO = {"igual", "diferente", "vazio", "preenchido", "em"}
OPERACOES_LINHA = {"I", "*I", "R", "*R", "M", "*M"}
DESTINOS = {"outros", "cabos"}
MAX_MODAIS, MAX_BLOCOS, MAX_REGRAS, MAX_GERA, MAX_COLUNAS, MAX_LINHAS_INTERNAS = 30, 40, 100, 20, 20, 500
_RE_ID = re.compile(r"^[A-Za-z0-9_-]{1,40}$")
_RE_CHAR_EXPR = re.compile(r"^[0-9+\-*/().\s]+$")


class ModaisPayload(BaseModel):
    projeto_codigo: str
    dados: dict


def _texto(v, maximo, rotulo, erros, obrigatorio=True):
    if v is None or v == "":
        if obrigatorio:
            erros.append(f"{rotulo}: obrigatório.")
        return
    if not isinstance(v, str) or len(v) > maximo:
        erros.append(f"{rotulo}: texto de até {maximo} caracteres.")


def _id(v, rotulo, erros):
    if not isinstance(v, str) or not _RE_ID.match(v):
        erros.append(f"{rotulo}: id inválido (letras, números, _ e -, até 40).")


def expressao_segura(expr) -> bool:
    """Número, ou texto só com dígitos, + - * / ( ) e {n} (capturas). Confere a sintaxe trocando {n} por 1."""
    if isinstance(expr, bool):
        return False
    if isinstance(expr, (int, float)):
        return True
    if not isinstance(expr, str) or not expr.strip() or len(expr) > 100:
        return False
    s = re.sub(r"\{\d+\}", "1", expr)
    if not _RE_CHAR_EXPR.match(s):
        return False
    try:
        arvore = ast.parse(s.strip(), mode="eval")
    except SyntaxError:
        return False
    for no in ast.walk(arvore):
        if not isinstance(no, (ast.Expression, ast.BinOp, ast.UnaryOp, ast.Constant, ast.Add, ast.Sub, ast.Mult, ast.Div, ast.USub, ast.UAdd)):
            return False
        if isinstance(no, ast.Constant) and not isinstance(no.value, (int, float)):
            return False
    return True


def _regex(v, rotulo, erros):
    if not isinstance(v, str) or not v or len(v) > 300:
        erros.append(f"{rotulo}: regex obrigatória (até 300 caracteres).")
        return
    try:
        re.compile(v)
    except re.error as e:
        erros.append(f"{rotulo}: regex inválida ({e}).")


def _condicao(c, rotulo, erros):
    if c is None:
        return
    if not isinstance(c, dict) or c.get("op") not in OPERACOES_CONDICAO or not isinstance(c.get("campo"), str):
        erros.append(f"{rotulo}: 'visivel_se' precisa de campo e op ({sorted(OPERACOES_CONDICAO)}).")
        return
    if c["op"] in ("igual", "diferente") and not isinstance(c.get("valor"), (str, int, float)):
        erros.append(f"{rotulo}: 'visivel_se.valor' obrigatório.")
    if c["op"] == "em" and not isinstance(c.get("valor"), list):
        erros.append(f"{rotulo}: 'visivel_se.valor' precisa ser uma lista para 'em'.")


def _bloco(b, i, ids_por_tipo, erros):
    p = f"Bloco #{i + 1}"
    if not isinstance(b, dict) or b.get("tipo") not in TIPOS_BLOCO:
        erros.append(f"{p}: tipo precisa ser um de {sorted(TIPOS_BLOCO)}.")
        return
    t = b["tipo"]
    p = f"{p} ({t})"
    _condicao(b.get("visivel_se"), p, erros)
    if t != "texto" and t != "acao_adicionar":
        _id(b.get("id"), f"{p}.id", erros)
        if isinstance(b.get("id"), str):
            if b["id"] in ids_por_tipo["__todos"]:
                erros.append(f"{p}: id '{b['id']}' repetido.")
            ids_por_tipo["__todos"].add(b["id"])
            ids_por_tipo.setdefault(t, set()).add(b["id"])
    if t == "texto":
        _texto(b.get("texto"), 2000, f"{p}.texto", erros)
    elif t == "campo":
        _texto(b.get("rotulo"), 80, f"{p}.rotulo", erros)
        if b.get("entrada") not in ENTRADAS:
            erros.append(f"{p}.entrada: precisa ser um de {sorted(ENTRADAS)}.")
        if b.get("entrada") in ("select", "radio"):
            op = b.get("opcoes")
            if not isinstance(op, list) or not op or len(op) > 100 or not all(isinstance(x, str) and 0 < len(x) <= 80 for x in op):
                erros.append(f"{p}.opcoes: lista de 1 a 100 textos.")
    elif t == "grade":
        cols = b.get("colunas")
        if not isinstance(cols, list) or not 1 <= len(cols) <= MAX_COLUNAS:
            erros.append(f"{p}.colunas: de 1 a {MAX_COLUNAS}.")
        else:
            vistos = set()
            for c in cols:
                if not isinstance(c, dict):
                    erros.append(f"{p}.colunas: cada coluna precisa de id e rotulo.")
                    continue
                _id(c.get("id"), f"{p}.colunas.id", erros)
                _texto(c.get("rotulo"), 60, f"{p}.colunas.rotulo", erros)
                if c.get("id") in vistos:
                    erros.append(f"{p}.colunas: id '{c.get('id')}' repetido.")
                vistos.add(c.get("id"))
        lm = b.get("linhas_min", 5)
        if not isinstance(lm, int) or isinstance(lm, bool) or not 1 <= lm <= 50:
            erros.append(f"{p}.linhas_min: de 1 a 50.")
    elif t == "interpretar":
        regras = b.get("regras")
        if not isinstance(regras, list) or not 1 <= len(regras) <= MAX_REGRAS:
            erros.append(f"{p}.regras: de 1 a {MAX_REGRAS}.")
            regras = []
        for k, r in enumerate(regras):
            q = f"{p}, regra #{k + 1}"
            if not isinstance(r, dict) or not isinstance(r.get("quando"), dict):
                erros.append(f"{q}: precisa de 'quando' e 'gera'.")
                continue
            _regex(r["quando"].get("regex"), f"{q}.quando.regex", erros)
            if r["quando"].get("flags") is not None and (not isinstance(r["quando"]["flags"], str) or not re.fullmatch(r"[gimsu]{0,5}", r["quando"]["flags"])):
                erros.append(f"{q}.quando.flags: só letras gimsu.")
            gera = r.get("gera")
            if not isinstance(gera, list) or not 1 <= len(gera) <= MAX_GERA:
                erros.append(f"{q}.gera: de 1 a {MAX_GERA} itens.")
                continue
            for g in gera:
                if not isinstance(g, dict):
                    erros.append(f"{q}.gera: item inválido.")
                    continue
                _texto(g.get("ativo"), 120, f"{q}.gera.ativo", erros)
                if "qtd" in g and not expressao_segura(g["qtd"]):
                    erros.append(f"{q}.gera.qtd: só números, + - * / ( ) e {{n}} (ex.: {{1}}*20).")
                if g.get("operacao") is not None and g["operacao"] not in OPERACOES_LINHA:
                    erros.append(f"{q}.gera.operacao: uma de {sorted(OPERACOES_LINHA)}.")
    elif t == "tabela_interna":
        cols = b.get("colunas")
        if not isinstance(cols, list) or not 1 <= len(cols) <= MAX_COLUNAS or not all(isinstance(c, str) and 0 < len(c) <= 40 for c in cols):
            erros.append(f"{p}.colunas: lista de 1 a {MAX_COLUNAS} nomes.")
            cols = []
        linhas = b.get("linhas")
        if not isinstance(linhas, list) or len(linhas) > MAX_LINHAS_INTERNAS:
            erros.append(f"{p}.linhas: lista de até {MAX_LINHAS_INTERNAS} linhas.")
        else:
            for ln in linhas:
                if not isinstance(ln, list) or len(ln) != len(cols) or not all(x is None or isinstance(x, (str, int, float)) for x in ln):
                    erros.append(f"{p}.linhas: cada linha precisa ter {len(cols)} valores simples.")
                    break
    elif t == "saida_texto":
        _texto(b.get("modelo"), 4000, f"{p}.modelo", erros)
        _texto(b.get("rotulo"), 80, f"{p}.rotulo", erros, obrigatorio=False)
    elif t == "destino":
        op = b.get("opcoes", ["outros", "cabos"])
        if not isinstance(op, list) or not op or not all(x in DESTINOS for x in op):
            erros.append(f"{p}.opcoes: use {sorted(DESTINOS)}.")
        if b.get("padrao") is not None and b["padrao"] not in DESTINOS:
            erros.append(f"{p}.padrao: use {sorted(DESTINOS)}.")
    elif t == "acao_adicionar":
        _texto(b.get("rotulo"), 60, f"{p}.rotulo", erros, obrigatorio=False)


def validar_modal(m, idx=0) -> list:
    erros = []
    p = f"Modal #{idx + 1}"
    if not isinstance(m, dict):
        return [f"{p}: precisa ser um objeto."]
    _id(m.get("id"), f"{p}.id", erros)
    _texto(m.get("nome"), 40, f"{p}.nome", erros)
    if m.get("tipo") not in TIPOS_MODAL:
        erros.append(f"{p}.tipo: precisa ser um de {sorted(TIPOS_MODAL)}.")
    blocos = m.get("blocos")
    if not isinstance(blocos, list) or not 1 <= len(blocos) <= MAX_BLOCOS:
        erros.append(f"{p}.blocos: de 1 a {MAX_BLOCOS}.")
        return erros
    ids = {"__todos": set()}
    for i, b in enumerate(blocos):
        _bloco(b, i, ids, erros)
    # referências entre blocos
    for i, b in enumerate(blocos):
        if not isinstance(b, dict):
            continue
        q = f"{p}, bloco #{i + 1}"
        if b.get("tipo") == "interpretar" and b.get("fonte") not in ids.get("grade", set()):
            erros.append(f"{q} (interpretar): 'fonte' precisa ser o id de uma grade deste modal.")
        if b.get("tipo") == "saida_texto" and b.get("fonte") is not None and b["fonte"] not in ids.get("grade", set()):
            erros.append(f"{q} (saida_texto): 'fonte' precisa ser o id de uma grade deste modal.")
        if b.get("tipo") == "acao_adicionar":
            if b.get("fonte") not in ids.get("interpretar", set()):
                erros.append(f"{q} (acao_adicionar): 'fonte' precisa ser o id de um bloco 'interpretar'.")
            d = b.get("destino")
            if d not in DESTINOS and d not in ids.get("destino", set()):
                erros.append(f"{q} (acao_adicionar): 'destino' é 'outros', 'cabos' ou o id de um bloco 'destino'.")
        c = b.get("visivel_se")
        if isinstance(c, dict) and c.get("campo") not in ids.get("campo", set()):
            erros.append(f"{q}: 'visivel_se.campo' precisa ser o id de um campo deste modal.")
    return erros


def validar(dados) -> list:
    if not isinstance(dados, dict) or dados.get("versao") != 1 or not isinstance(dados.get("modais"), list):
        return ["Formato: precisa ser {\"versao\": 1, \"modais\": [...]}."]
    if len(dados["modais"]) > MAX_MODAIS:
        return [f"No máximo {MAX_MODAIS} modais por projeto."]
    erros, ids = [], set()
    for i, m in enumerate(dados["modais"]):
        erros += validar_modal(m, i)
        if isinstance(m, dict) and m.get("id") in ids:
            erros.append(f"Modal #{i + 1}: id '{m.get('id')}' repetido.")
        if isinstance(m, dict):
            ids.add(m.get("id"))
    return erros


def _vazio():
    return {"versao": 1, "modais": []}


def _ler(projeto: str) -> dict:
    supabase = get_supabase()
    if supabase:
        try:
            res = supabase.table("atalhos_modais").select("dados_json").eq("projeto_codigo", projeto).eq("modal", CHAVE).execute()
            if res.data:
                return json.loads(res.data[0]["dados_json"])
        except Exception as e:
            logger.warning(f"Falha ao buscar modais montáveis no Supabase: {e}")
    conn = get_connection()
    try:
        row = conn.execute("SELECT dados_json FROM atalhos_modais WHERE projeto_codigo = ? AND modal = ?", (projeto, CHAVE)).fetchone()
    finally:
        conn.close()
    if row:
        try:
            return json.loads(row[0])
        except Exception:
            pass
    return _vazio()


@router.get("")
def obter(projeto_codigo: str = "DEFAULT"):
    return {"dados": _ler(projeto_codigo)}


@router.post("/validar", dependencies=[Depends(require_role("admin"))])
def validar_rota(payload: ModaisPayload):
    return {"erros": validar(payload.dados)}


@router.put("", dependencies=[Depends(require_role("admin"))])
def gravar(payload: ModaisPayload):
    erros = validar(payload.dados)
    if erros:
        raise HTTPException(status_code=400, detail={"erros": erros[:30]})
    valor = json.dumps(payload.dados, ensure_ascii=False)
    conn = get_connection()
    try:
        conn.execute("INSERT OR REPLACE INTO atalhos_modais (projeto_codigo, modal, dados_json) VALUES (?, ?, ?)",
                     (payload.projeto_codigo, CHAVE, valor))
        conn.commit()
    finally:
        conn.close()
    supabase = get_supabase()
    if supabase:
        try:
            supabase.table("atalhos_modais").upsert({"projeto_codigo": payload.projeto_codigo, "modal": CHAVE, "dados_json": valor},
                                                    on_conflict="projeto_codigo,modal").execute()
        except Exception as e:
            logger.warning(f"Erro ao salvar modais montáveis no Supabase: {e}")
            raise HTTPException(status_code=500, detail=f"Modais salvos localmente, mas falharam ao sincronizar com a nuvem (Supabase): {e}")
    return {"status": "success", "dados": payload.dados}
