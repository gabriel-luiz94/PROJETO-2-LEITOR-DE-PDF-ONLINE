"""
routers/regras_leitor.py — Rotas para as regras do leitor (TASK-006): Processamento e
Classificação, duas tabelas por projeto que alimentam o motor de regras
(static/regras_leitor_engine.js). Mesmo padrão de sincronização de routers/regras.py
(get_regras_conversao/save_regras_conversao, TASK-003): nuvem primeiro, fallback local, erro de
sincronização visível.

TASK-007 acrescenta: validação de schema antes de salvar (nunca aceita um payload que quebraria o
motor), histórico de versões com reversão, e proteção por role admin nas rotas de escrita.
"""
import json
import re
from fastapi import APIRouter, HTTPException, Depends, Request
from pydantic import BaseModel
from database import get_connection
from services.supabase_client import get_supabase
from middleware.auth_middleware import require_role
from config import logger

router = APIRouter(prefix="/api/regras-leitor", tags=["regras-leitor"])


class RegrasLeitorPayload(BaseModel):
    projeto_codigo: str
    regras: list


class ReverterPayload(BaseModel):
    projeto_codigo: str
    historico_id: int


# ═══════════════════════════════════════════════════════════════════════════
# VALIDAÇÃO — espelha exatamente os campos que static/regras_leitor_engine.js
# de fato interpreta (ver comentário no topo do arquivo do motor). Uma regra
# que falhe aqui nunca chega a ser salva.
# ═══════════════════════════════════════════════════════════════════════════
MODOS_VALIDOS = {"DEFINIR", "SUBSTITUIR", "SUBSTITUIR_TOTAL"}
FASES_VALIDAS = {1, 2, 3}
CAMPOS_LISTA = ("cor_em", "cor_nao_em", "layer_em")


def _compila_regex(valor, campo: str, erros: list):
    if valor is None or valor == "":
        return
    if not isinstance(valor, str):
        erros.append(f"{campo}: precisa ser uma string.")
        return
    try:
        re.compile(valor)
    except re.error as e:
        erros.append(f"{campo}: regex inválida ({e}).")


def _valida_campos_lista(regra: dict, prefixo: str, erros: list):
    for campo in CAMPOS_LISTA:
        valor = regra.get(campo)
        if valor is not None and not isinstance(valor, list):
            erros.append(f"{prefixo}: '{campo}' precisa ser uma lista.")


def _valida_regra_processamento(regra: dict, idx: int, erros: list):
    prefixo = f"Regra de Processamento #{idx + 1}"

    ordem = regra.get("ordem")
    if not isinstance(ordem, (int, float)) or isinstance(ordem, bool):
        erros.append(f"{prefixo}: campo 'ordem' é obrigatório e precisa ser numérico.")

    fase = regra.get("fase", 3)
    if fase not in FASES_VALIDAS:
        erros.append(f"{prefixo}: campo 'fase' precisa ser 1, 2 ou 3 (recebido: {fase!r}).")

    modo = regra.get("modo") or "DEFINIR"
    if modo not in MODOS_VALIDOS:
        erros.append(f"{prefixo}: campo 'modo' precisa ser um de {sorted(MODOS_VALIDOS)} (recebido: {modo!r}).")

    if modo in ("SUBSTITUIR", "SUBSTITUIR_TOTAL") and not regra.get("ativo_regex"):
        erros.append(
            f"{prefixo}: modo '{modo}' exige o campo 'ativo_regex' preenchido "
            f"(sem ele a regra nunca casa com nada)."
        )

    _compila_regex(regra.get("texto_regex"), f"{prefixo}.texto_regex", erros)
    _compila_regex(regra.get("ativo_regex"), f"{prefixo}.ativo_regex", erros)

    vizinhanca = regra.get("vizinhanca")
    if vizinhanca is not None:
        if not isinstance(vizinhanca, dict):
            erros.append(f"{prefixo}: 'vizinhanca' precisa ser um objeto ({{regex, janela, mesma_pagina}}).")
        else:
            _compila_regex(vizinhanca.get("regex"), f"{prefixo}.vizinhanca.regex", erros)
            janela = vizinhanca.get("janela")
            if janela is not None and not isinstance(janela, (int, float)):
                erros.append(f"{prefixo}: 'vizinhanca.janela' precisa ser numérico.")

    _valida_campos_lista(regra, prefixo, erros)


def _valida_regra_classificacao(regra: dict, idx: int, erros: list):
    prefixo = f"Regra de Classificação #{idx + 1}"

    ordem = regra.get("ordem")
    if not isinstance(ordem, (int, float)) or isinstance(ordem, bool):
        erros.append(f"{prefixo}: campo 'ordem' é obrigatório e precisa ser numérico.")

    _compila_regex(regra.get("ativo_regex"), f"{prefixo}.ativo_regex", erros)
    _compila_regex(regra.get("texto_regex"), f"{prefixo}.texto_regex", erros)

    operacao_em = regra.get("operacao_em")
    if operacao_em is not None and not isinstance(operacao_em, list):
        erros.append(f"{prefixo}: 'operacao_em' precisa ser uma lista.")

    _valida_campos_lista(regra, prefixo, erros)


LIMITE_REGRAS_ARQUIVO = 1000   # TASK-061: regras por importação
LIMITE_REGEX_ARQUIVO = 500     # TASK-061: tamanho de cada regex importada (contra regex catastrófica que travaria a tela)


def validar_importacao(tabela: str, regras) -> list:
    """Erros (vazia = válida) de uma lista de regras que veio de ARQUIVO: tudo o que `validar_regras` confere + limites de quantidade e de
    tamanho de regex. Não grava nada (a gravação continua sendo o Salvar de sempre)."""
    if not isinstance(regras, list):
        return ["O arquivo precisa conter uma lista de regras."]
    if len(regras) > LIMITE_REGRAS_ARQUIVO:
        return [f"Regras demais no arquivo ({len(regras)}; máx. {LIMITE_REGRAS_ARQUIVO})."]
    erros = []
    for i, r in enumerate(regras):
        if not isinstance(r, dict):
            continue                      # validar_regras acusa abaixo
        alvos = [("texto_regex", r.get("texto_regex")), ("ativo_regex", r.get("ativo_regex"))]
        if isinstance(r.get("vizinhanca"), dict):
            alvos.append(("vizinhanca.regex", r["vizinhanca"].get("regex")))
        for campo, valor in alvos:
            if isinstance(valor, str) and len(valor) > LIMITE_REGEX_ARQUIVO:
                erros.append(f"Regra #{i + 1}: '{campo}' passa de {LIMITE_REGEX_ARQUIVO} caracteres.")
    return erros + validar_regras(tabela, regras)


def validar_regras(tabela: str, regras) -> list:
    if not isinstance(regras, list):
        return ["O payload de regras precisa ser uma lista."]
    erros: list = []
    validador = _valida_regra_processamento if tabela == "processamento" else _valida_regra_classificacao
    for idx, regra in enumerate(regras):
        if not isinstance(regra, dict):
            erros.append(f"Regra #{idx + 1}: precisa ser um objeto.")
            continue
        validador(regra, idx, erros)
    return erros


# ═══════════════════════════════════════════════════════════════════════════
# LEITURA / ESCRITA
# ═══════════════════════════════════════════════════════════════════════════
def _get_regras(tabela: str, projeto_codigo: str):
    chave_local = f"regras_leitor_{tabela}"

    supabase = get_supabase()
    if supabase:
        try:
            res = supabase.table(chave_local).select("regras_json").eq("projeto_codigo", projeto_codigo).execute()
            if res.data:
                return json.loads(res.data[0]["regras_json"])
        except Exception as e:
            logger.warning(f"Falha ao buscar regras_leitor_{tabela} no Supabase: {e}")

    conn = get_connection()
    cursor = conn.cursor()
    cursor.execute(f"SELECT regras_json FROM {chave_local} WHERE projeto_codigo = ?", (projeto_codigo,))
    row = cursor.fetchone()
    conn.close()
    if row:
        try:
            return json.loads(row[0])
        except Exception:
            return []
    return []


def _save_regras(tabela: str, payload: RegrasLeitorPayload, usuario_email: str | None):
    erros = validar_regras(tabela, payload.regras)
    if erros:
        raise HTTPException(status_code=400, detail={"erros": erros})

    chave_local = f"regras_leitor_{tabela}"
    chave_historico = f"regras_leitor_{tabela}_historico"
    valor = json.dumps(payload.regras, ensure_ascii=False)

    conn = get_connection()
    cursor = conn.cursor()

    # Preserva a versão atual no histórico ANTES de sobrescrever (se já existir alguma).
    cursor.execute(f"SELECT regras_json FROM {chave_local} WHERE projeto_codigo = ?", (payload.projeto_codigo,))
    atual = cursor.fetchone()
    if atual:
        cursor.execute(
            f"INSERT INTO {chave_historico} (projeto_codigo, regras_json, criado_por) VALUES (?, ?, ?)",
            (payload.projeto_codigo, atual[0], usuario_email)
        )

    cursor.execute(
        f"INSERT OR REPLACE INTO {chave_local} (projeto_codigo, regras_json) VALUES (?, ?)",
        (payload.projeto_codigo, valor)
    )
    conn.commit()
    conn.close()

    supabase = get_supabase()
    if supabase:
        try:
            if atual:
                supabase.table(chave_historico).insert({
                    "projeto_codigo": payload.projeto_codigo,
                    "regras_json": atual[0],
                    "criado_por": usuario_email
                }).execute()
            supabase.table(chave_local).upsert({
                "projeto_codigo": payload.projeto_codigo,
                "regras_json": valor
            }).execute()
        except Exception as e:
            logger.warning(f"Erro ao salvar regras_leitor_{tabela} no Supabase: {e}")
            raise HTTPException(status_code=500, detail=f"Regras salvas localmente, mas falharam ao sincronizar com a nuvem (Supabase): {e}")

    return {"status": "success"}


def _listar_historico(tabela: str, projeto_codigo: str):
    chave_historico = f"regras_leitor_{tabela}_historico"

    supabase = get_supabase()
    if supabase:
        try:
            res = (
                supabase.table(chave_historico)
                .select("id, regras_json, criado_em, criado_por")
                .eq("projeto_codigo", projeto_codigo)
                .order("criado_em", desc=True)
                .execute()
            )
            if res.data:
                return [
                    {"id": r["id"], "criado_em": r["criado_em"], "criado_por": r.get("criado_por"),
                     "quantidade_regras": len(json.loads(r["regras_json"]))}
                    for r in res.data
                ]
        except Exception as e:
            logger.warning(f"Falha ao buscar histórico de regras_leitor_{tabela} no Supabase: {e}")

    conn = get_connection()
    cursor = conn.cursor()
    cursor.execute(
        f"SELECT id, regras_json, criado_em, criado_por FROM {chave_historico} "
        f"WHERE projeto_codigo = ? ORDER BY criado_em DESC, id DESC",
        (projeto_codigo,)
    )
    rows = cursor.fetchall()
    conn.close()
    resultado = []
    for row in rows:
        try:
            qtd = len(json.loads(row[1]))
        except Exception:
            qtd = 0
        resultado.append({"id": row[0], "criado_em": row[2], "criado_por": row[3], "quantidade_regras": qtd})
    return resultado


def _buscar_versao_historico(tabela: str, historico_id: int, projeto_codigo: str):
    chave_historico = f"regras_leitor_{tabela}_historico"

    supabase = get_supabase()
    if supabase:
        try:
            res = (
                supabase.table(chave_historico)
                .select("regras_json")
                .eq("id", historico_id)
                .eq("projeto_codigo", projeto_codigo)
                .execute()
            )
            if res.data:
                return json.loads(res.data[0]["regras_json"])
        except Exception as e:
            logger.warning(f"Falha ao buscar versão do histórico de regras_leitor_{tabela} no Supabase: {e}")

    conn = get_connection()
    cursor = conn.cursor()
    cursor.execute(
        f"SELECT regras_json FROM {chave_historico} WHERE id = ? AND projeto_codigo = ?",
        (historico_id, projeto_codigo)
    )
    row = cursor.fetchone()
    conn.close()
    if not row:
        return None
    return json.loads(row[0])


def _extrair_email(request: Request) -> str | None:
    user = getattr(request.state, "user", None)
    return user.get("email") if user else None


# ═══════════════════════════════════════════════════════════════════════════
# ROTAS
# ═══════════════════════════════════════════════════════════════════════════
@router.get("/processamento")
def get_processamento(projeto_codigo: str = "DEFAULT"):
    return {"regras": _get_regras("processamento", projeto_codigo)}


@router.post("/processamento/validar", dependencies=[Depends(require_role("admin"))])
def validar_importacao_processamento(payload: RegrasLeitorPayload):
    """Confere uma lista importada de arquivo SEM gravar (TASK-061)."""
    return {"erros": validar_importacao("processamento", payload.regras)[:30], "total": len(payload.regras)}


@router.post("/classificacao/validar", dependencies=[Depends(require_role("admin"))])
def validar_importacao_classificacao(payload: RegrasLeitorPayload):
    return {"erros": validar_importacao("classificacao", payload.regras)[:30], "total": len(payload.regras)}


@router.post("/processamento", dependencies=[Depends(require_role("admin"))])
def save_processamento(payload: RegrasLeitorPayload, request: Request):
    return _save_regras("processamento", payload, _extrair_email(request))


@router.get("/classificacao")
def get_classificacao(projeto_codigo: str = "DEFAULT"):
    return {"regras": _get_regras("classificacao", projeto_codigo)}


@router.post("/classificacao", dependencies=[Depends(require_role("admin"))])
def save_classificacao(payload: RegrasLeitorPayload, request: Request):
    return _save_regras("classificacao", payload, _extrair_email(request))


@router.get("/processamento/historico", dependencies=[Depends(require_role("admin"))])
def get_historico_processamento(projeto_codigo: str = "DEFAULT"):
    return {"historico": _listar_historico("processamento", projeto_codigo)}


@router.get("/classificacao/historico", dependencies=[Depends(require_role("admin"))])
def get_historico_classificacao(projeto_codigo: str = "DEFAULT"):
    return {"historico": _listar_historico("classificacao", projeto_codigo)}


@router.post("/processamento/reverter", dependencies=[Depends(require_role("admin"))])
def reverter_processamento(payload: ReverterPayload, request: Request):
    versao = _buscar_versao_historico("processamento", payload.historico_id, payload.projeto_codigo)
    if versao is None:
        raise HTTPException(status_code=404, detail="Versão de histórico não encontrada para este projeto.")
    novo_payload = RegrasLeitorPayload(projeto_codigo=payload.projeto_codigo, regras=versao)
    return _save_regras("processamento", novo_payload, _extrair_email(request))


@router.post("/classificacao/reverter", dependencies=[Depends(require_role("admin"))])
def reverter_classificacao(payload: ReverterPayload, request: Request):
    versao = _buscar_versao_historico("classificacao", payload.historico_id, payload.projeto_codigo)
    if versao is None:
        raise HTTPException(status_code=404, detail="Versão de histórico não encontrada para este projeto.")
    novo_payload = RegrasLeitorPayload(projeto_codigo=payload.projeto_codigo, regras=versao)
    return _save_regras("classificacao", novo_payload, _extrair_email(request))
