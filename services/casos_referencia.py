"""
services/casos_referencia.py — Casos de referência por projeto (TASK-178, ADR-007): pedidos com o resultado esperado, para conferir a base antes de gravar.

Caso: {"nome", "pedido": {"cabos": [...], "outros": [...]} (o mesmo payload do cálculo), "obra_cabecalho": {municipio, bairro, utd, utep, contexto}, "modo": "calculo",
       "esperado": [{"operacao", "mdo", "codigo", "qtd"}], "tolerancia": 0.005, "origem": "programa" | "ICP" | "manual", "nota", "historico": [{quando, quem, nota}]}
`conferir` roda os casos em lote, com a base carregada UMA vez (não muda o cálculo): cada caso vira igual / mudou (com o antes → depois código a código) / erro.
"""
import json
import time
from datetime import datetime, timezone

import database
from services import orcamento_calc as calc

ORIGENS = ("programa", "ICP", "manual")
TOLERANCIA_PADRAO = 0.005
LIMITE_CASOS = 2000


class ErroCaso(ValueError):
    """Caso recusado; `args[0]` = mensagem em português."""


def _trim(t) -> str:
    return str(t if t is not None else "").strip()


def _num(v, padrao=0.0) -> float:
    try:
        return round(float(str(v).replace(",", ".")), 6) if not isinstance(v, bool) else padrao
    except (TypeError, ValueError):
        return padrao


def chave(item) -> str:
    return f"{_trim(item.get('operacao')).upper() or 'I'}|{_trim(item.get('mdo')).upper()}|{_trim(item.get('codigo'))}"


def esperado_do_resultado(resultado) -> list:
    """Resultado do cálculo → esperado (soma por operação · tipo de mão de obra · código), ordenado."""
    soma = {}
    for r in resultado or []:
        k = chave(r)
        soma[k] = soma.get(k, 0.0) + _num(r.get("total"))
    return [{"operacao": k.split("|")[0], "mdo": k.split("|")[1], "codigo": k.split("|", 2)[2], "qtd": round(v, 6)} for k, v in sorted(soma.items())]


def normalizar_caso(c) -> dict:
    if not isinstance(c, dict):
        raise ErroCaso("O caso precisa ser um objeto.")
    nome = _trim(c.get("nome"))[:120]
    if not nome:
        raise ErroCaso("Dê um nome ao caso de referência.")
    pedido = c.get("pedido") if isinstance(c.get("pedido"), dict) else {}
    cabos, outros = pedido.get("cabos") or [], pedido.get("outros") or []
    if not isinstance(cabos, list) or not isinstance(outros, list) or not all(isinstance(x, dict) for x in cabos + outros):
        raise ErroCaso("O pedido do caso precisa ter listas 'cabos' e 'outros' de linhas.")
    if not cabos and not outros:
        raise ErroCaso("O pedido do caso está vazio.")
    esperado = []
    for x in c.get("esperado") if isinstance(c.get("esperado"), list) else []:
        if isinstance(x, dict) and _trim(x.get("codigo")):
            esperado.append({"operacao": _trim(x.get("operacao")).upper() or "I", "mdo": _trim(x.get("mdo")).upper(), "codigo": _trim(x["codigo"]), "qtd": _num(x.get("qtd"))})
    tol = _num(c.get("tolerancia"), TOLERANCIA_PADRAO)
    origem = _trim(c.get("origem")) if _trim(c.get("origem")) in ORIGENS else "manual"
    obra = c.get("obra_cabecalho") if isinstance(c.get("obra_cabecalho"), dict) else {}
    hist = [h for h in (c.get("historico") if isinstance(c.get("historico"), list) else []) if isinstance(h, dict)][-50:]
    return {"nome": nome, "pedido": {"cabos": cabos, "outros": outros}, "obra_cabecalho": {k: _trim(v)[:200] for k, v in obra.items() if isinstance(k, str)},
            "modo": "calculo", "esperado": sorted(esperado, key=lambda x: chave(x)), "tolerancia": tol if tol >= 0 else TOLERANCIA_PADRAO, "origem": origem, "nota": _trim(c.get("nota"))[:300], "historico": hist}


def calcular(caso, rows, projeto_nome, convencao_postes=None, perfil=None) -> list:
    """O que o cálculo dá hoje para o caso (esperado-like). Levanta o erro do cálculo, se houver."""
    r = calc.processar_calculo(caso["pedido"]["cabos"], caso["pedido"]["outros"], projeto_nome, rows, obra=caso.get("obra_cabecalho") or None, convencao_postes=convencao_postes, perfil_exportacao=perfil)
    return esperado_do_resultado(r.get("resultado"))


def comparar(esperado, obtido, tolerancia=TOLERANCIA_PADRAO) -> list:
    """Diferenças [{codigo, operacao, mdo, antes, depois, tipo}] (antes = esperado; depois = obtido); lista vazia = igual."""
    a = {chave(x): x for x in esperado}
    b = {chave(x): x for x in obtido}
    dif = []
    for k in sorted(set(a) | set(b)):
        x, y = a.get(k), b.get(k)
        antes, depois = (x["qtd"] if x else None), (y["qtd"] if y else None)
        ref = x or y
        if x is None:
            tipo = "novo"
        elif y is None:
            tipo = "sumiu"
        elif abs(antes - depois) > tolerancia:
            tipo = "qtd"
        else:
            continue
        dif.append({"codigo": ref["codigo"], "operacao": ref["operacao"], "mdo": ref["mdo"], "antes": antes, "depois": depois, "tipo": tipo})
    return dif


def conferir(casos, rows, projeto_nome, convencao_postes=None, perfil=None) -> dict:
    """Roda `casos` [{id?, nome, ...}] com a base `rows` carregada uma vez. {"casos": [{id, nome, situacao: igual|mudou|erro, diferencas, erro?}], "resumo": {...}, "segundos"}."""
    t0 = time.perf_counter()
    saida = []
    for c in casos:
        item = {"id": c.get("id"), "nome": c.get("nome"), "situacao": "igual", "diferencas": []}
        try:
            obtido = calcular(c, rows, projeto_nome, convencao_postes, perfil)
            item["diferencas"] = comparar(c.get("esperado") or [], obtido, c.get("tolerancia", TOLERANCIA_PADRAO))
            if item["diferencas"]:
                item["situacao"] = "mudou"
        except Exception as e:  # noqa: BLE001
            item["situacao"], item["erro"] = "erro", str(e)[:300]
        saida.append(item)
    resumo = {s: sum(1 for i in saida if i["situacao"] == s) for s in ("igual", "mudou", "erro")}
    return {"casos": saida, "resumo": resumo, "total": len(saida), "segundos": round(time.perf_counter() - t0, 3)}


def relatorio_markdown(resultado) -> str:
    """Relatório das mudanças (no estilo do CON_EMB_mudancas.md): cada caso que mudou, código a código (antes → depois)."""
    linhas = [f"# Conferência da base — {resultado['resumo'].get('mudou', 0)} caso(s) mudaram, {resultado['resumo'].get('erro', 0)} com erro, {resultado['resumo'].get('igual', 0)} iguais", ""]
    for c in resultado["casos"]:
        if c["situacao"] == "igual":
            continue
        linhas.append(f"## {c['nome']} — {c['situacao']}")
        if c.get("erro"):
            linhas.append(f"- erro: {c['erro']}")
        for d in c["diferencas"]:
            linhas.append(f"- {d['codigo']} ({d['operacao']} · {d['mdo'] or '-'}): {'—' if d['antes'] is None else d['antes']} → {'—' if d['depois'] is None else d['depois']} ({d['tipo']})")
        linhas.append("")
    return "\n".join(linhas)


# ── armazenamento ─────────────────────────────────────────────────────────────────────────────
def _tabela(conn):
    conn.execute("CREATE TABLE IF NOT EXISTS casos_referencia (id INTEGER PRIMARY KEY AUTOINCREMENT, projeto_codigo TEXT NOT NULL, nome TEXT NOT NULL, caso_json TEXT NOT NULL, "
                 "criado_em TEXT DEFAULT (datetime('now')), criado_por TEXT)")


def listar(projeto_codigo) -> list:
    conn = database.get_connection()
    try:
        _tabela(conn)
        rows = conn.execute("SELECT id, caso_json FROM casos_referencia WHERE projeto_codigo = ? ORDER BY id", (projeto_codigo,)).fetchall()
    finally:
        conn.close()
    return [dict(json.loads(r[1]), id=r[0]) for r in rows]


def gravar(projeto_codigo, caso, usuario=None) -> dict:
    c = normalizar_caso(caso)
    conn = database.get_connection()
    try:
        _tabela(conn)
        if conn.execute("SELECT COUNT(*) FROM casos_referencia WHERE projeto_codigo = ?", (projeto_codigo,)).fetchone()[0] >= LIMITE_CASOS:
            raise ErroCaso(f"Limite de {LIMITE_CASOS} casos por projeto.")
        cur = conn.execute("INSERT INTO casos_referencia (projeto_codigo, nome, caso_json, criado_por) VALUES (?, ?, ?, ?)", (projeto_codigo, c["nome"], json.dumps(c, ensure_ascii=False), usuario))
        conn.commit()
        novo_id = cur.lastrowid
    finally:
        conn.close()
    _espelhar(projeto_codigo, novo_id, c, usuario)
    return dict(c, id=novo_id)


def atualizar(projeto_codigo, caso_id, caso) -> dict:
    c = normalizar_caso(caso)
    conn = database.get_connection()
    try:
        _tabela(conn)
        cur = conn.execute("UPDATE casos_referencia SET nome = ?, caso_json = ? WHERE id = ? AND projeto_codigo = ?", (c["nome"], json.dumps(c, ensure_ascii=False), caso_id, projeto_codigo))
        conn.commit()
        if not cur.rowcount:
            raise ErroCaso("Caso não encontrado neste projeto.")
    finally:
        conn.close()
    _espelhar(projeto_codigo, caso_id, c, None)
    return dict(c, id=caso_id)


def excluir(projeto_codigo, caso_id) -> bool:
    conn = database.get_connection()
    try:
        _tabela(conn)
        cur = conn.execute("DELETE FROM casos_referencia WHERE id = ? AND projeto_codigo = ?", (caso_id, projeto_codigo))
        conn.commit()
        return bool(cur.rowcount)
    finally:
        conn.close()


def aceitar_mudanca(projeto_codigo, caso, obtido, usuario, nota) -> dict:
    """"Mudança esperada": o resultado de hoje vira o `esperado` do caso, com a nota de quem aceitou e por quê."""
    nota = _trim(nota)
    if not nota:
        raise ErroCaso("Diga por que a mudança é esperada (a nota fica no histórico do caso).")
    novo = dict(caso, esperado=obtido)
    novo["historico"] = (caso.get("historico") or []) + [{"quando": datetime.now(timezone.utc).isoformat(timespec="seconds"), "quem": usuario or "", "nota": nota[:300]}]
    return atualizar(projeto_codigo, caso["id"], novo)


def _espelhar(projeto_codigo, caso_id, caso, usuario):
    """Cópia na nuvem (Supabase `casos_referencia`), se houver; falha não perde a gravação local."""
    try:
        from services.supabase_client import get_supabase
        sb = get_supabase()
        if sb:
            sb.table("casos_referencia").upsert({"id": caso_id, "projeto_codigo": projeto_codigo, "nome": caso["nome"], "caso_json": json.dumps(caso, ensure_ascii=False), "criado_por": usuario}).execute()
    except Exception as e:  # noqa: BLE001
        from config import logger
        logger.warning(f"Caso de referência gravado só neste computador (Supabase indisponível): {e}")
