"""
services/autonomo/aprendizado_repo.py — Armazenamento LOCAL do aprendizado (TASK-059), no mesmo estilo de `execucoes.py`.

Tabelas SQLite criadas sob demanda (`database.init_db` não foi alterado; nada vai ao Supabase — é registro operacional da máquina):
  aprendizado_correcoes  uma linha por obra do autônomo que o usuário corrigiu e salvou (só as DIFERENÇAS, ver `aprendizado.comparar`)
  aprendizado_propostas  propostas de ajuste (estatística ou IA) com status pendente / aprovada / recusada
Tudo é por usuário e por projeto ("por enquanto só com o meu trabalho"). Não registrar conteúdo de planilhas em log.
"""
import json
import logging
from datetime import datetime

import database
from services.autonomo import aprendizado, execucoes

logger = logging.getLogger(__name__)
STATUS_PROPOSTA = ("pendente", "aprovada", "recusada")


def _agora():
    return datetime.now().isoformat(timespec="seconds")


def _conectar():
    conn = database.get_row_connection()
    conn.execute("""CREATE TABLE IF NOT EXISTS aprendizado_correcoes (
        execucao_id TEXT NOT NULL, user_id TEXT NOT NULL, projeto_codigo TEXT NOT NULL, arquivo TEXT, obra_salva TEXT,
        eventos_json TEXT, base_json TEXT, linhas_json TEXT, criado_em TEXT, atualizado_em TEXT,
        PRIMARY KEY (execucao_id, user_id))""")
    conn.execute("""CREATE TABLE IF NOT EXISTS aprendizado_propostas (
        id INTEGER PRIMARY KEY AUTOINCREMENT, user_id TEXT NOT NULL, projeto_codigo TEXT NOT NULL, chave TEXT NOT NULL, tipo TEXT, tabela TEXT,
        origem TEXT, status TEXT NOT NULL DEFAULT 'pendente', descricao TEXT, ocorrencias INTEGER, execucoes INTEGER, consistencia REAL,
        receita_json TEXT, criado_em TEXT, decidido_em TEXT, UNIQUE (user_id, projeto_codigo, chave))""")
    return conn


# ── correções ───────────────────────────────────────────────────────────────
def _tabelas_da_obra(dados_json):
    try:
        snap = json.loads(dados_json) if isinstance(dados_json, str) else (dados_json or {})
    except (TypeError, ValueError):
        return None
    return {"cabos": snap.get("cabos"), "outros": snap.get("outros")} if isinstance(snap, dict) else None


def _obra_local(obra_id):
    conn = database.get_connection()
    try:
        r = conn.execute("SELECT dados_json FROM obras WHERE id = ?", (obra_id,)).fetchone()
    finally:
        conn.close()
    return r[0] if r else None


def registrar_correcao(user_id, execucao_id, obra_salva_id, dados_salvos_json):
    """Compara a obra que o autônomo gerou com a que o usuário salvou e guarda as diferenças (uma por execução e usuário; salvar de novo
    substitui). Devolve o resumo ou None quando não se aplica (execução alheia/inexistente, obra do autônomo revertida ou ausente)."""
    ex = execucoes.buscar(execucao_id)
    if not ex or ex.get("user_id") != user_id or ex.get("status") not in ("ok", "com_pendencias") or not ex.get("obra_id"):
        return None
    agente = _tabelas_da_obra(_obra_local(ex["obra_id"]))
    usuario = _tabelas_da_obra(dados_salvos_json)
    if not agente or not usuario:
        return None
    r = aprendizado.comparar(agente, usuario)
    conn = _conectar()
    try:
        conn.execute("""INSERT INTO aprendizado_correcoes (execucao_id, user_id, projeto_codigo, arquivo, obra_salva, eventos_json, base_json, linhas_json, criado_em, atualizado_em)
            VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?) ON CONFLICT(execucao_id, user_id) DO UPDATE SET
            obra_salva=excluded.obra_salva, eventos_json=excluded.eventos_json, base_json=excluded.base_json, linhas_json=excluded.linhas_json, atualizado_em=excluded.atualizado_em""",
                     (execucao_id, user_id, ex["projeto_codigo"], ex.get("arquivo"), obra_salva_id, json.dumps(r["eventos"], ensure_ascii=False),
                      json.dumps(r["base"], ensure_ascii=False), json.dumps(r["linhas"]), _agora(), _agora()))
        conn.commit()
    finally:
        conn.close()
    return {"projeto": ex["projeto_codigo"], "eventos": len(r["eventos"])}


def listar_registros(user_id, projeto) -> list:
    conn = _conectar()
    try:
        rows = conn.execute("SELECT * FROM aprendizado_correcoes WHERE user_id = ? AND projeto_codigo = ? ORDER BY atualizado_em DESC", (user_id, projeto)).fetchall()
    finally:
        conn.close()
    return [{"execucao_id": r["execucao_id"], "arquivo": r["arquivo"], "obra_salva": r["obra_salva"], "atualizado_em": r["atualizado_em"],
             "eventos": json.loads(r["eventos_json"] or "[]"), "base": json.loads(r["base_json"] or "{}"), "linhas": json.loads(r["linhas_json"] or "{}")}
            for r in rows]


def apagar_historico(user_id, projeto) -> dict:
    """Apaga as correções registradas E as propostas ainda pendentes/recusadas do usuário no projeto (as aprovadas já viraram ajustes)."""
    conn = _conectar()
    try:
        a = conn.execute("DELETE FROM aprendizado_correcoes WHERE user_id = ? AND projeto_codigo = ?", (user_id, projeto)).rowcount
        b = conn.execute("DELETE FROM aprendizado_propostas WHERE user_id = ? AND projeto_codigo = ? AND status != 'aprovada'", (user_id, projeto)).rowcount
        conn.commit()
    finally:
        conn.close()
    return {"correcoes": a, "propostas": b}


# ── propostas ───────────────────────────────────────────────────────────────
def _linha_proposta(r) -> dict:
    d = dict(r)
    d["receita"] = json.loads(d.pop("receita_json") or "null")
    return d


def listar_propostas(user_id, projeto, status=None) -> list:
    conn = _conectar()
    try:
        sql, args = "SELECT * FROM aprendizado_propostas WHERE user_id = ? AND projeto_codigo = ?", [user_id, projeto]
        if status:
            sql += " AND status = ?"
            args.append(status)
        rows = conn.execute(sql + " ORDER BY CASE status WHEN 'pendente' THEN 0 ELSE 1 END, ocorrencias DESC, id DESC", args).fetchall()
    finally:
        conn.close()
    return [_linha_proposta(r) for r in rows]


def buscar_proposta(user_id, proposta_id):
    conn = _conectar()
    try:
        r = conn.execute("SELECT * FROM aprendizado_propostas WHERE id = ? AND user_id = ?", (proposta_id, user_id)).fetchone()
    finally:
        conn.close()
    return _linha_proposta(r) if r else None


def guardar_propostas(user_id, projeto, padroes: list, origem="estatistica") -> int:
    """Guarda como PENDENTES os padrões ainda desconhecidos (chave nova). Recusada/aprovada/pendente já existente não é recriada nem reaberta."""
    novas = 0
    conn = _conectar()
    try:
        for p in padroes:
            cur = conn.execute("""INSERT OR IGNORE INTO aprendizado_propostas (user_id, projeto_codigo, chave, tipo, tabela, origem, status, descricao, ocorrencias, execucoes, consistencia, receita_json, criado_em)
                VALUES (?, ?, ?, ?, ?, ?, 'pendente', ?, ?, ?, ?, ?, ?)""",
                               (user_id, projeto, p["chave"], p["tipo"], p["tabela"], origem, p["descricao"], p["ocorrencias"], p["execucoes"], p["consistencia"],
                                json.dumps(p["receita"], ensure_ascii=False) if p.get("receita") else None, _agora()))
            novas += cur.rowcount
        conn.commit()
    finally:
        conn.close()
    return novas


def decidir_proposta(user_id, proposta_id, status) -> None:
    conn = _conectar()
    try:
        conn.execute("UPDATE aprendizado_propostas SET status = ?, decidido_em = ? WHERE id = ? AND user_id = ?", (status, _agora(), proposta_id, user_id))
        conn.commit()
    finally:
        conn.close()


def reanalisar(user_id, projeto, grupos=None) -> dict:
    """Recalcula os padrões das correções do usuário no projeto e guarda as propostas novas. Barato: só agregação de contadores."""
    registros = listar_registros(user_id, projeto)
    padroes = aprendizado.detectar_padroes(registros, grupos)
    return {"obras": len(registros), "padroes": len(padroes), "novas": guardar_propostas(user_id, projeto, padroes)}


def resumo(user_id, projeto) -> dict:
    regs = listar_registros(user_id, projeto)
    props = listar_propostas(user_id, projeto)
    cont = {s: sum(1 for p in props if p["status"] == s) for s in STATUS_PROPOSTA}
    return {"obras_corrigidas": len(regs), "obras_sem_correcao": sum(1 for r in regs if not r["eventos"]),
            "correcoes_total": sum(len(r["eventos"]) for r in regs), "propostas": cont,
            "ia_minimo_obras": aprendizado.IA_MINIMO_OBRAS, "ia_disponivel": len(regs) >= aprendizado.IA_MINIMO_OBRAS,
            "limiares": {"ocorrencias": aprendizado.MINIMO_OCORRENCIAS, "obras": aprendizado.MINIMO_EXECUCOES, "consistencia": aprendizado.CONSISTENCIA_MINIMA}}
