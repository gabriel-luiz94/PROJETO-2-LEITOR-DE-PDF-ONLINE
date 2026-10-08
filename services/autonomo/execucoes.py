"""
services/autonomo/execucoes.py — Histórico das execuções do modo autônomo (TASK-031, fase B).

Tabela LOCAL `execucoes_autonomas` (SQLite), criada sob demanda aqui — `database.init_db` não foi alterado. É um registro operacional
da máquina que roda o autônomo; não é sincronizada com o Supabase (as OBRAS geradas continuam indo para o Supabase como as manuais).
Cada linha guarda o necessário para CONFIRMAR exclusões, REVERTER e REPROCESSAR: as tabelas originais (antes dos ajustes) e o diff.
"""
import json
import uuid
from datetime import datetime

import database

STATUS = ("processando", "aguardando_confirmacao", "ok", "com_pendencias", "erro", "revertida")
_JSON = ("relatorio", "originais", "diff", "decisoes", "itens_origem")


def _agora():
    return datetime.now().isoformat(timespec="seconds")


def garantir_tabela(conn):
    conn.execute("""
        CREATE TABLE IF NOT EXISTS execucoes_autonomas (
            id TEXT PRIMARY KEY, arquivo TEXT, arquivo_hash TEXT, projeto_codigo TEXT, user_id TEXT, status TEXT,
            obra_id TEXT, pasta_saida TEXT, mensagem TEXT, relatorio_json TEXT, originais_json TEXT, diff_json TEXT,
            decisoes_json TEXT, criado_em TEXT, atualizado_em TEXT)""")
    try:
        conn.execute("ALTER TABLE execucoes_autonomas ADD COLUMN arquivo_caminho TEXT")   # onde o original está agora (para reprocessar)
    except Exception:  # noqa: BLE001 — coluna já existe
        pass
    try:
        conn.execute("ALTER TABLE execucoes_autonomas ADD COLUMN itens_origem_json TEXT")   # TASK-059: de onde veio cada linha (aprendizado de regras do leitor)
    except Exception:  # noqa: BLE001 — coluna já existe
        pass
    conn.execute("CREATE INDEX IF NOT EXISTS idx_exec_aut_hash ON execucoes_autonomas (arquivo_hash, projeto_codigo)")


def _conectar():
    conn = database.get_row_connection()
    garantir_tabela(conn)
    return conn


def _linha(r) -> dict:
    d = dict(r)
    for k in _JSON:
        bruto = d.pop(f"{k}_json", None)
        d[k] = json.loads(bruto) if bruto else None
    return d


def criar(arquivo: str, arquivo_hash: str, projeto_codigo: str, user_id: str, status: str = "processando") -> str:
    exec_id = uuid.uuid4().hex[:12]
    conn = _conectar()
    try:
        conn.execute("INSERT INTO execucoes_autonomas (id, arquivo, arquivo_hash, projeto_codigo, user_id, status, criado_em, atualizado_em) "
                     "VALUES (?, ?, ?, ?, ?, ?, ?, ?)", (exec_id, arquivo, arquivo_hash, projeto_codigo, user_id, status, _agora(), _agora()))
        conn.commit()
    finally:
        conn.close()
    return exec_id


def atualizar(exec_id: str, **campos) -> None:
    sets, valores = [], []
    for k, v in campos.items():
        if k in _JSON:
            k, v = f"{k}_json", json.dumps(v, ensure_ascii=False)
        sets.append(f"{k} = ?")
        valores.append(v)
    sets.append("atualizado_em = ?")
    valores.append(_agora())
    conn = _conectar()
    try:
        conn.execute(f"UPDATE execucoes_autonomas SET {', '.join(sets)} WHERE id = ?", (*valores, exec_id))
        conn.commit()
    finally:
        conn.close()


def buscar(exec_id: str):
    conn = _conectar()
    try:
        r = conn.execute("SELECT * FROM execucoes_autonomas WHERE id = ?", (exec_id,)).fetchone()
    finally:
        conn.close()
    return _linha(r) if r else None


def buscar_por_hash(arquivo_hash: str, projeto_codigo: str):
    """A execução mais recente (não-erro) do mesmo conteúdo no mesmo projeto — para não processar duas vezes."""
    conn = _conectar()
    try:
        r = conn.execute("SELECT * FROM execucoes_autonomas WHERE arquivo_hash = ? AND projeto_codigo = ? AND status NOT IN ('erro', 'revertida') "
                         "ORDER BY criado_em DESC LIMIT 1", (arquivo_hash, projeto_codigo)).fetchone()
    finally:
        conn.close()
    return _linha(r) if r else None


def listar(status: str = None, projeto: str = None, limite: int = 100) -> list:
    """Resumo das execuções (sem os JSONs grandes), mais recentes primeiro. Filtros por status e/ou projeto (TASK-035)."""
    conn = _conectar()
    try:
        sql = ("SELECT id, arquivo, arquivo_caminho, projeto_codigo, user_id, status, obra_id, pasta_saida, mensagem, criado_em, atualizado_em "
               "FROM execucoes_autonomas")
        condicoes, args = [], []
        if status:
            condicoes.append("status = ?")
            args.append(status)
        if projeto:
            condicoes.append("projeto_codigo = ?")
            args.append(projeto)
        if condicoes:
            sql += " WHERE " + " AND ".join(condicoes)
        rows = conn.execute(sql + " ORDER BY criado_em DESC, rowid DESC LIMIT ?", (*args, limite)).fetchall()
    finally:
        conn.close()
    return [dict(r) for r in rows]
