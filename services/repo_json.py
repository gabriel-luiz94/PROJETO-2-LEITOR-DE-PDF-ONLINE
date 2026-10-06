"""
services/repo_json.py — Armazenamento "um JSON por projeto + histórico" com nuvem primeiro e fallback local (TASK-023).

Mesmo padrão de routers/validacao_regras.py (regras de domínio): lê do Supabase e, se falhar/vazio, do SQLite; grava no
SQLite e sincroniza o Supabase (a falha de sincronização é sinalizada ao chamador, com o dado já salvo localmente).
"""
import json

from config import logger
from database import get_connection
from services.supabase_client import get_supabase


class ErroSincronizacao(Exception):
    """Salvo localmente, mas a nuvem (Supabase) recusou/falhou."""


class RepoJson:
    def __init__(self, tabela: str, tabela_hist: str, coluna: str):
        self.tabela, self.hist, self.coluna = tabela, tabela_hist, coluna

    def buscar(self, projeto: str):
        supabase = get_supabase()
        if supabase:
            try:
                res = supabase.table(self.tabela).select(self.coluna).eq("projeto_codigo", projeto).execute()
                if res.data:
                    return json.loads(res.data[0][self.coluna])
            except Exception as e:
                logger.warning(f"Falha ao buscar {self.tabela} no Supabase: {e}")
        conn = get_connection()
        row = conn.execute(f"SELECT {self.coluna} FROM {self.tabela} WHERE projeto_codigo = ?", (projeto,)).fetchone()
        conn.close()
        return json.loads(row[0]) if row else None

    def gravar(self, projeto: str, objeto, email):
        valor = json.dumps(objeto, ensure_ascii=False)
        conn = get_connection()
        cur = conn.cursor()
        cur.execute(f"SELECT {self.coluna} FROM {self.tabela} WHERE projeto_codigo = ?", (projeto,))
        atual = cur.fetchone()
        if atual:  # preserva a versão atual ANTES de sobrescrever
            cur.execute(f"INSERT INTO {self.hist} (projeto_codigo, {self.coluna}, criado_por) VALUES (?, ?, ?)",
                        (projeto, atual[0], email))
        cur.execute(f"INSERT OR REPLACE INTO {self.tabela} (projeto_codigo, {self.coluna}) VALUES (?, ?)", (projeto, valor))
        conn.commit()
        conn.close()
        supabase = get_supabase()
        if supabase:
            try:
                if atual:
                    supabase.table(self.hist).insert({"projeto_codigo": projeto, self.coluna: atual[0], "criado_por": email}).execute()
                supabase.table(self.tabela).upsert({"projeto_codigo": projeto, self.coluna: valor}).execute()
            except Exception as e:
                logger.warning(f"Erro ao salvar {self.tabela} no Supabase: {e}")
                raise ErroSincronizacao(str(e))

    def historico(self, projeto: str) -> list:
        supabase = get_supabase()
        if supabase:
            try:
                res = (supabase.table(self.hist).select(f"id, {self.coluna}, criado_em, criado_por")
                       .eq("projeto_codigo", projeto).order("criado_em", desc=True).execute())
                if res.data:
                    return [{"id": r["id"], "criado_em": r["criado_em"], "criado_por": r.get("criado_por"),
                             "bruto": json.loads(r[self.coluna])} for r in res.data]
            except Exception as e:
                logger.warning(f"Falha ao buscar histórico de {self.tabela} no Supabase: {e}")
        conn = get_connection()
        rows = conn.execute(
            f"SELECT id, {self.coluna}, criado_em, criado_por FROM {self.hist} WHERE projeto_codigo = ? "
            "ORDER BY criado_em DESC, id DESC", (projeto,)).fetchall()
        conn.close()
        return [{"id": r[0], "criado_em": r[2], "criado_por": r[3], "bruto": json.loads(r[1])} for r in rows]


class RepoJsonContexto:
    """Mesmo padrão de `RepoJson` (nuvem primeiro, fallback local, histórico), mas chaveado por
    `(projeto_codigo, contexto)` em vez de só `projeto_codigo` (TASK-044)."""

    def __init__(self, tabela: str, tabela_hist: str, coluna: str):
        self.tabela, self.hist, self.coluna = tabela, tabela_hist, coluna

    def buscar(self, projeto: str, contexto: str):
        supabase = get_supabase()
        if supabase:
            try:
                res = (supabase.table(self.tabela).select(self.coluna)
                       .eq("projeto_codigo", projeto).eq("contexto", contexto).execute())
                if res.data:
                    return json.loads(res.data[0][self.coluna])
            except Exception as e:
                logger.warning(f"Falha ao buscar {self.tabela} no Supabase: {e}")
        conn = get_connection()
        row = conn.execute(
            f"SELECT {self.coluna} FROM {self.tabela} WHERE projeto_codigo = ? AND contexto = ?",
            (projeto, contexto)).fetchone()
        conn.close()
        return json.loads(row[0]) if row else None

    def gravar(self, projeto: str, contexto: str, objeto, email):
        valor = json.dumps(objeto, ensure_ascii=False)
        conn = get_connection()
        cur = conn.cursor()
        cur.execute(f"SELECT {self.coluna} FROM {self.tabela} WHERE projeto_codigo = ? AND contexto = ?",
                    (projeto, contexto))
        atual = cur.fetchone()
        if atual:
            cur.execute(
                f"INSERT INTO {self.hist} (projeto_codigo, contexto, {self.coluna}, criado_por) VALUES (?, ?, ?, ?)",
                (projeto, contexto, atual[0], email))
        cur.execute(
            f"INSERT OR REPLACE INTO {self.tabela} (projeto_codigo, contexto, {self.coluna}) VALUES (?, ?, ?)",
            (projeto, contexto, valor))
        conn.commit()
        conn.close()
        supabase = get_supabase()
        if supabase:
            try:
                if atual:
                    supabase.table(self.hist).insert(
                        {"projeto_codigo": projeto, "contexto": contexto, self.coluna: atual[0], "criado_por": email}
                    ).execute()
                supabase.table(self.tabela).upsert(
                    {"projeto_codigo": projeto, "contexto": contexto, self.coluna: valor}
                ).execute()
            except Exception as e:
                logger.warning(f"Erro ao salvar {self.tabela} no Supabase: {e}")
                raise ErroSincronizacao(str(e))

    def excluir(self, projeto: str, contexto: str):
        """Remove a linha do contexto (não toca no histórico — só auditoria, não é revertido daqui)."""
        conn = get_connection()
        conn.execute(f"DELETE FROM {self.tabela} WHERE projeto_codigo = ? AND contexto = ?", (projeto, contexto))
        conn.commit()
        conn.close()
        supabase = get_supabase()
        if supabase:
            try:
                supabase.table(self.tabela).delete().eq("projeto_codigo", projeto).eq("contexto", contexto).execute()
            except Exception as e:
                logger.warning(f"Erro ao excluir {self.tabela} no Supabase: {e}")
                raise ErroSincronizacao(str(e))

    def historico(self, projeto: str, contexto: str) -> list:
        supabase = get_supabase()
        if supabase:
            try:
                res = (supabase.table(self.hist).select(f"id, {self.coluna}, criado_em, criado_por")
                       .eq("projeto_codigo", projeto).eq("contexto", contexto).order("criado_em", desc=True).execute())
                if res.data:
                    return [{"id": r["id"], "criado_em": r["criado_em"], "criado_por": r.get("criado_por"),
                             "bruto": json.loads(r[self.coluna])} for r in res.data]
            except Exception as e:
                logger.warning(f"Falha ao buscar histórico de {self.tabela} no Supabase: {e}")
        conn = get_connection()
        rows = conn.execute(
            f"SELECT id, {self.coluna}, criado_em, criado_por FROM {self.hist} "
            "WHERE projeto_codigo = ? AND contexto = ? ORDER BY criado_em DESC, id DESC",
            (projeto, contexto)).fetchall()
        conn.close()
        return [{"id": r[0], "criado_em": r[2], "criado_por": r[3], "bruto": json.loads(r[1])} for r in rows]
