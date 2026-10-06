"""
services/contextos.py — Metadados de "contextos" dentro de um projeto (TASK-044).

Um contexto é só um nome, cadastrado por projeto (tabela `contextos`). Pensado para ser
compartilhado por mais de um mecanismo no futuro (hoje só Ajustes consome, via
services/repo_json.py::RepoJsonContexto para o overlay de cada contexto). Mesmo padrão de
nuvem-primeiro-com-fallback-local do resto do projeto.
"""
from config import logger
from database import get_connection
from services.supabase_client import get_supabase


def listar(projeto_codigo: str) -> list:
    """Nomes dos contextos cadastrados num projeto, em ordem alfabética."""
    supabase = get_supabase()
    if supabase:
        try:
            res = supabase.table("contextos").select("contexto").eq("projeto_codigo", projeto_codigo).execute()
            if res.data:
                return sorted({r["contexto"] for r in res.data})
        except Exception as e:
            logger.warning(f"Falha ao listar contextos no Supabase: {e}")
    conn = get_connection()
    rows = conn.execute(
        "SELECT contexto FROM contextos WHERE projeto_codigo = ? ORDER BY contexto", (projeto_codigo,)
    ).fetchall()
    conn.close()
    return [r[0] for r in rows]


def existe(projeto_codigo: str, contexto: str) -> bool:
    return contexto in listar(projeto_codigo)


def criar(projeto_codigo: str, contexto: str, email: str | None) -> None:
    """Idempotente: criar um contexto que já existe não faz nada (não apaga o overlay dele)."""
    conn = get_connection()
    conn.execute(
        "INSERT OR IGNORE INTO contextos (projeto_codigo, contexto, criado_por) VALUES (?, ?, ?)",
        (projeto_codigo, contexto, email),
    )
    conn.commit()
    conn.close()
    supabase = get_supabase()
    if supabase:
        try:
            supabase.table("contextos").upsert(
                {"projeto_codigo": projeto_codigo, "contexto": contexto, "criado_por": email}
            ).execute()
        except Exception as e:
            logger.warning(f"Erro ao criar contexto no Supabase: {e}")


def excluir(projeto_codigo: str, contexto: str) -> None:
    """Remove o metadado do contexto. Quem chama também deve excluir o overlay correspondente
    (services/repo_json.py::RepoJsonContexto.excluir) — não é feito aqui para manter esta função
    sem conhecimento de qual mecanismo guarda o quê."""
    conn = get_connection()
    conn.execute("DELETE FROM contextos WHERE projeto_codigo = ? AND contexto = ?", (projeto_codigo, contexto))
    conn.commit()
    conn.close()
    supabase = get_supabase()
    if supabase:
        try:
            supabase.table("contextos").delete().eq("projeto_codigo", projeto_codigo).eq("contexto", contexto).execute()
        except Exception as e:
            logger.warning(f"Erro ao excluir contexto no Supabase: {e}")
