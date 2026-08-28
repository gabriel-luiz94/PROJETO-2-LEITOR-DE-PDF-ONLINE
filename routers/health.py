"""
routers/health.py — Rotas para diagnóstico e exportação de backup.
"""
import os
import sys
from fastapi import APIRouter
from fastapi.responses import JSONResponse
from database import get_row_connection
from config import DB_PATH, PROMPT_PATH, SEED_CSV_PATH, STATIC_DIR
from services.sync_service import sync_tabela_master

router = APIRouter(tags=["health"])


@router.get("/api/health")
def health_check():
    """Diagnóstico para uso na empresa: verifica DB, prompt e IA."""
    info = {}
    # DB
    try:
        conn = get_row_connection()
        cur = conn.cursor()
        cur.execute("SELECT COUNT(*) FROM tabela_orcamento")
        info["tabela_orcamento_count"] = cur.fetchone()[0]
        cur.execute("SELECT COUNT(*) FROM configuracoes")
        info["configuracoes_count"] = cur.fetchone()[0]
        conn.close()
        info["db_path"] = DB_PATH
        info["db_ok"] = True
    except Exception as e:
        info["db_ok"] = False
        info["db_error"] = str(e)
        
    # Prompt
    info["prompt_path"] = PROMPT_PATH
    info["prompt_exists"] = os.path.exists(PROMPT_PATH)
    info["seed_path"] = SEED_CSV_PATH
    info["seed_exists"] = os.path.exists(SEED_CSV_PATH)
    
    # ENV
    info["env_gemini_key"] = bool(os.getenv("GEMINI_API_KEY") or os.getenv("GOOGLE_API_KEY"))
    info["frozen"] = getattr(sys, "frozen", False)
    info["static_dir"] = STATIC_DIR
    return info


@router.get("/api/backup/export")
def backup_export():
    """Exporta tabela_orcamento + regras + projetos como JSON para backup portátil."""
    try:
        conn = get_row_connection()
        cur = conn.cursor()
        cur.execute("SELECT * FROM tabela_orcamento")
        orcamento = [dict(r) for r in cur.fetchall()]
        cur.execute("SELECT * FROM regras")
        regras = [dict(r) for r in cur.fetchall()]
        cur.execute("SELECT * FROM projetos")
        projetos = [dict(r) for r in cur.fetchall()]
        conn.close()
        return {"tabela_orcamento": orcamento, "regras": regras, "projetos": projetos}
    except Exception as e:
        return JSONResponse(status_code=500, content={"error": str(e)})

@router.post("/api/health/sync-master")
def trigger_sync_master():
    """Baixa o banco mestre do Supabase para o cache local."""
    sucesso = sync_tabela_master()
    return {"status": "ok" if sucesso else "error"}
