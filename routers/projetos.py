"""
routers/projetos.py — Rotas para gerenciamento de projetos base.
"""
from fastapi import APIRouter
from fastapi.responses import JSONResponse
from fastapi import Request
from database import get_row_connection, get_connection
from models import ProjetoRequest
from middleware.auth_middleware import get_current_user_from_state

router = APIRouter(prefix="/api/projetos", tags=["projetos"])


from services.supabase_client import get_supabase

@router.get("")
def get_projetos(request: Request):
    supabase = get_supabase()

    # Sempre lê o local primeiro (fonte de verdade local)
    conn = get_row_connection()
    cursor = conn.cursor()
    cursor.execute("SELECT nome, codigo FROM projetos ORDER BY nome")
    local_rows = {r["nome"]: dict(r) for r in cursor.fetchall()}
    conn.close()

    # Tenta buscar do Supabase e mescla
    if supabase:
        try:
            res = supabase.table("projetos").select("nome,codigo").order("nome").execute()
            if res.data:
                # Sincroniza projetos do Supabase para local (sem sobrescrever os locais)
                conn2 = get_connection()
                cur2 = conn2.cursor()
                for p in res.data:
                    nome_p = p.get("nome", "")
                    codigo_p = p.get("codigo", "")
                    if nome_p:
                        cur2.execute(
                            "INSERT OR IGNORE INTO projetos (nome, codigo) VALUES (?, ?)",
                            (nome_p, codigo_p)
                        )
                        # Adiciona ao mapa mesclado se não existia localmente
                        if nome_p not in local_rows:
                            local_rows[nome_p] = {"nome": nome_p, "codigo": codigo_p}
                conn2.commit()
                conn2.close()
        except Exception as e:
            pass

    # Retorna lista mesclada ordenada por nome
    merged = sorted(local_rows.values(), key=lambda x: x["nome"])
    return {"projetos": merged}



from fastapi import HTTPException

@router.post("")
def add_projeto(req: ProjetoRequest, request: Request):
    from fastapi import HTTPException
    from config import logger
    user = get_current_user_from_state(request)
    if user.get("role") != "admin":
        raise HTTPException(status_code=403, detail="Apenas administradores podem cadastrar novos projetos.")

    nome = req.nome.strip().upper()
    codigo = req.codigo.strip()

    supabase = get_supabase()
    if supabase:
        try:
            # Verifica se o nome já existe no Supabase
            existing = supabase.table("projetos").select("nome").eq("nome", nome).execute()
            if existing.data and len(existing.data) > 0:
                # Atualiza o código
                supabase.table("projetos").update({"codigo": codigo}).eq("nome", nome).execute()
            else:
                # Insere novo
                supabase.table("projetos").insert({"nome": nome, "codigo": codigo}).execute()
        except Exception as e:
            logger.warning(f"Erro ao salvar projeto no Supabase: {e}")

    try:
        conn = get_connection()
        cursor = conn.cursor()
        cursor.execute("INSERT OR REPLACE INTO projetos (nome, codigo) VALUES (?, ?)", (nome, codigo))
        conn.commit()
        conn.close()
        return {"status": "ok", "nome": nome, "codigo": codigo}
    except Exception as e:
        logger.error(f"Erro ao salvar projeto localmente: {e}")
        return JSONResponse(status_code=400, content={"error": str(e)})
