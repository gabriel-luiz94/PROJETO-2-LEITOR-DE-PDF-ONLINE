"""
routers/admin.py — Rotas para gerenciamento local de usuários (Painel Admin).
"""
import hashlib
from fastapi import APIRouter, HTTPException, Depends
from pydantic import BaseModel
from database import get_connection, get_row_connection
from routers.auth import get_current_user
from services.supabase_client import get_supabase

router = APIRouter(prefix="/api/admin", tags=["admin"])


class UserCreate(BaseModel):
    email: str
    password: str
    is_admin: bool


class UserUpdate(BaseModel):
    is_admin: bool


def require_admin(token: str):
    user = get_current_user(token)
    if not user.get("is_admin"):
        raise HTTPException(status_code=403, detail="Acesso negado. Apenas administradores podem acessar este recurso.")
    return user


@router.get("/users")
def list_users(token: str):
    """Lista todos os usuários. Sincroniza da nuvem para o local se houver internet."""
    require_admin(token)
    supabase = get_supabase()
    conn = get_row_connection()
    cur = conn.cursor()
    
    if supabase:
        try:
            # Sincroniza Nuvem -> Local
            res = supabase.table("usuarios_nuvem").select("*").execute()
            cloud_users = res.data
            for cu in cloud_users:
                cur.execute('''
                    INSERT OR REPLACE INTO usuarios_locais (email, senha_hash, is_admin)
                    VALUES (?, ?, ?)
                ''', (cu["email"], cu["senha_hash"], int(cu["is_admin"])))
            conn.commit()
        except Exception:
            pass # Ignora falhas de sync
            
    cur.execute("SELECT id, email, is_admin FROM usuarios_locais")
    users = [dict(row) for row in cur.fetchall()]
    conn.close()
    return users


@router.post("/users")
def add_user(req: UserCreate, token: str):
    """Adiciona um novo usuário (Nuvem + Local)."""
    require_admin(token)
    
    if len(req.password) < 6:
        raise HTTPException(status_code=400, detail="A senha deve ter pelo menos 6 caracteres.")
        
    senha_hash = hashlib.sha256(req.password.encode()).hexdigest()
    supabase = get_supabase()
    
    if supabase:
        try:
            # Tenta inserir na nuvem primeiro
            supabase.table("usuarios_nuvem").insert({
                "email": req.email,
                "senha_hash": senha_hash,
                "is_admin": req.is_admin
            }).execute()
        except Exception as e:
            raise HTTPException(status_code=400, detail=f"Erro ao inserir na nuvem: o e-mail já pode existir.")
            
    conn = get_connection()
    cur = conn.cursor()
    try:
        cur.execute('''
            INSERT OR REPLACE INTO usuarios_locais (email, senha_hash, is_admin)
            VALUES (?, ?, ?)
        ''', (req.email, senha_hash, int(req.is_admin)))
        conn.commit()
        user_id = cur.lastrowid
    except Exception as e:
        conn.close()
        raise HTTPException(status_code=400, detail="Erro local: E-mail já está em uso.")
        
    conn.close()
    
    return {"id": user_id, "email": req.email, "is_admin": req.is_admin, "msg": "Usuário criado com sucesso!"}


@router.delete("/users/{user_id}")
def delete_user(user_id: int, token: str):
    """Remove um usuário."""
    require_admin(token)
    conn = get_connection()
    cur = conn.cursor()
    
    # Precisamos do e-mail para deletar na nuvem, pois o ID pode ser diferente
    cur.execute("SELECT email FROM usuarios_locais WHERE id = ?", (user_id,))
    row = cur.fetchone()
    if not row:
        conn.close()
        raise HTTPException(status_code=404, detail="Usuário não encontrado.")
    email = row[0]
    
    supabase = get_supabase()
    if supabase:
        try:
            supabase.table("usuarios_nuvem").delete().eq("email", email).execute()
        except Exception:
            pass
            
    cur.execute("DELETE FROM usuarios_locais WHERE id = ?", (user_id,))
    conn.commit()
    conn.close()
    return {"msg": "Usuário removido com sucesso."}


@router.put("/users/{user_id}/admin")
def toggle_admin(user_id: int, req: UserUpdate, token: str):
    """Ativa ou desativa os privilégios de admin."""
    require_admin(token)
    conn = get_connection()
    cur = conn.cursor()
    
    cur.execute("SELECT email FROM usuarios_locais WHERE id = ?", (user_id,))
    row = cur.fetchone()
    if not row:
        conn.close()
        raise HTTPException(status_code=404, detail="Usuário não encontrado.")
    email = row[0]
    
    supabase = get_supabase()
    if supabase:
        try:
            supabase.table("usuarios_nuvem").update({"is_admin": req.is_admin}).eq("email", email).execute()
        except Exception:
            pass
            
    cur.execute("UPDATE usuarios_locais SET is_admin = ? WHERE id = ?", (int(req.is_admin), user_id))
    conn.commit()
    conn.close()
class PasswordUpdate(BaseModel):
    password: str

@router.put("/users/{user_id}/password")
def update_password(user_id: int, req: PasswordUpdate, token: str):
    """Redefine a senha de um usuário."""
    require_admin(token)
    
    if len(req.password) < 6:
        raise HTTPException(status_code=400, detail="A senha deve ter pelo menos 6 caracteres.")
        
    senha_hash = hashlib.sha256(req.password.encode()).hexdigest()
    
    conn = get_connection()
    cur = conn.cursor()
    
    cur.execute("SELECT email FROM usuarios_locais WHERE id = ?", (user_id,))
    row = cur.fetchone()
    if not row:
        conn.close()
        raise HTTPException(status_code=404, detail="Usuário não encontrado.")
    email = row[0]
    
    supabase = get_supabase()
    if supabase:
        try:
            supabase.table("usuarios_nuvem").update({"senha_hash": senha_hash}).eq("email", email).execute()
        except Exception:
            pass
            
    cur.execute("UPDATE usuarios_locais SET senha_hash = ? WHERE id = ?", (senha_hash, user_id))
    conn.commit()
    conn.close()
    return {"msg": "Senha redefinida com sucesso."}
