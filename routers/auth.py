"""
routers/auth.py — Rotas para Autenticação usando Supabase.
"""
from fastapi import APIRouter, HTTPException, Depends
from pydantic import BaseModel
from services.supabase_client import get_supabase
from database import get_connection, get_row_connection
from config import logger
from typing import Optional

router = APIRouter(prefix="/api/auth", tags=["auth"])


class LoginRequest(BaseModel):
    email: str
    password: str
    lembrar: bool = False


class AuthResponse(BaseModel):
    access_token: str
    user_id: str
    email: str
    is_admin: bool


@router.post("/login", response_model=AuthResponse)
def login(req: LoginRequest):
    import hashlib
    import uuid
    
    # 1. Primeiro tentamos validar no banco de usuários locais.
    conn = get_row_connection()
    cur = conn.cursor()
    cur.execute("SELECT * FROM usuarios_locais WHERE email = ?", (req.email,))
    user_local = cur.fetchone()
    conn.close()

    if user_local:
        senha_hash_entrada = hashlib.sha256(req.password.encode()).hexdigest()
        if senha_hash_entrada == user_local["senha_hash"]:
            # Login local bem sucedido
            user_id = str(user_local["id"])
            email = user_local["email"]
            is_admin = bool(user_local["is_admin"])
            
            # Gera um token falso para sessao local offline
            access_token = f"local_token_{uuid.uuid4().hex}"
            
            conn = get_connection()
            cur = conn.cursor()
            cur.execute('''
                INSERT OR REPLACE INTO sessoes (user_id, email, access_token, refresh_token, is_admin, ultimo_login)
                VALUES (?, ?, ?, ?, ?, datetime('now'))
            ''', (user_id, email, access_token, "", int(is_admin)))
            conn.commit()
            conn.close()
            
            return {
                "access_token": access_token,
                "user_id": user_id,
                "email": email,
                "is_admin": is_admin
            }

    # 2. Se não encontrou usuário local, tenta o Supabase (online) na tabela `usuarios_nuvem`
    supabase = get_supabase()
    if not supabase:
        # Se não há conexão e também não era usuário local
        raise HTTPException(status_code=503, detail="Sem conexão com servidor de login e usuário local não encontrado.")

    try:
        res = supabase.table("usuarios_nuvem").select("*").eq("email", req.email).execute()
        if not res.data:
            raise HTTPException(status_code=401, detail="Credenciais inválidas ou usuário não encontrado.")
            
        user_cloud = res.data[0]
        senha_hash_entrada = hashlib.sha256(req.password.encode()).hexdigest()
        
        if senha_hash_entrada != user_cloud.get("senha_hash"):
            raise HTTPException(status_code=401, detail="Credenciais inválidas.")
            
        user_id = str(user_cloud.get("id"))
        email = user_cloud.get("email")
        is_admin = bool(user_cloud.get("is_admin"))
        
        # Gera token temporário
        access_token = f"cloud_token_{uuid.uuid4().hex}"
        refresh_token = ""
        
        # Salva na sessão local (SQLite)
        conn = get_connection()
        cur = conn.cursor()
        cur.execute('''
            INSERT OR REPLACE INTO sessoes (user_id, email, access_token, refresh_token, is_admin, ultimo_login)
            VALUES (?, ?, ?, ?, ?, datetime('now'))
        ''', (user_id, email, access_token, refresh_token, int(is_admin)))
        
        # Sincroniza esse usuário para a tabela de offline
        cur.execute('''
            INSERT OR REPLACE INTO usuarios_locais (email, senha_hash, is_admin)
            VALUES (?, ?, ?)
        ''', (email, user_cloud.get("senha_hash"), int(is_admin)))
        
        conn.commit()
        conn.close()
        
        return {
            "access_token": access_token,
            "user_id": user_id,
            "email": email,
            "is_admin": is_admin
        }
        
    except HTTPException:
        raise
    except Exception as e:
        logger.error(f"Erro no login (cloud): {e}")
        raise HTTPException(status_code=401, detail="Erro de conexão com servidor ou credenciais inválidas.")


@router.get("/me")
def get_current_user(token: str):
    """
    Valida o token localmente e retorna os dados do usuário.
    """
    conn = get_row_connection()
    cur = conn.cursor()
    cur.execute("SELECT * FROM sessoes WHERE access_token = ?", (token,))
    sessao = cur.fetchone()
    conn.close()
    
    if not sessao:
        raise HTTPException(status_code=401, detail="Sessão inválida ou expirada.")
        
    return {
        "user_id": sessao["user_id"],
        "email": sessao["email"],
        "is_admin": bool(sessao["is_admin"])
    }
