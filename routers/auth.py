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
    supabase = get_supabase()
    if not supabase:
        # Modo Offline: Tenta buscar na sessão local
        conn = get_row_connection()
        cur = conn.cursor()
        cur.execute("SELECT * FROM sessoes WHERE email = ?", (req.email,))
        sessao = cur.fetchone()
        conn.close()
        
        if sessao:
            # Em modo estritamente offline, não podemos validar a senha de forma segura se não salvamos o hash.
            # Se a sessão existe localmente, assumimos que o token ainda é válido para acesso offline.
            # NOTA: Em produção real, deveríamos salvar um hash local ou usar o refresh token quando voltar a internet.
            return {
                "access_token": sessao["access_token"],
                "user_id": sessao["user_id"],
                "email": sessao["email"],
                "is_admin": bool(sessao["is_admin"])
            }
        raise HTTPException(status_code=503, detail="Sem conexão com servidor de login e sem sessão offline salva.")

    try:
        # Tenta logar via Supabase
        res = supabase.auth.sign_in_with_password({"email": req.email, "password": req.password})
        
        user_id = res.user.id
        email = res.user.email
        access_token = res.session.access_token
        refresh_token = res.session.refresh_token
        
        # Simples verificação de admin
        ADMIN_EMAILS = ["valdecinunesaf@gmail.com"]
        is_admin = email in ADMIN_EMAILS 
        
        # Salva na sessão local (SQLite)
        conn = get_connection()
        cur = conn.cursor()
        cur.execute('''
            INSERT OR REPLACE INTO sessoes (user_id, email, access_token, refresh_token, is_admin, ultimo_login)
            VALUES (?, ?, ?, ?, ?, datetime('now'))
        ''', (user_id, email, access_token, refresh_token, int(is_admin)))
        conn.commit()
        conn.close()
        
        return {
            "access_token": access_token,
            "user_id": user_id,
            "email": email,
            "is_admin": is_admin
        }
        
    except Exception as e:
        logger.error(f"Erro no login: {e}")
        raise HTTPException(status_code=401, detail="Credenciais inválidas.")


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
