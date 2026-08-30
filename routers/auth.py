"""
routers/auth.py — Rotas para Autenticação com JWT + bcrypt.

Login com fallback offline: tenta Supabase (nuvem) → fallback SQLite (local).
Tokens JWT com expiração. Migração automática de senhas SHA-256 legadas para bcrypt.
"""
import uuid
from fastapi import APIRouter, HTTPException, Request
from pydantic import BaseModel
from services.supabase_client import get_supabase
from database import (
    get_connection, get_row_connection,
    hash_password, verify_password, _migrate_legacy_password
)
from middleware.auth_middleware import (
    create_jwt_token, create_refresh_token, decode_jwt_token
)
from config import logger

router = APIRouter(prefix="/api/auth", tags=["auth"])


class LoginRequest(BaseModel):
    email: str
    password: str
    lembrar: bool = False


class AuthResponse(BaseModel):
    access_token: str
    refresh_token: str
    user_id: str
    email: str
    is_admin: bool
    role: str


@router.post("/login", response_model=AuthResponse)
def login(req: LoginRequest):
    """
    Login com fallback offline:
    1. Valida no banco de usuários locais (SQLite)
    2. Se não encontrar, tenta Supabase (nuvem)
    3. Gera token JWT com expiração
    """

    # 1. Tentar validação direta no Supabase (Nuvem) primeiro
    supabase = get_supabase()
    if supabase:
        try:
            res = supabase.table("usuarios_nuvem").select("*").eq("email", req.email).execute()
            if res.data and len(res.data) > 0:
                user_cloud = res.data[0]
                if verify_password(req.password, user_cloud.get("senha_hash", "")):
                    user_id = str(user_cloud.get("id"))
                    email = user_cloud.get("email")
                    is_admin = bool(user_cloud.get("is_admin"))
                    role = user_cloud.get("role", "admin" if is_admin else "operador")

                    access_token = create_jwt_token(user_id, email, role)
                    refresh_token = create_refresh_token(user_id)

                    # Atualiza o cache local
                    conn = get_connection()
                    cur = conn.cursor()
                    cur.execute('''
                        INSERT OR REPLACE INTO usuarios_locais (email, senha_hash, is_admin, role)
                        VALUES (?, ?, ?, ?)
                    ''', (email, user_cloud.get("senha_hash"), int(is_admin), role))
                    cur.execute('''
                        INSERT OR REPLACE INTO sessoes (user_id, email, access_token, refresh_token, is_admin, role, ultimo_login)
                        VALUES (?, ?, ?, ?, ?, ?, datetime('now'))
                    ''', (user_id, email, access_token, refresh_token, int(is_admin), role))
                    conn.commit()
                    conn.close()

                    return {
                        "access_token": access_token,
                        "refresh_token": refresh_token,
                        "user_id": user_id,
                        "email": email,
                        "is_admin": is_admin,
                        "role": role,
                    }
                else:
                    raise HTTPException(status_code=401, detail="Senha incorreta.")
        except HTTPException:
            raise
        except Exception as e:
            logger.warning(f"Aviso ao consultar Supabase auth: {e}. Tentando fallback local...")

    # 2. Fallback: Tentar login local (SQLite)
    conn = get_row_connection()
    cur = conn.cursor()
    cur.execute("SELECT * FROM usuarios_locais WHERE email = ?", (req.email,))
    user_local = cur.fetchone()
    conn.close()

    if user_local:
        if verify_password(req.password, user_local["senha_hash"]):
            user_id = str(user_local["id"])
            email = user_local["email"]
            is_admin = bool(user_local["is_admin"])
            role = user_local["role"] if user_local["role"] else ("admin" if is_admin else "operador")

            # Migrar senha legada SHA-256 → bcrypt se necessário
            if not user_local["senha_hash"].startswith("$2"):
                conn = get_connection()
                _migrate_legacy_password(email, req.password, conn)
                conn.close()

            access_token = create_jwt_token(user_id, email, role)
            refresh_token = create_refresh_token(user_id)

            conn = get_connection()
            cur = conn.cursor()
            cur.execute('''
                INSERT OR REPLACE INTO sessoes (user_id, email, access_token, refresh_token, is_admin, role, ultimo_login)
                VALUES (?, ?, ?, ?, ?, ?, datetime('now'))
            ''', (user_id, email, access_token, refresh_token, int(is_admin), role))
            conn.commit()
            conn.close()

            return {
                "access_token": access_token,
                "refresh_token": refresh_token,
                "user_id": user_id,
                "email": email,
                "is_admin": is_admin,
                "role": role,
            }
        else:
            raise HTTPException(status_code=401, detail="Senha incorreta.")

    raise HTTPException(status_code=401, detail="Credenciais inválidas ou usuário não encontrado.")

    try:
        res = supabase.table("usuarios_nuvem").select("*").eq("email", req.email).execute()
        if not res.data:
            raise HTTPException(status_code=401, detail="Credenciais inválidas ou usuário não encontrado.")

        user_cloud = res.data[0]

        if not verify_password(req.password, user_cloud.get("senha_hash", "")):
            raise HTTPException(status_code=401, detail="Credenciais inválidas.")

        user_id = str(user_cloud.get("id"))
        email = user_cloud.get("email")
        is_admin = bool(user_cloud.get("is_admin"))
        role = user_cloud.get("role", "admin" if is_admin else "operador")

        # Gerar tokens JWT
        access_token = create_jwt_token(user_id, email, role)
        refresh_token = create_refresh_token(user_id)

        # Salvar na sessão local (SQLite)
        conn = get_connection()
        cur = conn.cursor()
        cur.execute('''
            INSERT OR REPLACE INTO sessoes (user_id, email, access_token, refresh_token, is_admin, role, ultimo_login)
            VALUES (?, ?, ?, ?, ?, ?, datetime('now'))
        ''', (user_id, email, access_token, refresh_token, int(is_admin), role))

        # Sincroniza usuário para tabela offline
        cloud_hash = user_cloud.get("senha_hash", "")
        cur.execute('''
            INSERT OR REPLACE INTO usuarios_locais (email, senha_hash, is_admin, role)
            VALUES (?, ?, ?, ?)
        ''', (email, cloud_hash, int(is_admin), role))

        conn.commit()
        conn.close()

        return {
            "access_token": access_token,
            "refresh_token": refresh_token,
            "user_id": user_id,
            "email": email,
            "is_admin": is_admin,
            "role": role,
        }

    except HTTPException:
        raise
    except Exception as e:
        logger.error(f"Erro no login (cloud): {e}")
        raise HTTPException(status_code=401, detail="Erro de conexão com servidor ou credenciais inválidas.")


@router.post("/refresh")
def refresh_token(request: Request):
    """Renova o access_token usando o refresh_token."""
    auth_header = request.headers.get("Authorization", "")
    if not auth_header.startswith("Bearer "):
        raise HTTPException(status_code=401, detail="Refresh token não fornecido.")

    token = auth_header[7:]
    payload = decode_jwt_token(token)

    if not payload or payload.get("type") != "refresh":
        raise HTTPException(status_code=401, detail="Refresh token inválido ou expirado.")

    user_id = payload.get("sub")

    # Buscar dados do usuário
    conn = get_row_connection()
    cur = conn.cursor()
    cur.execute("SELECT * FROM sessoes WHERE user_id = ?", (user_id,))
    sessao = cur.fetchone()
    conn.close()

    if not sessao:
        raise HTTPException(status_code=401, detail="Sessão não encontrada.")

    email = sessao["email"]
    role = sessao["role"] or ("admin" if sessao["is_admin"] else "operador")

    new_access_token = create_jwt_token(user_id, email, role)

    # Atualizar sessão
    conn = get_connection()
    cur = conn.cursor()
    cur.execute(
        "UPDATE sessoes SET access_token = ? WHERE user_id = ?",
        (new_access_token, user_id)
    )
    conn.commit()
    conn.close()

    return {
        "access_token": new_access_token,
        "user_id": user_id,
        "email": email,
        "role": role,
    }


@router.get("/me")
def get_current_user(request: Request):
    """
    Retorna os dados do usuário autenticado.
    Usa o user injetado pelo middleware JWT.
    """
    user = getattr(request.state, "user", None)
    if not user:
        raise HTTPException(status_code=401, detail="Não autenticado.")
    return user
