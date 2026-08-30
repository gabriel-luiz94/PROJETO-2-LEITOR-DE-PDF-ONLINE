"""
middleware/auth_middleware.py — Middleware global de autenticação JWT.

Protege todas as rotas /api/* automaticamente (exceto whitelist).
Extrai user_id e role do JWT e injeta em request.state.
"""
import os
import jwt
from datetime import datetime, timezone
from fastapi import Request, HTTPException
from starlette.middleware.base import BaseHTTPMiddleware
from starlette.responses import JSONResponse
from config import logger

# Chave secreta para assinar tokens JWT
JWT_SECRET = os.environ.get("JWT_SECRET", "dev-secret-change-in-production")
JWT_ALGORITHM = "HS256"
JWT_EXPIRY_HOURS = 24

# Rotas que NÃO exigem autenticação
PUBLIC_ROUTES = {
    "/api/auth/login",
    "/api/health",
    "/api/health/sync-master",
    "/api/update/check",
    "/api/projetos",
    "/api/orcamento/dados",
    "/api/orcamento/calcular",
    "/api/orcamento/search",
    "/api/orcamento/detalhes",
    "/trigger-file",
}

# Prefixos que NÃO exigem autenticação (frontend, static, etc.)
PUBLIC_PREFIXES = (
    "/static",
    "/",
    "/login",
    "/resumo",
    "/admin",
    "/upload",     # Upload de arquivo PDF/DXF (protegido por IS_FROZEN no router)
    "/extract-local",
    "/ws",
    "/docs",
    "/openapi.json",
)


def create_jwt_token(user_id: str, email: str, role: str) -> str:
    """Cria um token JWT com expiração."""
    from datetime import timedelta
    now = datetime.now(timezone.utc)
    payload = {
        "sub": user_id,
        "email": email,
        "role": role,
        "iat": now,
        "exp": now + timedelta(hours=JWT_EXPIRY_HOURS),
    }
    return jwt.encode(payload, JWT_SECRET, algorithm=JWT_ALGORITHM)


def create_refresh_token(user_id: str) -> str:
    """Cria um refresh token com expiração longa (30 dias)."""
    from datetime import timedelta
    now = datetime.now(timezone.utc)
    payload = {
        "sub": user_id,
        "type": "refresh",
        "iat": now,
        "exp": now + timedelta(days=30),
    }
    return jwt.encode(payload, JWT_SECRET, algorithm=JWT_ALGORITHM)


def decode_jwt_token(token: str) -> dict | None:
    """Decodifica e valida um token JWT. Retorna None se inválido/expirado."""
    try:
        payload = jwt.decode(token, JWT_SECRET, algorithms=[JWT_ALGORITHM])
        return payload
    except jwt.ExpiredSignatureError:
        return None
    except jwt.InvalidTokenError:
        return None


class AuthMiddleware(BaseHTTPMiddleware):
    """
    Middleware que intercepta requests para /api/* e valida o JWT.
    Rotas públicas e prefixos públicos são ignorados.
    O user é injetado em request.state.user (dict com sub, email, role).
    """

    async def dispatch(self, request: Request, call_next):
        if request.scope.get("type") != "http":
            return await call_next(request)

        path = request.url.path

        # Rotas públicas — passa direto
        if path in PUBLIC_ROUTES:
            return await call_next(request)

        # Prefixos públicos — passa direto
        for prefix in PUBLIC_PREFIXES:
            if prefix == "/":
                if path == "/":
                    return await call_next(request)
            elif path.startswith(prefix):
                return await call_next(request)

        # Rotas /api/* — exigem autenticação
        if path.startswith("/api/"):
            token = None

            # Tenta extrair do header Authorization: Bearer <token>
            auth_header = request.headers.get("Authorization", "")
            if auth_header.startswith("Bearer "):
                token = auth_header[7:]

            # Fallback: query parameter ?token=...
            if not token:
                token = request.query_params.get("token", "")

            if not token:
                return JSONResponse(
                    status_code=401,
                    content={"detail": "Token de autenticação não fornecido."}
                )

            payload = decode_jwt_token(token)
            if not payload:
                return JSONResponse(
                    status_code=401,
                    content={"detail": "Token inválido ou expirado."}
                )

            # Injeta dados do usuário no request
            request.state.user = {
                "user_id": payload.get("sub"),
                "email": payload.get("email"),
                "role": payload.get("role", "operador"),
            }

        return await call_next(request)


def require_role(*roles: str):
    """
    Dependency para exigir role específica em uma rota.
    Uso: @router.get("/...", dependencies=[Depends(require_role("admin"))])
    """
    def _dependency(request: Request):
        user = getattr(request.state, "user", None)
        if not user:
            raise HTTPException(status_code=401, detail="Não autenticado.")
        if user.get("role") not in roles:
            raise HTTPException(
                status_code=403,
                detail=f"Acesso negado. Requer role: {', '.join(roles)}"
            )
        return user
    return _dependency


def get_current_user_from_state(request: Request) -> dict:
    """Extrai o usuário do request.state (populado pelo middleware)."""
    user = getattr(request.state, "user", None)
    if not user:
        raise HTTPException(status_code=401, detail="Não autenticado.")
    return user
