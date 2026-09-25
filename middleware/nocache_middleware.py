"""
middleware/nocache_middleware.py — Middleware de no-cache para arquivos HTML estáticos.

O StaticFiles mount do Starlette serve arquivos sem cabecalhos de cache-control,
o que faz o navegador guardar versoes antigas de .html. Este middleware injeta
NO_CACHE_HEADERS em toda resposta de /static/*.html, garantindo que o browser
sempre busque a versao mais recente do servidor.
"""
from starlette.middleware.base import BaseHTTPMiddleware
from config import NO_CACHE_HEADERS


class NoCacheHtmlMiddleware(BaseHTTPMiddleware):
    """Injeta NO_CACHE_HEADERS em respostas de /static/*.html."""

    async def dispatch(self, request, call_next):
        response = await call_next(request)
        path = request.url.path
        # Aplica apenas a arquivos HTML servidos pelo mount /static/
        if path.startswith("/static/") and path.endswith(".html"):
            for key, value in NO_CACHE_HEADERS.items():
                response.headers[key] = value
        return response
