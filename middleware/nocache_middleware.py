"""
middleware/nocache_middleware.py — Middleware de no-cache para arquivos HTML estáticos.

O StaticFiles mount do Starlette serve arquivos sem cabecalhos de cache-control,
o que faz o navegador guardar versoes antigas de .html. Este middleware injeta
NO_CACHE_HEADERS em toda resposta de /static/*.html, garantindo que o browser
sempre busque a versao mais recente do servidor.
"""
import uuid

from starlette.middleware.base import BaseHTTPMiddleware
from config import APP_MODE, NO_CACHE_HEADERS

# TASK-168: identificador desta abertura do programa (muda a cada execução). O navegador embutido guarda o localStorage entre aberturas; o JS compara este
# cookie com o que guardou e, se mudou, esquece o projeto aberto: o programa sempre abre no projeto default e o usuário abre o projeto.
SESSAO_ID = uuid.uuid4().hex
COOKIE_SESSAO = "gridr_sessao"


class NoCacheHtmlMiddleware(BaseHTTPMiddleware):
    """Injeta NO_CACHE_HEADERS em respostas de /static/*.html."""

    async def dispatch(self, request, call_next):
        response = await call_next(request)
        path = request.url.path
        # Aplica apenas a arquivos HTML servidos pelo mount /static/
        if path.startswith("/static/") and path.endswith(".html"):
            for key, value in NO_CACHE_HEADERS.items():
                response.headers[key] = value
        if APP_MODE != "server" and "text/html" in (response.headers.get("content-type") or "") and request.cookies.get(COOKIE_SESSAO) != SESSAO_ID:
            response.set_cookie(COOKIE_SESSAO, SESSAO_ID, path="/", samesite="lax")          # só no programa instalado; a versão online não muda
        return response
