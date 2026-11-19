"""
routers/desktop_arquivos.py (TASK-193) — entrega os bytes de um arquivo que o usuário ESCOLHEU na caixa «Abrir» da janela desktop.

`GET /api/desktop/arquivo/{id}`: o id vem de `ApiArquivos.abrir_arquivos` (services/arquivos_desktop.py) e vale 10 minutos; só existem ids de arquivos escolhidos pelo
usuário nesta sessão (nenhum caminho arbitrário é servido) e a rota exige login como as demais. Na versão online não há janela nativa: nunca há ids (sempre 404).
"""
from fastapi import APIRouter, HTTPException
from fastapi.responses import FileResponse

from services import arquivos_desktop

router = APIRouter(prefix="/api/desktop", tags=["desktop"])


@router.get("/arquivo/{arquivo_id}")
def baixar_escolhido(arquivo_id: str):
    caminho = arquivos_desktop.caminho_escolhido(arquivo_id)
    if not caminho:
        raise HTTPException(status_code=404, detail="Arquivo não encontrado (escolha-o de novo).")
    return FileResponse(caminho, media_type="application/octet-stream", headers={"Cache-Control": "no-store"})
