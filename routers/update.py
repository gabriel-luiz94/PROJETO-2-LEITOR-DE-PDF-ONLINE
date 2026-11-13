"""
routers/update.py — Rota para verificar e aplicar atualizações.

A rota /check é pública, usada pelos clientes desktop para ver se há nova versão.
A rota /apply (POST) é usada pelo cliente desktop para baixar e instalar a nova versão.
"""
from fastapi import APIRouter, HTTPException
from pydantic import BaseModel
from config import APP_VERSION, APP_MODE
from services.auto_updater import cancelar_atualizacao, check_for_updates, estado_atualizacao, iniciar_atualizacao

router = APIRouter(prefix="/api/update", tags=["update"])


class UpdateRequest(BaseModel):
    download_url: str | None = None      # TASK-127: ignorado — a URL é sempre a da Release do GitHub consultada pelo servidor local


@router.get("/check")
def check_update():
    """
    Retorna as informações da versão mais recente.
    Se estiver rodando no modo 'server', retorna os dados reais de release 
    (no momento simulados, em produção pode ler do banco ou Github Releases).
    Se estiver rodando no modo 'desktop', chama o serviço local para consultar o server.
    """
    if APP_MODE == "server":
        # Simulando que o servidor sabe qual é a última versão. 
        # Em produção, você poderia atualizar isso numa tabela do banco via Painel Admin.
        # Ex: SELECT valor FROM configuracoes WHERE chave = 'latest_desktop_version'
        from database import get_row_connection
        conn = get_row_connection()
        cur = conn.cursor()
        cur.execute("SELECT valor FROM configuracoes WHERE chave = 'latest_desktop_version'")
        row_version = cur.fetchone()
        
        cur.execute("SELECT valor FROM configuracoes WHERE chave = 'latest_desktop_url'")
        row_url = cur.fetchone()
        conn.close()
        
        latest_version = row_version["valor"] if row_version else APP_VERSION
        latest_url = row_url["valor"] if row_url else ""
        
        return {
            "latest_version": latest_version,
            "download_url": latest_url,
            "release_notes": "Correção de bugs e melhorias no Sync."
        }
    else:
        # Modo Desktop consultando o Servidor
        return check_for_updates()


@router.post("/apply")
def apply_update(req: UpdateRequest):
    """(Apenas Desktop) Começa a baixar a Release mais recente do GitHub em segundo plano (conferindo o SHA-256) e reiniciar o app. O corpo do pedido não decide
    o que é baixado. A resposta é imediata; o andamento vem de GET /status e dá para cancelar em POST /cancel (TASK-159)."""
    if APP_MODE == "server":
        raise HTTPException(status_code=400, detail="Esta rota só funciona no cliente Desktop.")

    try:
        iniciar_atualizacao()
        return {"status": "iniciada", "msg": "Atualização iniciada."}
    except RuntimeError as e:
        raise HTTPException(status_code=400, detail=str(e))
    except Exception as e:
        raise HTTPException(status_code=500, detail=f"Erro ao aplicar atualização: {e}")


@router.get("/status")
def status_update():
    """Andamento da atualização: fase (ocioso | verificando | baixando | conferindo | pronto | cancelado | erro), bytes baixados, total e percentual."""
    return estado_atualizacao()


@router.post("/cancel")
def cancelar_update():
    """Cancela a atualização em andamento (o .exe atual não é tocado e o arquivo parcial é apagado)."""
    return {"cancelado": cancelar_atualizacao()}
