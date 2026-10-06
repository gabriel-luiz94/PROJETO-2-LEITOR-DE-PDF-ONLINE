"""
services/cloudconvert_service.py — Conversão DWG → DXF via CloudConvert (TASK-043).

Função com I/O real (chamadas HTTP externas) — diferente dos motores de ajuste/validação,
que são puros. Usada só quando o arquivo é .dwg; .dxf/.pdf continuam 100% offline, sem
nenhuma chamada daqui.

Fluxo da API v2 do CloudConvert (https://cloudconvert.com/api/v2):
  1. Cria um Job com 3 tasks: import/upload, convert (dwg→dxf), export/url.
  2. Envia o arquivo (multipart) para a URL de upload que a task de import devolveu.
  3. Espera (polling) o job terminar.
  4. Baixa o arquivo resultante da task de export.
"""
import time

import httpx

CLOUDCONVERT_API_BASE = "https://api.cloudconvert.com/v2"
TIMEOUT_SEGUNDOS = 90
INTERVALO_POLL_SEGUNDOS = 2


class ErroConversaoDwg(Exception):
    """Falha ao converter um .dwg para .dxf via CloudConvert (chave ausente/inválida, cota
    esgotada, timeout, arquivo inválido etc.) — sempre com uma mensagem segura para o usuário."""


def _criar_job(api_key: str) -> dict:
    resp = httpx.post(
        f"{CLOUDCONVERT_API_BASE}/jobs",
        headers={"Authorization": f"Bearer {api_key}"},
        json={
            "tasks": {
                "importar": {"operation": "import/upload"},
                "converter": {"operation": "convert", "input": "importar",
                              "input_format": "dwg", "output_format": "dxf"},
                "exportar": {"operation": "export/url", "input": "converter"},
            }
        },
        timeout=30,
    )
    if resp.status_code >= 400:
        raise ErroConversaoDwg(f"CloudConvert recusou a criação do job ({resp.status_code}): {resp.text[:300]}")
    return resp.json()["data"]


def _enviar_arquivo(tarefa_import: dict, conteudo: bytes, nome_arquivo: str):
    form = tarefa_import["result"]["form"]
    resp = httpx.post(
        form["url"],
        data=form["parameters"],
        files={"file": (nome_arquivo, conteudo, "application/octet-stream")},
        timeout=60,
    )
    if resp.status_code >= 400:
        raise ErroConversaoDwg(f"Falha ao enviar o arquivo para o CloudConvert ({resp.status_code}): {resp.text[:300]}")


def _esperar_job(job_id: str, api_key: str) -> dict:
    limite = time.monotonic() + TIMEOUT_SEGUNDOS
    while True:
        resp = httpx.get(
            f"{CLOUDCONVERT_API_BASE}/jobs/{job_id}",
            headers={"Authorization": f"Bearer {api_key}"},
            timeout=30,
        )
        if resp.status_code >= 400:
            raise ErroConversaoDwg(f"Falha ao consultar o job no CloudConvert ({resp.status_code}): {resp.text[:300]}")
        job = resp.json()["data"]
        status = job.get("status")
        if status == "finished":
            return job
        if status == "error":
            tarefas = {t.get("name"): t for t in job.get("tasks", [])}
            erro_tarefa = next((t.get("message") for t in tarefas.values() if t.get("status") == "error"), None)
            raise ErroConversaoDwg(f"CloudConvert não conseguiu converter o arquivo: {erro_tarefa or 'erro desconhecido'}")
        if time.monotonic() >= limite:
            raise ErroConversaoDwg(f"Conversão não terminou em {TIMEOUT_SEGUNDOS}s (timeout).")
        time.sleep(INTERVALO_POLL_SEGUNDOS)


def _baixar_resultado(job: dict) -> bytes:
    tarefas = {t.get("name"): t for t in job.get("tasks", [])}
    exportar = tarefas.get("exportar") or {}
    arquivos = (exportar.get("result") or {}).get("files") or []
    if not arquivos:
        raise ErroConversaoDwg("CloudConvert terminou o job, mas não devolveu nenhum arquivo exportado.")
    resp = httpx.get(arquivos[0]["url"], timeout=60)
    if resp.status_code >= 400:
        raise ErroConversaoDwg(f"Falha ao baixar o DXF convertido ({resp.status_code}).")
    return resp.content


def converter_dwg_para_dxf(conteudo_dwg: bytes, nome_arquivo: str, api_key: str | None) -> bytes:
    """Converte bytes de um .dwg em bytes de um .dxf via CloudConvert. Levanta ErroConversaoDwg
    em qualquer falha (sem API key, rede, cota, arquivo inválido etc.)."""
    if not api_key:
        raise ErroConversaoDwg(
            "Nenhuma API key do CloudConvert configurada (nem sua, nem padrão do sistema) — "
            "necessária para converter .dwg (crie uma gratuita em cloudconvert.com)."
        )
    job = _criar_job(api_key)
    tarefas = {t["name"]: t for t in job["tasks"]}
    _enviar_arquivo(tarefas["importar"], conteudo_dwg, nome_arquivo)
    job_final = _esperar_job(job["id"], api_key)
    return _baixar_resultado(job_final)
