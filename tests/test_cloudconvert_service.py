"""tests/test_cloudconvert_service.py — conversão DWG → DXF via CloudConvert (TASK-043).

Nenhuma chamada de rede real: `httpx.post`/`httpx.get` são sempre mockados. `conftest.py` já zera
CLOUDCONVERT_API_KEY do ambiente para todos os testes."""
import pytest

from services import cloudconvert_service as cc


class _Resp:
    def __init__(self, status_code=200, dados=None, texto="", conteudo=b""):
        self.status_code = status_code
        self._dados = dados
        self.text = texto
        self.content = conteudo

    def json(self):
        return self._dados


def _job(status="waiting", com_form=True, com_arquivo_exportado=True, mensagem_erro=None):
    tarefas = [
        {"name": "importar", "operation": "import/upload",
         "result": {"form": {"url": "https://upload.cloudconvert.com/xyz", "parameters": {"a": "1"}}} if com_form else {}},
        {"name": "converter", "operation": "convert", "status": "error" if status == "error" else "finished",
         "message": mensagem_erro},
        {"name": "exportar", "operation": "export/url",
         "result": {"files": [{"url": "https://storage.cloudconvert.com/resultado.dxf"}]} if com_arquivo_exportado else {}},
    ]
    return {"id": "job-123", "status": status, "tasks": tarefas}


def test_sem_api_key_nao_chama_rede(monkeypatch):
    chamadas = []
    monkeypatch.setattr(cc.httpx, "post", lambda *a, **k: chamadas.append("post") or _Resp())
    monkeypatch.setattr(cc.httpx, "get", lambda *a, **k: chamadas.append("get") or _Resp())
    with pytest.raises(cc.ErroConversaoDwg, match="Nenhuma API key"):
        cc.converter_dwg_para_dxf(b"conteudo", "x.dwg", None)
    assert chamadas == []


def test_fluxo_completo_de_sucesso(monkeypatch):
    chamadas = []

    def post(url, **kw):
        chamadas.append(("post", url))
        if url == f"{cc.CLOUDCONVERT_API_BASE}/jobs":
            return _Resp(200, {"data": _job(status="waiting")})
        if url == "https://upload.cloudconvert.com/xyz":
            return _Resp(200)
        raise AssertionError(f"POST inesperado: {url}")

    def get(url, **kw):
        chamadas.append(("get", url))
        if url == f"{cc.CLOUDCONVERT_API_BASE}/jobs/job-123":
            return _Resp(200, {"data": _job(status="finished")})
        if url == "https://storage.cloudconvert.com/resultado.dxf":
            return _Resp(200, conteudo=b"CONTEUDO DXF CONVERTIDO")
        raise AssertionError(f"GET inesperado: {url}")

    monkeypatch.setattr(cc.httpx, "post", post)
    monkeypatch.setattr(cc.httpx, "get", get)
    resultado = cc.converter_dwg_para_dxf(b"bytes do dwg", "projeto.dwg", "chave-valida")
    assert resultado == b"CONTEUDO DXF CONVERTIDO"
    assert ("post", f"{cc.CLOUDCONVERT_API_BASE}/jobs") in chamadas
    assert ("post", "https://upload.cloudconvert.com/xyz") in chamadas


def test_poll_espera_o_job_terminar(monkeypatch):
    """1ª consulta ainda 'waiting', 2ª 'finished' — sem dormir de verdade."""
    respostas = iter([_job(status="waiting"), _job(status="finished")])
    monkeypatch.setattr(cc.time, "sleep", lambda *_: None)
    monkeypatch.setattr(cc.httpx, "post", lambda url, **kw: (
        _Resp(200, {"data": _job()}) if url.endswith("/jobs") else _Resp(200)
    ))
    monkeypatch.setattr(cc.httpx, "get", lambda url, **kw: (
        _Resp(200, {"data": next(respostas)}) if "/jobs/" in url else _Resp(200, conteudo=b"OK")
    ))
    resultado = cc.converter_dwg_para_dxf(b"x", "a.dwg", "chave")
    assert resultado == b"OK"


def test_erro_na_criacao_do_job(monkeypatch):
    monkeypatch.setattr(cc.httpx, "post", lambda *a, **k: _Resp(401, texto="unauthorized"))
    with pytest.raises(cc.ErroConversaoDwg, match="recusou a criação do job"):
        cc.converter_dwg_para_dxf(b"x", "a.dwg", "chave-invalida")


def test_erro_ao_enviar_arquivo(monkeypatch):
    def post(url, **kw):
        if url.endswith("/jobs"):
            return _Resp(200, {"data": _job()})
        return _Resp(500, texto="falhou")
    monkeypatch.setattr(cc.httpx, "post", post)
    with pytest.raises(cc.ErroConversaoDwg, match="Falha ao enviar o arquivo"):
        cc.converter_dwg_para_dxf(b"x", "a.dwg", "chave")


def test_job_termina_com_erro(monkeypatch):
    monkeypatch.setattr(cc.httpx, "post", lambda url, **kw: (
        _Resp(200, {"data": _job()}) if url.endswith("/jobs") else _Resp(200)
    ))
    monkeypatch.setattr(cc.httpx, "get", lambda *a, **k: _Resp(200, {"data": _job(status="error", mensagem_erro="arquivo corrompido")}))
    with pytest.raises(cc.ErroConversaoDwg, match="arquivo corrompido"):
        cc.converter_dwg_para_dxf(b"x", "a.dwg", "chave")


def test_timeout_de_espera(monkeypatch):
    contador = {"n": 0}

    def monotonic_falso():
        contador["n"] += 1
        return 0 if contador["n"] == 1 else 10_000  # 1ª chamada define o limite; a 2ª já está "no futuro"
    monkeypatch.setattr(cc.time, "monotonic", monotonic_falso)
    monkeypatch.setattr(cc.time, "sleep", lambda *_: pytest.fail("não deveria dormir no caminho de timeout"))
    monkeypatch.setattr(cc.httpx, "post", lambda url, **kw: (
        _Resp(200, {"data": _job()}) if url.endswith("/jobs") else _Resp(200)
    ))
    monkeypatch.setattr(cc.httpx, "get", lambda *a, **k: _Resp(200, {"data": _job(status="waiting")}))
    with pytest.raises(cc.ErroConversaoDwg, match="timeout"):
        cc.converter_dwg_para_dxf(b"x", "a.dwg", "chave")


def test_job_finaliza_sem_arquivo_exportado(monkeypatch):
    monkeypatch.setattr(cc.httpx, "post", lambda url, **kw: (
        _Resp(200, {"data": _job()}) if url.endswith("/jobs") else _Resp(200)
    ))
    monkeypatch.setattr(cc.httpx, "get", lambda *a, **k: _Resp(200, {"data": _job(status="finished", com_arquivo_exportado=False)}))
    with pytest.raises(cc.ErroConversaoDwg, match="não devolveu nenhum arquivo"):
        cc.converter_dwg_para_dxf(b"x", "a.dwg", "chave")


def test_falha_ao_baixar_resultado(monkeypatch):
    def get(url, **kw):
        if "/jobs/" in url:
            return _Resp(200, {"data": _job(status="finished")})
        return _Resp(500)
    monkeypatch.setattr(cc.httpx, "post", lambda url, **kw: (
        _Resp(200, {"data": _job()}) if url.endswith("/jobs") else _Resp(200)
    ))
    monkeypatch.setattr(cc.httpx, "get", get)
    with pytest.raises(cc.ErroConversaoDwg, match="Falha ao baixar"):
        cc.converter_dwg_para_dxf(b"x", "a.dwg", "chave")
