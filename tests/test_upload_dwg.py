"""tests/test_upload_dwg.py — dispatch de .dwg em POST /upload e GET /extract-local (TASK-043).

A conversão real (services/cloudconvert_service.converter_dwg_para_dxf) é sempre mockada — nenhuma
chamada de rede aqui; esses testes cobrem só a resolução de credencial e a ligação com o resto do
pipeline de extração (que já é testado em outros arquivos)."""
import io

import ezdxf
import pytest
from fastapi.testclient import TestClient

import app as appmod
import database
from routers import upload as upload_router
from services.cloudconvert_service import ErroConversaoDwg


@pytest.fixture(autouse=True)
def _banco_inicializado():
    database.init_db()


@pytest.fixture
def client():
    return TestClient(appmod.app)


def _dxf_bytes(textos=("1-U4",)) -> bytes:
    doc = ezdxf.new()
    msp = doc.modelspace()
    for i, t in enumerate(textos):
        msp.add_text(t, dxfattribs={"color": 1, "insert": (i * 10, 0), "height": 2})
    buf = io.StringIO()
    doc.write(buf)
    return buf.getvalue().encode("utf-8")


# ── resolução de credencial (precedência: usuário > salva > padrão do ambiente) ────────────────
def test_resolver_sem_nada_configurado_devolve_nenhuma():
    assert upload_router.resolver_chave_cloudconvert(None) == (None, "nenhuma")


def test_resolver_usa_env_como_padrao(monkeypatch):
    monkeypatch.setenv("CLOUDCONVERT_API_KEY", "chave-do-ambiente")
    assert upload_router.resolver_chave_cloudconvert(None) == ("chave-do-ambiente", "padrao")


def test_resolver_prefere_salva_sobre_padrao(monkeypatch):
    monkeypatch.setenv("CLOUDCONVERT_API_KEY", "chave-do-ambiente")
    conn = database.get_connection()
    conn.execute("INSERT INTO configuracoes (chave, valor) VALUES ('cloudconvert_api_key', 'chave-salva')")
    conn.commit()
    conn.close()
    assert upload_router.resolver_chave_cloudconvert(None) == ("chave-salva", "salva")


def test_resolver_prefere_header_sobre_todo_o_resto(monkeypatch):
    monkeypatch.setenv("CLOUDCONVERT_API_KEY", "chave-do-ambiente")
    conn = database.get_connection()
    conn.execute("INSERT INTO configuracoes (chave, valor) VALUES ('cloudconvert_api_key', 'chave-salva')")
    conn.commit()
    conn.close()
    assert upload_router.resolver_chave_cloudconvert("chave-do-usuario") == ("chave-do-usuario", "usuario")


def test_origem_chave_ativa_nunca_expõe_o_valor(monkeypatch):
    monkeypatch.setenv("CLOUDCONVERT_API_KEY", "segredo")
    assert upload_router.origem_chave_cloudconvert_ativa() == "padrao"


# ── POST /upload com .dwg ───────────────────────────────────────────────────────────────────────
def test_upload_dwg_sem_chave_devolve_erro_sem_crashar_nem_chamar_rede():
    """Sem mock: a função real de conversão roda, mas recusa antes de qualquer chamada httpx
    (coberto em test_cloudconvert_service.py::test_sem_api_key_nao_chama_rede). Aqui confirmamos
    que a rota devolve um erro limpo em vez de deixar a exceção estourar."""
    client = TestClient(appmod.app)
    resp = client.post("/upload", files={"file": ("projeto.dwg", b"bytes nao usados", "application/octet-stream")})
    assert resp.status_code == 200
    body = resp.json()
    assert body["data"] == [] and "Nenhuma API key" in body["error"]


def test_upload_dwg_convertido_com_sucesso_extrai_como_dxf(client, monkeypatch):
    monkeypatch.setattr(upload_router, "converter_dwg_para_dxf", lambda conteudo, nome, chave: _dxf_bytes(["1-U4", "2-CFU"]))
    resp = client.post(
        "/upload",
        files={"file": ("projeto.dwg", b"conteudo binario do dwg", "application/octet-stream")},
        headers={"X-CloudConvert-Key": "chave-do-usuario"},
    )
    assert resp.status_code == 200
    body = resp.json()
    assert "error" not in body
    ativos = [item.get("texto") for item in body["data"]]
    assert "1-U4" in ativos and "2-CFU" in ativos


def test_upload_dwg_chave_do_usuario_e_persistida_apos_sucesso(client, monkeypatch):
    monkeypatch.setattr(upload_router, "converter_dwg_para_dxf", lambda conteudo, nome, chave: _dxf_bytes())
    resp = client.post(
        "/upload",
        files={"file": ("a.dwg", b"x", "application/octet-stream")},
        headers={"X-CloudConvert-Key": "chave-nova-do-usuario"},
    )
    assert resp.status_code == 200 and "error" not in resp.json()
    assert upload_router.resolver_chave_cloudconvert(None) == ("chave-nova-do-usuario", "salva")


def test_upload_dwg_falha_de_conversao_devolve_erro_sem_crash(client, monkeypatch):
    def quebra(conteudo, nome, chave):
        raise ErroConversaoDwg("CloudConvert não conseguiu converter o arquivo: formato inválido")
    monkeypatch.setattr(upload_router, "converter_dwg_para_dxf", quebra)
    resp = client.post(
        "/upload",
        files={"file": ("ruim.dwg", b"x", "application/octet-stream")},
        headers={"X-CloudConvert-Key": "chave"},
    )
    assert resp.status_code == 200  # nunca derruba a rota, mesmo padrão do .dxf/.pdf quebrados
    body = resp.json()
    assert body["data"] == [] and "formato inválido" in body["error"]
    # a chave que falhou não é persistida (só persiste quando a conversão dá certo)
    assert upload_router.resolver_chave_cloudconvert(None) == (None, "nenhuma")


def test_upload_dxf_continua_funcionando_sem_cloudconvert(client, monkeypatch):
    chamadas = []
    monkeypatch.setattr(upload_router, "converter_dwg_para_dxf", lambda *a, **k: chamadas.append(1))
    resp = client.post("/upload", files={"file": ("normal.dxf", _dxf_bytes(["1-U4"]), "application/octet-stream")})
    assert resp.status_code == 200
    body = resp.json()
    assert "error" not in body and chamadas == []  # .dxf nunca passa pelo CloudConvert
