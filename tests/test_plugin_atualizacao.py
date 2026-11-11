"""TASK-130 — atualização automática do plugin nanoCAD: a aplicação publica a DLL do repositório (versão + SHA-256), sem login."""
import hashlib
import json
import re
from pathlib import Path

import pytest
from fastapi.testclient import TestClient

import app as appmod
import routers.plugin_cad as plugin_cad

RAIZ = Path(__file__).resolve().parent.parent
PASTA = RAIZ / "plugin_nanocad" / "appload" / "nanoCAD26"


@pytest.fixture
def client():
    return TestClient(appmod.app)


@pytest.mark.parametrize("pasta", ["nanoCAD26", "nanoCAD5"])
def test_manifesto_versionado_confere_com_a_dll_e_com_a_versao_do_plugin(pasta):
    p = RAIZ / "plugin_nanocad" / "appload" / pasta
    m = json.loads((p / "versao.json").read_text(encoding="utf-8"))
    assert m["sha256"] == hashlib.sha256((p / "OrcamentoNanoCAD.dll").read_bytes()).hexdigest(), \
        "rode plugin_nanocad/compilar.sh (o manifesto ficou velho)"
    cs = (RAIZ / "plugin_nanocad" / "src" / "Nucleo" / "ObraArquivo.cs").read_text(encoding="utf-8")
    assert m["versao"] == re.search(r'VersaoPlugin = "([^"]+)"', cs).group(1)
    props = (RAIZ / "plugin_nanocad" / "src" / "Versao.props").read_text(encoding="utf-8")
    assert m["versao"] == re.search(r"<Version>([^<]+)</Version>", props).group(1)


def test_versao_e_arquivo_publicos_sem_login(client):
    r = client.get("/api/plugin/atualizacao?alvo=nanocad26")
    assert r.status_code == 200
    j = r.json()
    dll = (PASTA / "OrcamentoNanoCAD.dll").read_bytes()
    assert j["sha256"] == hashlib.sha256(dll).hexdigest() and j["tamanho"] == len(dll)
    a = client.get(j["url"])
    assert a.status_code == 200 and a.content == dll
    assert a.headers["x-plugin-versao"] == j["versao"]
    # o resto da API do plugin continua exigindo login
    assert client.get("/api/plugin/versao?projeto=081").status_code == 401


def test_nanocad5_tambem_e_publicado(client):
    r = client.get("/api/plugin/atualizacao?alvo=nanocad5")
    assert r.status_code == 200 and r.json()["versao"]


def test_alvo_desconhecido_ou_sem_publicacao_da_404(client, tmp_path, monkeypatch):
    assert client.get("/api/plugin/atualizacao?alvo=autocad").status_code == 404
    monkeypatch.setattr(plugin_cad, "PASTA_APPLOAD", str(tmp_path))
    assert client.get("/api/plugin/atualizacao?alvo=nanocad5").status_code == 404


def test_manifesto_que_nao_confere_nao_publica(client, tmp_path, monkeypatch):
    p = tmp_path / "nanoCAD26"
    p.mkdir()
    (p / "OrcamentoNanoCAD.dll").write_bytes(b"MZ nova")
    (p / "versao.json").write_text(json.dumps({"versao": "9.9.9", "sha256": "0" * 64}), encoding="utf-8")
    monkeypatch.setattr(plugin_cad, "PASTA_APPLOAD", str(tmp_path))
    assert client.get("/api/plugin/atualizacao?alvo=nanocad26").status_code == 503
    assert client.get("/api/plugin/atualizacao/arquivo?alvo=nanocad26").status_code == 503
