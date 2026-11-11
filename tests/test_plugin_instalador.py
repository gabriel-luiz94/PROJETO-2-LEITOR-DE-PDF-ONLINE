"""TASK-131 — o programa principal traz o plugin do nanoCAD e o instala em %APPDATA%\\OrcamentoNanoCAD\\plugin (sem voltar versão)."""
import json
from pathlib import Path

import pytest
from fastapi.testclient import TestClient

import app as appmod
import config
import routers.plugin_cad as plugin_cad
from middleware.auth_middleware import create_jwt_token
from services import plugin_instalador as inst

RAIZ = Path(__file__).resolve().parent.parent


@pytest.fixture
def pastas(tmp_path, monkeypatch):
    origem = tmp_path / "appload"
    for pasta, conteudo in (("nanoCAD26", b"MZ 26 nova"), ("nanoCAD5", b"MZ 5 nova")):
        (origem / pasta).mkdir(parents=True)
        (origem / pasta / "OrcamentoNanoCAD.dll").write_bytes(conteudo)
        (origem / pasta / "versao.json").write_text(json.dumps({"versao": "0.10.0"}), encoding="utf-8")
    destino = tmp_path / "destino"
    monkeypatch.setattr(plugin_cad, "PASTA_APPLOAD", str(origem))
    monkeypatch.setenv("PLUGIN_NANOCAD_DESTINO", str(destino))
    return origem, destino


def _res(r):
    return {x["alvo"]: x for x in r}


def test_instala_os_dois_e_depois_fica_em_dia(pastas):
    _, destino = pastas
    r = _res(inst.instalar())
    assert r["nanocad26"]["resultado"] == "instalado" and r["nanocad5"]["resultado"] == "instalado"
    assert (destino / "nanoCAD26" / "OrcamentoNanoCAD.dll").read_bytes() == b"MZ 26 nova"
    assert json.loads((destino / "nanoCAD5" / "versao.json").read_text(encoding="utf-8"))["versao"] == "0.10.0"
    assert all(x["resultado"] == "em dia" for x in inst.instalar())
    s = _res(inst.situacao())
    assert s["nanocad26"]["instalado"] and not s["nanocad26"]["desatualizado"]


def test_versao_mais_velha_e_trocada_e_a_anterior_fica_guardada(pastas):
    _, destino = pastas
    d = destino / "nanoCAD26"
    d.mkdir(parents=True)
    (d / "OrcamentoNanoCAD.dll").write_bytes(b"MZ velha")
    (d / "versao.json").write_text(json.dumps({"versao": "0.9.5"}), encoding="utf-8")       # 0.9.5 < 0.10.0 (número a número)
    assert _res(inst.situacao())["nanocad26"]["desatualizado"]
    r = _res(inst.instalar())["nanocad26"]
    assert r["resultado"] == "instalado" and r["anterior"] == "0.9.5"
    assert (d / "OrcamentoNanoCAD.dll").read_bytes() == b"MZ 26 nova"
    assert (d / "OrcamentoNanoCAD.dll.anterior").read_bytes() == b"MZ velha"
    assert not (d / "OrcamentoNanoCAD.dll.nova").exists()


def test_nao_volta_versao_que_o_plugin_ja_atualizou_sozinho(pastas):
    _, destino = pastas
    d = destino / "nanoCAD26"
    d.mkdir(parents=True)
    (d / "OrcamentoNanoCAD.dll").write_bytes(b"MZ do plugin")
    (d / "versao.json").write_text(json.dumps({"versao": "0.11.0", "origem": "atualização automática do plugin"}), encoding="utf-8")
    assert _res(inst.instalar())["nanocad26"]["resultado"] == "em dia"
    assert (d / "OrcamentoNanoCAD.dll").read_bytes() == b"MZ do plugin"


def test_rotas_no_desktop_e_no_servidor(pastas, monkeypatch):
    c = TestClient(appmod.app)
    h = {"Authorization": f"Bearer {create_jwt_token('u1', 'u1@x.com', 'operador')}"}
    assert c.get("/api/plugin/instalacao").status_code == 401
    monkeypatch.setattr(config, "APP_MODE", "desktop")
    j = c.get("/api/plugin/instalacao", headers=h).json()
    assert j["desktop"] and {a["alvo"] for a in j["alvos"]} == {"nanocad26", "nanocad5"}
    assert c.post("/api/plugin/instalacao", json={}, headers=h).json()["resultado"][0]["resultado"] == "instalado"
    monkeypatch.setattr(config, "APP_MODE", "server")
    j = c.get("/api/plugin/instalacao", headers=h).json()
    assert not j["desktop"] and "alvos" not in j
    assert c.post("/api/plugin/instalacao", json={}, headers=h).status_code == 403


def test_build_bat_e_release_levam_as_duas_dlls():
    bat = (RAIZ / "build.bat").read_text(encoding="utf-8")
    assert "plugin_nanocad\\appload\\nanoCAD26" in bat and "plugin_nanocad\\appload\\nanoCAD5" in bat
    wf = (RAIZ / ".github" / "workflows" / "release-exe.yml").read_text(encoding="utf-8")
    assert "OrcamentoNanoCAD_nanoCAD5.dll" in wf and "OrcamentoNanoCAD_nanoCAD26.dll" in wf
