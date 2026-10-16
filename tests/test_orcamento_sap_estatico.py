"""TASK-070 — a visão SAP está ligada na tela: alternador, módulo carregado antes do script da página, rota sem cache."""
import os

import pytest
from fastapi.testclient import TestClient

import config
import database


@pytest.fixture
def client(tmp_path, monkeypatch):
    monkeypatch.setattr(database, "DB_PATH", str(tmp_path / "teste.db"))
    database.init_db()
    import app as appmod
    return TestClient(appmod.app)


def _html():
    with open(os.path.join(config.STATIC_DIR, "resultado_orcamento.html"), encoding="utf-8") as f:
        return f.read()


def test_alternador_e_exportar_excel_na_tela():
    h = _html()
    assert "definirVisao('rec')" in h and "definirVisao('sap')" in h
    assert 'id="btn-exportar-xlsx"' in h and 'id="sap-grid"' in h
    assert h.index("/static/orcamento_sap.js") < h.index("const tbody = document.getElementById('resultado-orcamento-tbody')")


def test_o_filtro_da_sap_nao_toca_nas_linhas_do_rec():
    """A SAP filtra só as suas linhas: o filtro do REC (que esconde linhas e, por isso, muda o .rec) fica desligado no modo SAP."""
    h = _html()
    assert "if (modoVisao === 'sap') { aplicarFiltrosSap(); return; }" in h


def test_rota_do_modulo_sem_cache(client):
    r = client.get("/static/orcamento_sap.js")
    assert r.status_code == 200 and "no-cache" in r.headers.get("cache-control", "").lower()


def test_modo_comparacao_ligado_na_tela_e_rota_sem_cache(client):
    """TASK-071: barra, botões, coluna REFERÊNCIA antes de QUANTIDADE e fantasmas fora do .rec."""
    h = _html()
    assert 'id="barra-comparar"' in h and "compararComRecBanco(" in h and 'id="file-rec-comparar"' in h
    assert h.index("/static/orcamento_comparar.js") < h.index("const tbody = document.getElementById('resultado-orcamento-tbody')")
    assert h.index("REFERÊNCIA ↕") < h.index("QUANTIDADE ↕")                       # coluna só de visualização, antes da quantidade
    assert "tr.classList.contains('fantasma')" in h                                # leitores do REC ignoram as linhas-fantasma
    r = client.get("/static/orcamento_comparar.js")
    assert r.status_code == 200 and "no-cache" in r.headers.get("cache-control", "").lower()
