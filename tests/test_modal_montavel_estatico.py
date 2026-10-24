"""TASK-077 — ligações estáticas dos modais montáveis (HTML, rotas de JS, host em resumo.js)."""
import os

import pytest
from fastapi.testclient import TestClient

import config

HTML = open(os.path.join(config.STATIC_DIR, "resumo.html"), encoding="utf-8").read()
RESUMO = open(os.path.join(config.STATIC_DIR, "resumo.js"), encoding="utf-8").read()
UI = open(os.path.join(config.STATIC_DIR, "modal_montavel_ui.js"), encoding="utf-8").read()


def test_scripts_antes_do_resumo_e_container_na_barra():
    i_motor = HTML.index("/static/modal_montavel_engine.js")
    i_ui = HTML.index("/static/modal_montavel_ui.js")
    assert HTML.index("/static/grade_tabela.js") < i_motor < i_ui < HTML.index("/static/resumo.js")
    assert 'id="barra-modais-montaveis"' in HTML
    assert HTML.index("abrirModalVinculacao()") < HTML.index('id="barra-modais-montaveis"')


def test_troca_de_projeto_recarrega_os_modais():
    assert HTML.count("ModalMontavelUI.recarregar()") >= 2   # onchange do select e carga inicial


def test_host_escreve_nas_duas_tabelas_com_historico():
    bloco = RESUMO[RESUMO.index("window.ModalMontavelHost"):]
    bloco = bloco[:bloco.index("ModalMontavelUI.iniciar")]
    assert "pushHistory()" in bloco and "tableStates.cabos.data.push" in bloco and "tableStates.outros.data.push" in bloco
    assert "it.qtd + '-' + it.ativo" in bloco


def test_ui_so_admin_ve_o_mais_e_nao_usa_eval():
    assert "ehAdmin()" in UI and "new Function" not in UI and "eval(" not in UI
    assert "innerHTML = '" not in UI.replace("innerHTML = ''", "")   # nunca injeta HTML vindo da definição


@pytest.mark.parametrize("arq", ["modal_montavel_engine.js", "modal_montavel_ui.js"])
def test_js_servido_sem_cache(arq):
    import app as appmod
    r = TestClient(appmod.app).get("/static/" + arq)
    assert r.status_code == 200
    assert "no-cache" in r.headers.get("cache-control", "").lower()
