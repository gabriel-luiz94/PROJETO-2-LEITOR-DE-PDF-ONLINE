"""TASK-084 — o modal de vinculação em abas: arquivos ligados na tela, sintaxe e contratos que o resumo.js usa."""
import os
import shutil
import subprocess

import pytest

import config

NODE = shutil.which("node")


def _ler(nome):
    with open(os.path.join(config.STATIC_DIR, nome), encoding="utf-8") as f:
        return f.read()


def test_scripts_na_ordem_certa_no_resumo():
    html = _ler("resumo.html")
    ordem = [html.index(f"/static/{n}.js") for n in ("anotacao_outros", "anotacao_cabos", "vinculacao_regras", "vinculacao_ui", "resumo")]
    assert ordem == sorted(ordem)


@pytest.mark.skipif(not NODE, reason="node indisponível")
@pytest.mark.parametrize("arq", ["vinculacao_ui.js", "vinculacao_regras.js", "resumo.js"])
def test_sintaxe_dos_arquivos(arq):
    r = subprocess.run([NODE, "--check", os.path.join(config.STATIC_DIR, arq)], capture_output=True, text=True)
    assert r.returncode == 0, r.stderr


def test_resumo_liga_o_host_e_a_totalizadora_ao_vinculo():
    js = _ler("resumo.js")
    for trecho in ["window.VinculacaoHost = {", "window.VinculacaoUI.abrir(window.VinculacaoHost)", "window.abrirModalConectores", "_parametrosVinculoPorLinha()",
                   "parametros_linha: parametrosVinculo[idxOutros]", "obj.cabosFicticios", "obj.vinculoOrigem", "/api/vinculacao-config?projeto_codigo=", "VinculacaoRegras.achados("]:
        assert trecho in js, trecho
    ui = _ler("vinculacao_ui.js")
    for metodo in ["host.cabos()", "host.outros()", "host.antes()", "host.desfazer()", "host.projeto()", "host.prefixoCabo", "host.codigoEstrutura", "host.postesDaLinha", "host.linhaSoPoste"]:
        assert metodo in ui, metodo
    for fn in ["cabos", "outros", "projeto", "prefixoCabo", "codigoEstrutura", "postesDaLinha", "linhaSoPoste", "antes", "desfazer", "depois"]:
        assert fn in js.split("window.VinculacaoHost = {")[1].split("window.abrirModalVinculacao")[0], fn


def test_modal_antigo_so_tem_conectores():
    html = _ler("resumo.html")
    assert "body-vinc-conectores" in html and "vinculacao-lista" not in html and "body-vinc-regras" not in html
