"""TASK-158 — Salvar / Salvar como / salvamento automático ligado-desligado / rascunho do projeto default."""
import json
import os

import pytest

import config
from services import projeto_desktop as pd, projeto_obra as po

DADOS = {"cabos": {"data": [{"entidade": "0", "operacao": "I", "ativo": "C35", "qtdAtivos": 16}]}, "outros": {"data": [{"entidade": "0", "operacao": "I", "ativo": "11300D"}]}}


def ler(*partes):
    return open(os.path.join(*partes), encoding="utf-8").read()


# ───────── rascunho do projeto default ─────────
def test_rascunho_grava_rotaciona_e_le(tmp_path):
    pasta = str(tmp_path / "rascunhos")
    assert pd.rascunho_info(pasta) == {"ok": True, "rascunhos": []}                              # ainda não há rascunho
    r = pd.rascunho_gravar({"dados": DADOS}, "081", True, pasta)
    assert r["ok"] and os.path.basename(r["caminho"]) == "rascunho_default.obra" and not os.path.exists(r["caminho"] + ".tmp")
    info = pd.rascunho_info(pasta)["rascunhos"]
    assert [x["id"] for x in info] == ["atual"] and info[0]["itens"] == 2 and info[0]["salvo_em"]
    # mesma sessão: continua substituindo o atual (não cria "anterior")
    assert pd.rascunho_gravar({"dados": {"cabos": [], "outros": [{"operacao": "I", "ativo": "N1"}]}}, "081", False, pasta)["ok"]
    assert [x["id"] for x in pd.rascunho_info(pasta)["rascunhos"]] == ["atual"]
    # sessão nova: o atual vira "anterior"
    assert pd.rascunho_gravar({"dados": DADOS}, "081", True, pasta)["ok"]
    info = pd.rascunho_info(pasta)["rascunhos"]
    assert [x["id"] for x in info] == ["atual", "anterior"] and info[1]["itens"] == 1
    lido = pd.rascunho_ler("anterior", "081", pasta)
    assert lido["ok"] and lido["projeto"]["dados"]["outros"][0]["ativo"] == "N1" and lido["projeto"]["nome"] == "Rascunho (projeto default)"
    assert pd.rascunho_ler("atual", "999", pasta)["ok"] is False                              # outro projeto de regras: recusa com mensagem
    assert pd.rascunho_ler("xyz", "081", pasta)["ok"] is False


def test_rascunho_fica_numa_pasta_junto_do_app():
    assert pd.pasta_rascunho() == os.path.join(config.BASE_DIR, "rascunhos")


def test_rascunho_aceita_dados_vazios_e_nao_levanta_excecao(tmp_path):
    assert pd.rascunho_gravar({}, "081", False, str(tmp_path / "a"))["ok"] is True
    assert pd.rascunho_gravar({"dados": DADOS}, "", False, str(tmp_path / "b"))["ok"] is False   # sem código do projeto de regras o arquivo não seria válido


def test_ponte_expoe_os_metodos_de_rascunho():
    api = pd.ApiProjeto()
    assert api.projeto_rascunho_info()["ok"] is True or api.projeto_rascunho_info()["ok"] is False
    assert api.projeto_rascunho_gravar("{", "081")["ok"] is False and api.projeto_rascunho_ler("xyz")["ok"] is False


# ───────── telas ─────────
def test_barra_com_salvar_salvar_como_e_automatico():
    js = ler(config.STATIC_DIR, "projeto.js")
    for trecho in ("btn-projeto-salvar", "btn-projeto-menu", "Salvar como…", "Salvamento automático", "Ctrl+Shift+S", "Recuperar rascunho do default…",
                   "gridr_auto_default", "auto: !!auto", "Salvar alterações em", "beforeunload", "alterado por fora — clique para resolver", "Não salvo", "atualizarEstado"):
        assert trecho in js, trecho
    assert "auto: copia ? !!(atual && atual.auto === true) : true" in js          # projeto novo: ligado; salvar como mantém
    assert "definirProjeto(r, false)" in js                                          # projeto aberto: desligado
    assert "if (!estado()) { abrirJanela('novo'); return Promise.resolve(); }" in js   # no default, Salvar abre Novo projeto
    assert "if (!resumo.temDados())" not in js.split("async function gravarRascunho")[1].split("async function recuperarRascunho")[0] or "!resumo.temDados()) return" in js
    assert "nunca grava a tela vazia" in js


def test_resultado_registra_coletor_e_nao_guarda_assinatura_velha():
    js = ler(config.STATIC_DIR, "projeto.js")
    assert "__registrarColetor(entregar)" in js and "__pendenteResultado" in js
    assert "atualizarEstado({ assinatura: r.assinatura })" in js                     # a assinatura é sempre atualizada sobre o estado ATUAL (sem corrida)
    assert "e.assinatura = r.assinatura; gravarEstado(e)" not in js
