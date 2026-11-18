"""TASK-188 — Leitor: vizinhança POR POSIÇÃO no desenho (DXF). O comprimento de um cabo é o texto «NN m» mais perto (abaixo) do rótulo, não o próximo da lista
(a lista sai ordenada por altura: o comprimento caía dezenas de itens adiante ou atrás e o cabo ficava sem comprimento ou com o de outro cabo)."""
import json
import os
import shutil
import subprocess

import pytest

import config
from routers import regras_leitor as rl

NODE = shutil.which("node")
MOTOR = os.path.join(config.STATIC_DIR, "regras_leitor_engine.js")

LAYERS = ["POSTE_INSTALANDO_LV", "POSTE_REMOVENDO_LV"]
LABEL = r"^(?!\d+(?:[.,]\d+)?\s*m$)((?:(?:C\s*)?(?:\d[xX]){0,2}[A-Z]{0,4}\d[\dA-Z/,.]*(?:\([^)]*\))?(?:\+\d+(?:[.,]\d+)?)?(?:\s*[-+]\s*\d+(?:[.,]\d+)?\s*KV)?))$"
COMP = r"^(\d+(?:[.,]\d+)?)\s*m$"


def regra(**viz):
    v = {"regex": COMP, "janela": 16, "mesma_pagina": True, "direcao": "DEPOIS"}
    v.update(viz)
    return {"fase": 1, "ordem": 30, "ativa": True, "modo": "DEFINIR", "layer_em": LAYERS, "texto_regex": LABEL, "ativo_template": "{1} {vizinhanca} m", "operacao": "", "parar": True, "vizinhanca": v}


def it(texto, x, y, layer="POSTE_INSTALANDO_LV"):
    return {"texto": texto, "cor": "#ffffff", "layer": layer, "pagina": 1, "_x": x, "_y": y}


_JS = r"""
const M = require(%(m)s); const e = JSON.parse(require('fs').readFileSync(0, 'utf8'));
globalThis.window = {};
console.log(JSON.stringify(e.casos.map(c => c.itens.map((x, i) => M.processarEClassificar({ texto: x.texto, cor: x.cor, layer: x.layer, pagina: x.pagina, index: i, allItems: c.itens }, [c.regra], []).ativo))));
"""


def roda(casos):
    r = subprocess.run([NODE, "-e", _JS % {"m": json.dumps(MOTOR)}], input=json.dumps({"casos": casos}), capture_output=True, text=True, timeout=30)
    assert r.returncode == 0, r.stderr
    return json.loads(r.stdout)


# o desenho do usuário: rótulo de cima, dois riscados (um MTEXT) e o 27 m logo abaixo; no meio da lista entram textos de OUTROS pontos (ordem por altura)
def desenho():
    lista = [it("3xA35P(7,94)-13,8kV", 100, 40.6), it("2xC25-13,8kV", 100, 38.0, "POSTE_REMOVENDO_LV"), it("1xS04-13,8kV", 100, 38.0, "POSTE_REMOVENDO_LV")]
    lista += [it(f"RUA {k}", 500 + k, 37.0 - k * 0.1, "TXT_RUAS") for k in range(20)]                  # 20 textos de outros lugares com altura entre os rótulos e o comprimento
    lista += [it("27 m", 100, 33.6, "0"), it("36 m", 10, 38.0, "0")]                                  # o 36 m (de outro cabo) está na mesma altura, 90 unidades à esquerda
    return lista


@pytest.mark.skipif(not NODE, reason="node indisponível")
def test_ordem_da_lista_erra_e_a_posicao_acerta():
    por_ordem, por_posicao = roda([{"itens": desenho(), "regra": regra()}, {"itens": desenho(), "regra": regra(geometrica=True, raio=12, posicao="ABAIXO")}])
    assert por_ordem[0] != "3xA35P(7,94)-13,8kV 27 m" and por_ordem[1] != "2xC25-13,8kV 27 m"             # pela ordem da lista não acha o 27 m (nem o certo)
    assert por_posicao[0] == "3xA35P(7,94)-13,8kV 27 m"
    assert por_posicao[1] == "2xC25-13,8kV 27 m" and por_posicao[2] == "1xS04-13,8kV 27 m"                  # os três dividem o mesmo comprimento


@pytest.mark.skipif(not NODE, reason="node indisponível")
def test_raio_posicao_e_sem_coordenadas():
    base = [it("3xA35P(7,94)-13,8kV", 0, 100), it("30 m", 0, 95, "0"), it("20 m", 0, 105, "0")]
    r = roda([
        {"itens": base, "regra": regra(geometrica=True, raio=12)},                                  # abaixo: 30 m (o 20 m está acima)
        {"itens": base, "regra": regra(geometrica=True, raio=12, posicao="ACIMA")},                 # acima: 20 m
        {"itens": base, "regra": regra(geometrica=True, raio=3)},                                   # raio curto: nada dentro dele
        {"itens": [it("3xA35P(7,94)-13,8kV", 0, 100), it("30 m", 0, 50, "0")], "regra": regra(geometrica=True, raio=12)},   # longe demais: não casa (não cai para a ordem)
        {"itens": [{"texto": "3xA35P(7,94)-13,8kV", "cor": "#fff", "layer": "POSTE_INSTALANDO_LV", "pagina": 1}, {"texto": "30 m", "cor": "#fff", "layer": "0", "pagina": 1}], "regra": regra(geometrica=True)},   # PDF, sem _x/_y: usa a ordem
    ])
    assert r[0][0] == "3xA35P(7,94)-13,8kV 30 m"
    assert r[1][0] == "3xA35P(7,94)-13,8kV 20 m"
    assert r[2][0] == ""                                                                            # nada dentro do raio: esta regra não monta o ativo (a regra dos itens das layers POSTE_* completa depois)
    assert r[3][0] == ""                                                                            # longe demais: não casa (e não cai para a ordem da lista)
    assert r[4][0] == "3xA35P(7,94)-13,8kV 30 m"


@pytest.mark.skipif(not NODE, reason="node indisponível")
def test_o_empate_fica_com_o_mais_perto_na_lista_e_a_pagina_conta():
    itens = [it("3xA35P(7,94)-13,8kV", 0, 100), it("30 m", 5, 98, "0"), it("40 m", -5, 98, "0")]
    outra = [it("3xA35P(7,94)-13,8kV", 0, 100), dict(it("30 m", 0, 98, "0"), pagina=2)]
    r = roda([{"itens": itens, "regra": regra(geometrica=True)}, {"itens": outra, "regra": regra(geometrica=True)}])
    assert r[0][0] == "3xA35P(7,94)-13,8kV 30 m"                                                    # mesma distância: o mais perto na lista
    assert r[1][0] == ""                                                                            # outra página não vale


@pytest.mark.skipif(not NODE, reason="node indisponível")
def test_rotulo_com_neutro_mais_n():
    itens = [it("3xC35(35)+35", 0, 100), it("29 m", 0, 97, "0"), it("3xC16(16)+16-13,8kV", 50, 100), it("11 m", 50, 97, "0")]
    r = roda([{"itens": itens, "regra": regra(geometrica=True)}])
    assert r[0][0] == "3xC35(35)+35 29 m" and r[0][2] == "3xC16(16)+16-13,8kV 11 m"


def _erros(**viz):
    erros = []
    rl._valida_regra_processamento({"fase": 1, "ordem": 1, "modo": "DEFINIR", "texto_regex": "x", "vizinhanca": {"regex": "^a$", "janela": 5, **viz}}, 1, erros)
    return erros


def test_validacao_dos_campos_novos():
    assert _erros(geometrica=True, raio=12, posicao="ABAIXO") == []
    assert _erros(geometrica=True) == []
    assert any("geometrica" in e for e in _erros(geometrica="sim"))
    for ruim in (0, -3, "12", True):
        assert any("raio" in e for e in _erros(geometrica=True, raio=ruim)), ruim
    assert any("posicao" in e for e in _erros(geometrica=True, posicao="DIREITA"))


def test_editor_de_regras_tem_os_campos_por_posicao():
    s = open(os.path.join(config.STATIC_DIR, "script.js"), encoding="utf-8").read()
    for trecho in ["id=\"rle-viz-geo\"", "id=\"rle-viz-raio\"", "id=\"rle-viz-posicao\"", "regra.vizinhanca.geometrica = true", "regra.vizinhanca.raio = parseFloat"]:
        assert trecho in s, trecho
    m = open(MOTOR, encoding="utf-8").read()
    assert "function _vizinhoPorPosicao(" in m and "r.vizinhanca.geometrica" in m
