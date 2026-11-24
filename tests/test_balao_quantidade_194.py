"""TASK-194 — balão de parâmetros: com quantidade no ativo («6 CON», «6-CON», «*2 CON», «11300D, 2 N4») o Enter agora GRAVA o parâmetro.
A leitura procurava a dica pelo nome SEM a quantidade e a gravação, pelo nome COM ela («6 CON» não existe na base): navegava e não gravava."""
import json
import os
import shutil
import subprocess

import pytest

import config

NODE = shutil.which("node")
pytestmark = pytest.mark.skipif(not NODE, reason="node indisponível")
S = config.STATIC_DIR

ROTEIRO = r"""
const path = require('path');
const AI = require(path.join(%(s)s, 'ativo_itens.js')), AD = require(path.join(%(s)s, 'ativo_dica.js'));
const e = JSON.parse(require('fs').readFileSync(0, 'utf8'));
const pedidos = [];
const dicaDe = nome => { pedidos.push(nome); return nome === 'CON' || nome === 'N4' ? { tipo: 'parametros', params: (nome === 'CON' ? ['cabo', 'cabo2', 'mdo', 'tipo'] : ['poste', 'tipocabo', 'cabo']).map(n => ({ nome: n, valores: [] })) } : { tipo: 'sem_parametros' }; };
console.log(JSON.stringify(e.map(c => AI.gravarParametro(c.texto, c.i, c.param, c.valor, dicaDe, AD.aplicarValor)).concat([pedidos])));
"""


def _gravar(casos):
    r = subprocess.run([NODE, "-e", ROTEIRO % {"s": json.dumps(S)}], input=json.dumps(casos), capture_output=True, text=True, timeout=30)
    assert r.returncode == 0, r.stderr
    saida = json.loads(r.stdout)
    return saida[:-1], saida[-1]


def test_grava_com_e_sem_quantidade_mantendo_a_quantidade_no_texto():
    r, pedidos = _gravar([
        {"texto": "CON", "i": 0, "param": "cabo", "valor": "A70"},
        {"texto": "6 CON", "i": 0, "param": "cabo", "valor": "A70"},
        {"texto": "6-CON", "i": 0, "param": "cabo", "valor": "A70"},
        {"texto": "*2 CON", "i": 0, "param": "tipo", "valor": "COMP"},
        {"texto": "0,5 CON", "i": 0, "param": "mdo", "valor": "N"}])
    assert r == ["CON(cabo=A70)", "6 CON(cabo=A70)", "6-CON(cabo=A70)", "*2 CON(tipo=COMP)", "0,5 CON(mdo=N)"]
    assert pedidos == ["CON"] * 5                       # a base é perguntada pelo nome SEM a quantidade (antes: «6 CON», que não existe)


def test_linha_com_varios_itens_so_muda_o_item_escolhido():
    r, _ = _gravar([
        {"texto": "11300D, 2 N4", "i": 1, "param": "tipocabo", "valor": "A"},
        {"texto": "3 CON, 2 CON", "i": 1, "param": "cabo", "valor": "S04"},
        {"texto": "3 CON(cabo=A70), 2 CON", "i": 0, "param": "cabo2", "valor": "A70"},
        {"texto": "6 CON(cabo=A70,cabo2=A70)", "i": 0, "param": "cabo", "valor": "S04"}])          # troca o valor de um parâmetro já escrito
    assert r == ["11300D, 2 N4(tipocabo=A)", "3 CON, 2 CON(cabo=S04)", "3 CON(cabo=A70,cabo2=A70), 2 CON", "6 CON(cabo=S04,cabo2=A70)"]


def test_limpar_um_parametro_e_casos_sem_gravacao():
    r, _ = _gravar([
        {"texto": "6 CON(cabo=A70,mdo=N)", "i": 0, "param": "cabo", "valor": ""},
        {"texto": "6 PR", "i": 0, "param": "cabo", "valor": "A70"},          # ativo sem parâmetros declarados
        {"texto": "6 CON", "i": 3, "param": "cabo", "valor": "A70"},         # item que não existe
        {"texto": "", "i": 0, "param": "cabo", "valor": "A70"}])
    assert r == ["6 CON(mdo=N)", None, None, None]


def test_resumo_usa_a_gravacao_nova_e_so_um_balao_e_iniciado():
    js = open(os.path.join(S, "resumo.js"), encoding="utf-8").read()
    assert js.count("window.AtivoBalao.iniciar(") == 1                      # um corte de texto errado já duplicou este bloco uma vez
    i = js.index("gravar: (input, i, nomeParam, valor)")
    corpo = js[i:js.index("carregarBase:", i)]
    assert "_gravarParametroNoTexto(row.ativo, i, nomeParam, valor)" in corpo           # TASK-249: a gravação foi para uma função (o balão do mapa usa a mesma)
    j = js.index("function _gravarParametroNoTexto(")
    funcao = js[j:js.index("\n}", j)]
    assert "AtivoItens.gravarParametro(" in funcao and "it.nome, c" not in funcao


def test_quantidade_decimal_com_virgula_nao_quebra_o_item():
    js = "const AI = require(%s); console.log(JSON.stringify([AI.tokens('0,5 CON, 1,5 N3(XLPE,D), 2 PR').map(t => t.bruto), AI.itens('0,5 CON, 11300D').map(i => [i.nome, i.base]), AI.tokens('A,B, C').map(t => t.bruto), AI.compacto('0,5 CON(cabo=A70), 2 N1(XLPE,D)')]))" % json.dumps(os.path.join(S, "ativo_itens.js"))
    r = subprocess.run([NODE, "-e", js], capture_output=True, text=True, timeout=30)
    assert r.returncode == 0, r.stderr
    tokens, itens, sem_espaco, compacto = json.loads(r.stdout)
    assert tokens == ["0,5 CON", "1,5 N3(XLPE,D)", "2 PR"]
    assert itens == [["0,5 CON", "CON"], ["11300D", "11300D"]]
    assert sem_espaco == ["A", "B", "C"]                 # vírgula entre LETRAS continua separando, como sempre
    assert compacto == "0,5 CON, 2 N1"
