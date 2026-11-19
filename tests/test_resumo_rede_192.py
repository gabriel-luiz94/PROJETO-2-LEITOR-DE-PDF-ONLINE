"""TASK-192 — Resumo da rede em Pernambuco: poste compacto (nome oficial da base = `12600D`), linha viva (`*I`) também instala; Paraíba segue igual."""
import json
import os
import shutil
import subprocess

import pytest

import config

NODE = shutil.which("node")
pytestmark = pytest.mark.skipif(not NODE, reason="node indisponível")
S = config.STATIC_DIR

# Como o resumo.js chama: as três funções de cabo/Outros que ele passa ao módulo saem do próprio resumo.js (corte por chaves), para o teste usar o código real.
ROTEIRO = r"""
const fs = require('fs'), path = require('path');
const S = %(static)s;
const AnotacaoOutros = require(path.join(S, 'anotacao_outros.js')), AnotacaoCabos = require(path.join(S, 'anotacao_cabos.js')), RR = require(path.join(S, 'resumo_rede.js'));
const src = fs.readFileSync(path.join(S, 'resumo.js'), 'utf8');
const pegar = nome => { const i = src.indexOf('function ' + nome + '('); let j = src.indexOf('{', i), n = 0; for (let k = j; k < src.length; k++) { if (src[k] === '{') n++; else if (src[k] === '}' && --n === 0) return src.slice(i, k + 1); } };
eval(['extrairParesQtdAtivoOutros', 'extrairPrefixoEComprimentoCabo', 'extrairPrefixoCaboSemComprimento'].map(pegar).join('\n'));
const e = JSON.parse(fs.readFileSync(0, 'utf8'));
const ctxOutros = AnotacaoOutros.contextoDaBase(e.base || [], e.projeto || '', undefined);
const r = RR.calcular({ cabos: e.cabos || [], outros: e.outros || [], ctxOutros, convencaoPostes: undefined, AnotacaoOutros, AnotacaoCabos, extrairPares: extrairParesQtdAtivoOutros,
                        extrairPrefixoEComprimento: extrairPrefixoEComprimentoCabo, extrairPrefixoSemComprimento: extrairPrefixoCaboSemComprimento });
console.log(JSON.stringify(r));
"""


def _calc(entrada):
    r = subprocess.run([NODE, "-e", ROTEIRO % {"static": json.dumps(S)}], input=json.dumps(entrada), capture_output=True, text=True, timeout=30)
    assert r.returncode == 0, r.stderr
    return json.loads(r.stdout)


def _o(op, ativo, **kw):
    return {"operacao": op, "ativo": ativo, **kw}


# A base de Pernambuco cadastra o poste com o nome oficial COMPACTO e o DT/RC como apelido (como no banco real).
BASE_PE = [{"ativo": a, "projeto": "PERNAMBUCO", "codigo": "1", "desc_ativo": "poste", "componente": "P"} for a in
           ("12600D=12600D(fases=TRI), DT12/600", "12600D(fases=TRI)", "121000R=121000R(fases=TRI), RC12/1000", "121000R(fases=TRI)", "11300D=11300D(fases=TRI), DT11/300", "9600R=9600R(fases=TRI), RC9/600")]


def test_pernambuco_poste_compacto_com_base_conta():
    r = _calc({"projeto": "PERNAMBUCO", "base": BASE_PE, "outros": [
        _o("I", "121000R, CE4, CE3, FLBIT"), _o("I", "12600D, CEJ2, TR, FLABIT"), _o("R", "11300D, B1-MT")]})
    assert r["postes"] == 2 and r["equipamentos"] == 1                       # o R não conta; o TR do poste de trafo é equipamento


def test_linha_viva_instala_tambem():
    r = _calc({"projeto": "PERNAMBUCO", "base": BASE_PE, "outros": [_o("*I", "12600D, CEJ2, TR"), _o("*R", "11300D, CFU"), _o("M", "9600R"), _o("*M", "9600R"), _o("", "9600R")]})
    assert r["postes"] == 1 and r["equipamentos"] == 1                       # só o *I; *R, M, *M e vazia não instalam


def test_poste_compacto_sem_base_e_com_parametros_e_quantidade():
    r = _calc({"projeto": "PERNAMBUCO", "base": [], "outros": [_o("I", "12600D(fases=TRI), TR"), _o("I", "2 9600R"), _o("I", "12600D, 2 CFU, ITF-R, ATRINT, CRUZETA 2,4 m")]})
    assert r["postes"] == 4 and r["equipamentos"] == 3                       # 1 + 2 + 1 postes; TR + 2 CFU


def test_itens_que_nao_sao_poste_nao_contam_como_poste():
    r = _calc({"projeto": "PERNAMBUCO", "base": BASE_PE, "outros": [_o("I", "CE4, B4-MT, IT-R, N3-2, FLABIT, ATRINT-EQ, PR, TT")]})
    assert r["postes"] == 0 and r["equipamentos"] == 0


def test_paraiba_continua_igual_formato_antigo():
    r = _calc({"projeto": "PARAIBA", "base": [], "outros": [_o("I", "DT11/300 1-CFU 1-TR125"), _o("I", "2-DT10/300 1-TR115"), _o("R", "DT11/300 1-CFU"), _o("*I", "CV11/600 1-CFA")]})
    assert r["postes"] == 4 and r["equipamentos"] == 4          # DT11/300 + 2 DT10/300 + CV11/600 (o *I agora conta); CFU + TR125 + TR115 + CFA


def test_cabos_pernambuco_media_baixa_e_linha_viva():
    r = _calc({"cabos": [
        _o("I", "3xA35P(7,94)-13,8kV 25 m"),          # média: o vão conta uma vez (25); o neutro 7,94 não entra
        _o("*I", "3xA185P(7,94)-13,8kV 10 m"),         # linha viva também: +10
        _o("R", "3xA35P(7,94)-13,8kV 99 m"),           # retirada não conta
        _o("I", "3xC16(16) 20 m"),                     # sem kV: baixa (o cabo 20 + o neutro 20)
        _o("M", "3xA35P(7,94)-13,8kV 50 m")]})
    assert r["redeMedia"] == 25 + 10 and r["redeBaixa"] == 40


def test_cabos_classico_paraiba_media_e_baixa():
    r = _calc({"cabos": [_o("I", "CAA2 ABC 35 m"), _o("I", "M3X1X35+35 40 m"), _o("R", "CAA2 ABC 99 m")]})
    assert r["redeMedia"] == 35 and r["redeBaixa"] == 40


def test_comprimento_na_linha_de_baixo_continua():
    r = _calc({"cabos": [_o("I", "3xA35P(7,94)-13,8kV"), _o("I", "45 m")]})
    assert r["redeMedia"] == 45


def test_instalando_e_sem_parametros():
    RR = os.path.join(S, "resumo_rede.js")
    js = "const R = require(%s); console.log(JSON.stringify([['I','*I',' *i ','R','*R','M','*M','',null].map(R.instalando), R.semParametros('12600D(fases=TRI)'), R.semParametros(null)]))" % json.dumps(RR)
    r = subprocess.run([NODE, "-e", js], capture_output=True, text=True, timeout=30)
    assert json.loads(r.stdout) == [[True, True, True, False, False, False, False, False, False], "12600D", ""]


def test_resumo_usa_o_modulo_e_e_servido_sem_cache():
    js, html = open(os.path.join(S, "resumo.js"), encoding="utf-8").read(), open(os.path.join(S, "resumo.html"), encoding="utf-8").read()
    corpo = js[js.index("function calcularResumoRede"):js.index("function formatarNumeroResumoRede")]
    assert "ResumoRede.calcular" in corpo and "__vinculacaoConfig.postes" in corpo
    assert html.index("/static/resumo_rede.js") < html.index("/static/resumo.js")
    from fastapi.testclient import TestClient
    import app as appmod
    r = TestClient(appmod.app).get("/static/resumo_rede.js")
    assert r.status_code == 200 and "no-store" in r.headers["cache-control"]
