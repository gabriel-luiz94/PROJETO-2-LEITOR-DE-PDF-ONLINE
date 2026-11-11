"""TASK-136 — sugestão de ativos/apelidos do PROJETO ao digitar (com botão liga/desliga) e ativo de outro projeto não vale no cálculo."""
import json
import os

import pytest

import config
from services import ativo_parametros as ap
from services import orcamento_calc as calc
from tests import oraculo_tela as o

JS = os.path.join(config.STATIC_DIR, "ativo_sugestao.js")
node = pytest.mark.skipif(not o.NODE, reason="node não instalado")


def L(ativo, proj="", origem="", comp="", cod="", fi=1.0):
    return {"ativo": ativo, "desc_ativo": "d", "componente": comp, "projeto": proj, "mdo": "MATERIAL", "codigo": cod, "desc_codigo": "d", "fator_i": fi, "fator_r": 0.0,
            "filtro": "", "origem": origem, "lv": "A", "arred": "NORMAL", "formula_i": "", "formula_r": ""}


BASE = [L("CHAVE_FACA=CHAVE_FACA(cabo=S04), SUH", "PERNAMBUCO"), L("CHAVE_FACA", "PERNAMBUCO", comp="CHAVE_FACA", cod="500109", fi=3), L("SUPORTE_PB", "PARAIBA", comp="SPB", cod="700"),
        L("SUA", "", "OUTROS", comp="SUA_C", cod="701"), L("[S336,4]", "PERNAMBUCO", comp="S3", cod="702"), L("N1C:=N1(C)+N2(C), NC", "PARAIBA"), L("ZZZ", "OUTRA", "CABOS", comp="Z", cod="9")]


def js(codigo, entrada):
    import subprocess
    r = subprocess.run([o.NODE, "-e", f"const S=require({json.dumps(JS)});const e=JSON.parse(require('fs').readFileSync(0,'utf8'));console.log(JSON.stringify({codigo}));"],
                       input=json.dumps(entrada), capture_output=True, text=True, timeout=30)
    assert r.returncode == 0, r.stderr
    return json.loads(r.stdout)


# ───────────────────────── nomes do projeto ─────────────────────────
@node
def test_so_ativos_e_apelidos_do_projeto():
    def nomes(proj, origem="OUTROS"):
        return js("S.indiceDeNomes(e.base,e.proj,e.origem).map(x=>[x.nome,x.tipo,x.de])", {"base": BASE, "proj": proj, "origem": origem})
    pe = nomes("PERNAMBUCO")
    assert ["SUH", "apelido", "CHAVE_FACA"] in pe and ["CHAVE_FACA", "ativo", None] in pe and ["SUA", "ativo", None] in pe and ["S336,4", "ativo", None] in pe
    assert not any(n[0] in ("SUPORTE_PB", "NC", "N1C", "ZZZ") for n in pe)                           # outro projeto / origem CABOS: não aparece
    pb = [n[0] for n in nomes("PARAIBA")]
    assert "SUH" not in pb and "CHAVE_FACA" not in pb and {"SUPORTE_PB", "SUA", "N1C", "NC"} <= set(pb)
    assert [n[0] for n in nomes("")] == ["SUA"]                                                       # sem projeto: só as genéricas
    assert "ZZZ" not in [n[0] for n in nomes("OUTRA", "OUTROS")]                                      # origem CABOS não entra em Outros


@node
def test_sugerir_ordem_e_limite():
    r = js("S.sugerir(e.q,S.indiceDeNomes(e.base,e.proj,'OUTROS')).map(x=>x.nome)", {"base": BASE, "proj": "PERNAMBUCO", "q": "SU"})
    assert r == ["SUA", "SUH"]
    assert js("S.sugerir('UH',S.indiceDeNomes(e.base,e.proj,'OUTROS')).map(x=>x.nome)", {"base": BASE, "proj": "PERNAMBUCO"}) == ["SUH"]       # contém, não só começa
    assert js("S.sugerir('SUH',S.indiceDeNomes(e.base,e.proj,'OUTROS'))", {"base": BASE, "proj": "PERNAMBUCO"}) == []                        # idêntico e sozinho: nada a sugerir
    assert js("S.sugerir('',[])", {}) == []
    grande = [{"nome": f"A{i:03d}", "texto": f"A{i:03d}", "tipo": "ativo", "de": None} for i in range(40)]
    assert len(js("S.sugerir('A',e)", grande)) == 12


@node
@pytest.mark.parametrize("texto,cursor,esperado", [
    ("SU", 2, {"inicio": 0, "fim": 2, "query": "SU"}), ("9400D, SU", 9, {"inicio": 7, "fim": 9, "query": "SU"}), ("2-SU", 4, {"inicio": 2, "fim": 4, "query": "SU"}),
    ("*1-SU", 5, {"inicio": 3, "fim": 5, "query": "SU"}), ("1-N1 1-SU", 9, {"inicio": 7, "fim": 9, "query": "SU"}), ("2 SU", 4, {"inicio": 2, "fim": 4, "query": "SU"}),
    ("N1(D, X", 7, None), ("2", 1, None), ("2 ", 2, None), ("", 0, None), ("[S33", 4, None), ("SUH", 2, {"inicio": 0, "fim": 3, "query": "SU"})])
def test_token_atual(texto, cursor, esperado):
    assert js("S.tokenAtual(e.t,e.c)", {"t": texto, "c": cursor}) == esperado


@node
def test_aplicar_troca_so_o_item():
    t = js("S.tokenAtual(e.t,e.c)", {"t": "9400D, SU, N1", "c": 9})
    r = js("S.aplicar(e.t,e.tok,{texto:'SUH'})", {"t": "9400D, SU, N1", "tok": t})
    assert r == {"texto": "9400D, SUH, N1", "cursor": 10}
    r = js("S.aplicar(e.t,e.tok,{texto:'[S336,4]'})", {"t": "S3", "tok": {"inicio": 0, "fim": 2, "query": "S3"}})
    assert r["texto"] == "[S336,4]"


# ───────────────────────── cálculo: ativo de outro projeto não vale ─────────────────────────
def total(texto, proj):
    r = calc.processar_calculo([], [{"ativo": texto, "operacao": "I", "entidade": "0"}], proj, ap.preencher_componentes(BASE)[0])
    return {x["codigo"]: x["total"] for x in r["resultado"]}, r["nao_encontrados"]


def test_ativo_de_outro_projeto_vira_nao_encontrado():
    assert total("1-CHAVE_FACA", "PERNAMBUCO") == ({"500109": 3.0}, [])
    assert total("1-SUH", "PERNAMBUCO")[0] == {"500109": 3.0}
    assert total("1-CHAVE_FACA", "PARAIBA") == ({}, ["CHAVE_FACA"])                  # antes: "achado" sem itens
    assert total("1-SUH", "PARAIBA") == ({}, ["SUH"])
    assert total("1-SUA", "PARAIBA")[0] == {"701": 1.0}                              # genérico vale em qualquer projeto
    assert total("1-SUPORTE_PB", "PERNAMBUCO") == ({}, ["SUPORTE_PB"])


# ───────────────────────── tela ─────────────────────────
def _ler(nome):
    with open(os.path.join(config.STATIC_DIR, nome), encoding="utf-8") as f:
        return f.read()


def test_tela_liga_a_sugestao_e_desliga_a_antiga_em_outros():
    html, rjs = _ler("resumo.html"), _ler("resumo.js")
    assert 'id="chip-sugestoes"' in html and "/static/ativo_sugestao.js" in html
    assert "AtivoSugestao.iniciar" in rjs and "if (type !== 'outros') showAutocomplete(inpAt, type);" in rjs and "e.key === 'ArrowDown' && type !== 'outros'" in rjs
    js_ = _ler("ativo_sugestao.js")
    assert "sugestao_ativos" in js_ and "getItem(CHAVE) !== '0'" in js_                  # padrão ligado, guarda a escolha
