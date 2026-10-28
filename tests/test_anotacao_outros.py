"""TASK-081 — notação compacta na tabela Outros ("11600D, 2 N3, T3"): tradução implícita, apelidos, regra do M e paridade com o JS."""
import json
import os
import random
import shutil
import subprocess

import pytest

import config
from services import anotacao_outros as ao
from services import ativo_parametros as ap
from services import orcamento_calc as calc
from services.autonomo import totalizadora as tt
from services.validacao_planilhas import validar_planilhas
from tests import oraculo_tela as o

NODE = shutil.which("node")
cl = calc._classico_outros


def L(ativo, comp="", cod="", fi=1, proj="PERNAMBUCO", lv="A", mdo="MATERIAL"):
    return {"ativo": ativo, "desc_ativo": ativo and "D", "componente": comp, "projeto": proj, "mdo": mdo, "codigo": cod, "desc_codigo": f"d {cod}",
            "fator_i": fi, "fator_r": 0.0, "filtro": "", "origem": "OUTROS", "lv": lv}


BASE = [
    L("DT11/600, 11600D", "DT11/600", "p600"), L("DT11/300, 11300D", "DT11/300", "p300"), L("RC10/300", "RC10/300", "rc300"), L("FV7/150", "FV7/150", "fv150"),
    L("N3=N3(poste=11600D)", "N3", "n3a"), L("", "N3", "n3b", fi=2),
    L("N3(poste=11600D)", "N3.11600D", "n3esp"),
    L("T3", "T3", "t3"),
]


def p(texto, op="I"):
    return {"ativo": texto, "operacao": op, "entidade": "0"}


def itens(texto, op="I", ctx=None):
    r = ao.interpretar_linha(texto, op, cl, ctx)
    return [(i["ativo"], i["qtd"], i["operacao"]) for i in r["itens"]], r["erros"]


# ── postes em notação compacta ──
@pytest.mark.parametrize("compacto,nome", [("11600D", "DT11/600"), ("11300D", "DT11/300"), ("10300R", "RC10/300"), ("9300F", "CV9/300"), ("7150FV", "FV7/150"),
                                           ("91000D", "DT9/1000"), ("12400d", "DT12/400")])
def test_compacto_para_poste(compacto, nome):
    r = ao.compacto_para_poste(compacto)
    assert r["ativo"] == nome and r["erro"] is None
    assert ao.poste_para_compacto(nome) == compacto.upper()


def test_compacto_ambiguo_ou_impossivel():
    assert ao.compacto_para_poste("N3") is None and ao.compacto_para_poste("T3") is None
    assert ao.compacto_para_poste("123D")["erro"]                     # 12/3 e 1/23: sem esforço de 3-4 dígitos plausível
    assert ao.compacto_para_poste("11600D", {"DT1/1600", "DT11/600"})["ativo"] == "DT11/600"      # prefere o que existe na base


# ── formato e vírgulas ──
def test_compacto_basico():
    assert itens("11600D, 2 N3, T3") == ([("DT11/600", 1.0, "I"), ("N3", 2.0, "I"), ("T3", 1.0, "I")], [])
    assert itens("2 11600D, 0,5 N3")[0] == [("DT11/600", 2.0, "I"), ("N3", 0.5, "I")]
    assert itens("*2 N3, T3")[0] == [("N3", -2.0, "I"), ("T3", 1.0, "I")]
    assert ao.eh_compacta("11600D, N3") and not ao.eh_compacta("0,5-N3") and not ao.eh_compacta("1-N1(XLPE,DT)")


def test_formato_antigo_continua_igual():
    assert itens("2-N3 1-T3")[0] == [("N3", 2.0, "I"), ("T3", 1.0, "I")]
    assert itens("2-N1(XLPE,DT)")[0] == [("N1(XLPE,DT)", 2.0, "I")]
    assert itens("0,5-N3")[0] == [("N3", 0.5, "I")]


def test_erros_de_formato():
    assert itens("11600D,N1")[1][0]["cod"] == "VIRGULA"
    assert itens("11600D, 1-N3")[1][0]["cod"] == "MISTURA"
    assert itens("11600D, 2 3 N3") == ([("DT11/600", 1.0, "I"), ("3 N3", 2.0, "I")], [])        # o único delimitador é ", ": "3 N3" é o nome (com espaço)
    assert itens("11600D, ")[1] != [] or itens("11600D, ")[0]          # vírgula no fim = item vazio
    assert itens("11600D, x N3") == ([("DT11/600", 1.0, "I"), ("X N3", 1.0, "I")], [])


# ── apelidos e regra do poste/M ──
def test_apelidos_e_poste_na_estrutura():
    ctx = ao.contexto_da_base(BASE, "PERNAMBUCO")
    assert ctx["apelidos"] == {"11600D": "DT11/600", "11300D": "DT11/300"}
    assert itens("11300D, 2 N3", ctx=ctx)[0] == [("DT11/300", 1.0, "I"), ("N3(POSTE=11300D)", 2.0, "I")]
    assert itens("11600D, N3, 11300D, N3, T3", ctx=ctx)[0][1::2][0][0] == "N3(POSTE=11600D)"
    r = itens("11600D, N3, 11300D, N3", ctx=ctx)[0]
    assert [x[0] for x in r] == ["DT11/600", "N3(POSTE=11600D)", "DT11/300", "N3(POSTE=11300D)"]       # cada estrutura usa o poste mais próximo à esquerda
    assert itens("N3, T3", ctx=ctx)[0] == [("N3", 1.0, "I"), ("T3", 1.0, "I")]                      # sem poste à esquerda: padrão


def test_operacao_M_ignora_poste_e_demais_ficam_I():
    ctx = ao.contexto_da_base(BASE, "PERNAMBUCO")
    assert itens("11600D, N3, T3", "M", ctx)[0] == [("N3(POSTE=11600D)", 1.0, "I"), ("T3", 1.0, "I")]
    assert itens("11600D, N3", "*M", ctx)[0] == [("N3(POSTE=11600D)", 1.0, "*I")]
    assert itens("1-DT11/600 1-N3", "M", ctx)[0] == [("N3(POSTE=11600D)", 1.0, "I")]               # vale também no formato antigo


def test_texto_equivalente_e_codigo_da_estrutura():
    ctx = ao.contexto_da_base(BASE, "PERNAMBUCO")
    assert ao.texto_equivalente("11600D, 2 N3", cl, ctx) == "1-DT11/600 2-N3"
    assert ao.codigo_estrutura("11600D, 2 N3", cl, ctx) == "N3"
    assert ao.texto_equivalente("2-N3", cl, ctx) == "2-N3"


# ── cálculo ponta a ponta ──
def totais(res):
    t = {}
    for r in res["resultado"]:
        t[r["codigo"]] = t.get(r["codigo"], 0) + r["total"]
    return {k: round(v, 2) for k, v in t.items()}


def test_calculo_compacto_igual_ao_antigo_equivalente():
    a = calc.processar_calculo([], [p("11600D, 2 N3, T3")], "PERNAMBUCO", BASE)
    b = calc.processar_calculo([], [p("1-DT11/600 2-N3(POSTE=11600D) 1-T3")], "PERNAMBUCO", BASE)
    assert totais(a) == totais(b) and a["nao_encontrados"] == []
    assert totais(a) == {"p600": 1, "n3a": 2, "n3b": 4, "n3esp": 2, "t3": 1}


def test_calculo_M_nao_conta_o_poste():
    r = calc.processar_calculo([], [p("11600D, N3", "M")], "PERNAMBUCO", BASE)
    assert "p600" not in totais(r) and totais(r)["n3esp"] == 1


def test_calculo_apelido_em_linha_antiga():
    r = calc.processar_calculo([], [p("2-11300D")], "PERNAMBUCO", BASE)
    assert totais(r) == {"p300": 2} and r["nao_encontrados"] == []


# ── apelidos na base ──
def test_base_apelidos_validacao():
    a = ap.analisar_celula_base("DT11/300, 11300D")
    assert (a["base"], a["apelidos"], a["erro"]) == ("DT11/300", ["11300D"], None)
    for celula in ("N1=N1(cabo=XLPE), X", "N1(cabo=XLPE), X"):                                   # TASK-102: apelido também vale em ativo com parâmetros
        r = ap.analisar_celula_base(celula)
        assert r["erro"] is None and r["apelidos"] == ["X"] and r["base"] == "N1", celula
    assert ap.analisar_celula_base("B1MT=B1MT(poste=%D,tipocabo=A), B1-MT")["apelidos"] == ["B1-MT"]       # hífen aceito no apelido
    assert ap.analisar_celula_base("DT9/150, a b")["erro"]
    assert ap.nome_base("DT11/300, 11300D") == "DT11/300"
    erros = ap.validar_base([{"ativo": "DT11/300, X1", "projeto": ""}, {"ativo": "DT9/150, X1", "projeto": ""}])["erros"]
    assert any("X1" in e for e in erros)
    erros = ap.validar_base([{"ativo": "DT11/300, DT9/150", "projeto": ""}, {"ativo": "DT9/150", "projeto": ""}])["erros"]
    assert any("DT9/150" in e for e in erros)


# ── validação (camada 1) ──
def test_validacao_camada1():
    def cods(texto):
        return {a["regra_id"] for a in validar_planilhas([], [p(texto)], BASE, "PERNAMBUCO")}
    assert cods("11600D, 2 N3, T3") == set()
    assert "C1-OUT-VIRGULA" in cods("11600D,N3")
    assert "C1-OUT-MISTURA" in cods("11600D, 1-N3")
    assert cods("2-N3 1-T3") == set()
    assert "C1-DUP" in cods("11600D, N3, N3")


# ── Totalizadora: o espelho Python continua igual ao JS real, agora com notação compacta ──
BASE_T = [{"ativo": "DT11/600, 11600D", "desc_ativo": "POSTE 11/600", "componente": "DT11/600", "projeto": "P1", "codigo": "p600", "desc_codigo": "d"},
          {"ativo": "N3=N3(poste=11600D)", "desc_ativo": "N3", "componente": "N3", "projeto": "P1", "codigo": "n3", "desc_codigo": "d"},
          {"ativo": "N3(poste=11600D)", "desc_ativo": "", "componente": "N3.11600D", "projeto": "P1", "codigo": "n3e", "desc_codigo": "d"},
          {"ativo": "T3", "desc_ativo": "T3", "componente": "T3", "projeto": "P1", "codigo": "t3", "desc_codigo": "d"},
          {"ativo": "CE3 PR", "desc_ativo": "CE3 PR", "componente": "CE3 PR", "projeto": "P1", "codigo": "ce3pr", "desc_codigo": "d"}]
LINHAS_T = ["CE3 PR", "2 CE3 PR", "11600D, CE3 PR, T3", "11600D, 2 CE3 PR", "11600D, 2 N3, T3", "2 11600D, N3", "0,5 T3, N3", "11600D,N3", "11600D, 1-N3", "2-N3 1-T3", "DT11/600 1-N3", "11600D, *2 N3", "9999X, N3", "N3, T3", "11600D, N3(A)", "11600D, N3(A,11600D)", "11600D", "9300F", "7150FV"]


@pytest.mark.skipif(not NODE, reason="node indisponível")
@pytest.mark.parametrize("operacao", ["I", "*I", "R", "M", "*M"])
def test_totalizadora_compacta_igual_ao_js(operacao):
    outros = [{"entidade": "ESTRUTURA", "operacao": operacao, "ativo": t} for t in LINHAS_T]
    js = o.totalizadora_e_payload([], outros, [], BASE_T, "P1")
    tot = tt.montar_totalizadora([], outros, [], BASE_T, "P1")
    assert tt.normalizar_para_json(tot) == js["totalizadora"]
    assert tt.payload_calculo(tot) == js["payload"]
    assert len(tot) > 5


# ── paridade do módulo (Python x JS) ──
_JS = """
const fs = require('fs');
globalThis.AtivoParametros = require(%(ap)s);
const AO = require(%(ao)s);
const e = JSON.parse(fs.readFileSync(0, 'utf8'));
const cl = t => { const s = []; AtivoParametros.dividirTokens(t).forEach(x => { const m = x.match(/^([\\*\\-]?\\d+(?:[.,]\\d+)?)[Xx\\-](.+)$/i); if (m) s.push({ativo: m[2].trim().toUpperCase(), qtd: parseFloat(m[1].replace(',', '.'))}); else s.push({ativo: x.toUpperCase(), qtd: 1}); }); return s; };
const ctx = AO.contextoDaBase(e.base, e.projeto);
console.log(JSON.stringify(e.casos.map(c => AO.interpretarLinha(c[0], c[1], cl, ctx))));
"""


def _classico_py(texto):
    out = []
    import re
    for x in ap.dividir_tokens(texto):
        m = re.match(r"^([*-]?\d+(?:[.,]\d+)?)[Xx-](.+)$", x)
        out.append({"ativo": (m.group(2).strip() if m else x).upper(), "qtd": float(m.group(1).replace("*", "-").replace(",", ".")) if m else 1.0})
    return out


@pytest.mark.skipif(not NODE, reason="node indisponível")
def test_paridade_python_js_do_modulo():
    casos = [[t, op] for t in LINHAS_T + ["7150FV, N3", "10300R, T3, 12400D, N3", "11600D, N3(poste=11600D)", "", "  ", "CE3 PR", "9400D, CE3 PR, N1", "9400D, 2 CE3 PR", "2 CE3 PR", "1-N1 2-CE3", "N1 N2"] for op in ("I", "*I", "M", "*M")]
    cod = _JS % {"ap": json.dumps(os.path.join(config.STATIC_DIR, "ativo_parametros.js")), "ao": json.dumps(os.path.join(config.STATIC_DIR, "anotacao_outros.js"))}
    r = subprocess.run([NODE, "-e", cod], input=json.dumps({"casos": casos, "base": BASE_T, "projeto": "P1"}), capture_output=True, text=True, timeout=60)
    assert r.returncode == 0, r.stderr
    js = json.loads(r.stdout)
    ctx = ao.contexto_da_base(BASE_T, "P1")
    py = [ao.interpretar_linha(t, op, _classico_py, ctx) for t, op in casos]
    assert py == js


def test_arquivos_publicados():
    from fastapi.testclient import TestClient
    r = TestClient(__import__("app").app).get("/static/anotacao_outros.js")
    assert r.status_code == 200 and "interpretarLinha" in r.text


def test_base_de_exemplo_e_valida():
    assert ap.validar_base([dict(l) for l in BASE])["erros"] == []


def test_poste_tambem_em_pedido_com_parenteses():
    ctx = ao.contexto_da_base([L("N1=N1(cabo=A, poste=11300D)", "N1", "x")], "PERNAMBUCO")
    f = lambda t: itens(t, ctx=ctx)[0][-1][0]
    assert f("7150FV, N1(P)") == "N1(P,POSTE=7150FV)"
    assert f("7150FV, N1(cabo=P)") == "N1(CABO=P,POSTE=7150FV)"
    assert f("7150FV, N1(P,11300D)") == "N1(P,11300D)"                  # já informou o poste: respeita
    assert f("7150FV, N1(poste=11300D)") == "N1(POSTE=11300D)"


# ── TASK-097: o único delimitador é ", " — ativo com espaço no nome é UM item ──
def test_nome_com_espaco_e_um_item_so():
    assert itens("9400D, CE3 PR")[0] == [("DT9/400", 1.0, "I"), ("CE3 PR", 1.0, "I")]
    assert itens("9400D, 2 CE3 PR, N1")[0] == [("DT9/400", 1.0, "I"), ("CE3 PR", 2.0, "I"), ("N1", 1.0, "I")]         # "2" é a quantidade; o resto é o nome
    assert itens("CE3 PR, N1")[0] == [("CE3 PR", 1.0, "I"), ("N1", 1.0, "I")]


def test_linha_sozinha_sem_quantidade_com_hifen_e_compacta():
    r = ao.interpretar_linha("CE3 PR", "I", cl, None)
    assert r["formato"] == "compacto" and [(i["ativo"], i["qtd"]) for i in r["itens"]] == [("CE3 PR", 1.0)]
    assert itens("2 CE3 PR")[0] == [("CE3 PR", 2.0, "I")]
    assert itens("N1")[0] == [("N1", 1.0, "I")]
    # com "qtd-ativo" em algum token continua o formato antigo (itens separados por espaço)
    r = ao.interpretar_linha("1-N1 2-CE3", "I", cl, None)
    assert r["formato"] == "antigo" and [(i["ativo"], i["qtd"]) for i in r["itens"]] == [("N1", 1.0), ("CE3", 2.0)]
