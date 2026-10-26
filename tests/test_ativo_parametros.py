"""TASK-079 — ativos com parâmetros por categoria (services/ativo_parametros.py e static/ativo_parametros.js, com paridade)."""
import json
import os
import subprocess

import pytest

import config
from services import ativo_parametros as ap
from tests import oraculo_tela as o

JS = os.path.join(config.STATIC_DIR, "ativo_parametros.js")

BASE = ["N1=N1(cabo=XLPE, poste=DT)", "N1", "N1(cabo=XLPE)", "N1(cabo=ALUM)", "N1(cabo=COBRE)", "N1(poste=DT)", "N1(poste=CC)",
        "N1(cabo=COBRE, poste=CC)", "P50", "DT9/150"]
SEM_PADRAO = ["N1=N1(cabo, poste)", "N1(cabo=XLPE)", "N1(cabo=ALUM)", "N1(poste=DT)", "N1(poste=CC)"]
SEM_DECL = ["N1(cabo=XLPE)", "N1(poste=DT)", "N1(cabo=ALUM)"]

TEXTOS = ["N1", "n1", " N1 ( xlpe , dt ) ", "N1(XLPE,DT)", "N1(ALUM)", "N1(poste=CC, cabo=COBRE)", "N1(XLIPE)", "N1(ALUM,XX)", "N1(A,B,C)", "N1(cabo=ALUM,cabo=XLPE)",
          "N1(cabo=XLPE,DT)", "N1(XLPE,poste=DT)", "N1()", "N1(,)", "N1(XL-PE)", "N1(cabo=)", "N1(=X)", "N1(XLPE", "N1)(", "P50(X)", "P50", "DT9/150", "N1=N1(a=b)",
          "N1=X1(a=b)", "N1=N1", "", "  ", "N1(cabo=XLPE)(poste=DT)", "N1(XLPE;DT)"]
CELULAS = ["N1=N1(cabo=XLPE, poste=DT)", "N1=N1(cabo, poste)", "N1=N1(cabo=XLPE, poste)", "N1(cabo=XLPE)", "N1=N1(a, a)", "N1=N1()", "N1=X(a=b)", "N1(XLPE)", "P50", ""]


def node(codigo, entrada):
    r = subprocess.run([o.NODE, "-e", f"const A=require({json.dumps(JS)});const e=JSON.parse(require('fs').readFileSync(0,'utf8'));console.log(JSON.stringify({codigo}));"],
                       input=json.dumps(entrada), capture_output=True, text=True, timeout=30)
    assert r.returncode == 0, r.stderr
    return json.loads(r.stdout)


# ── Python ──
def test_analisar_formas():
    a = ap.analisar_celula_base("N1=N1(cabo=XLPE, poste=DT)")
    assert (a["tipo"], a["base"], a["erro"]) == ("declaracao", "N1", None)
    assert a["params"] == [{"nome": "CABO", "op": "=", "valor": "XLPE", "alts": ["XLPE"]}, {"nome": "POSTE", "op": "=", "valor": "DT", "alts": ["DT"]}]
    assert ap.analisar_celula_base("N1=N1(cabo, poste)")["params"] == [{"nome": "CABO", "op": "=", "valor": None, "alts": []}, {"nome": "POSTE", "op": "=", "valor": None, "alts": []}]
    assert ap.analisar_celula_base("N1(cabo=XLPE)")["tipo"] == "grupo"
    assert ap.analisar_celula_base("DT9/150") == {"texto": "DT9/150", "base": "DT9/150", "tipo": "simples", "params": [], "erro": None}
    for ruim in ["N1(XL-PE)", "N1()", "N1(cabo=)", "N1(", "N1=X1(a=b)", "N1=N1", "N1(a=1,a=2)", ""]:
        assert ap.analisar_celula_base(ruim)["erro"], ruim


def test_nome_base():
    assert [ap.nome_base(x) for x in ["N1(XLPE,DT)", "n1=N1(cabo=XLPE)", "N1", " DT9/150 ", "N1(cabo=A)"]] == ["N1", "N1", "N1", "DT9/150", "N1"]


def test_assinatura_com_declaracao_e_padroes():
    s = ap.montar_assinaturas(BASE)["N1"]
    assert s["nomes"] == ["CABO", "POSTE"] and s["padroes"] == {"CABO": "XLPE", "POSTE": "DT"} and s["declarada"]
    assert s["valores"] == {"CABO": ["XLPE", "ALUM", "COBRE"], "POSTE": ["DT", "CC"]}
    assert "P50" not in ap.montar_assinaturas(BASE)


def test_assinatura_sem_padrao_e_sem_declaracao():
    s = ap.montar_assinaturas(SEM_PADRAO)["N1"]
    assert s["nomes"] == ["CABO", "POSTE"] and s["padroes"] == {}
    s2 = ap.montar_assinaturas(SEM_DECL)["N1"]
    assert s2["nomes"] == ["CABO", "POSTE"] and not s2["declarada"] and s2["padroes"] == {}       # ordem = 1ª aparição


def test_resolver_com_padrao():
    sig = ap.montar_assinaturas(BASE)
    f = lambda t: ap.resolver_pedido(t, sig)
    assert f("N1")["texto"] == "N1(XLPE,DT)" and f("n1(alum)")["texto"] == "N1(ALUM,DT)"
    assert f("N1(poste=CC)")["texto"] == "N1(XLPE,CC)" and f("N1(poste=CC, cabo=COBRE)")["texto"] == "N1(COBRE,CC)"
    assert f("N1(ALUM,CC)")["params"] == {"CABO": "ALUM", "POSTE": "CC"}
    assert f("N1(XLIPE)")["erros"] and "XLPE, ALUM, COBRE" in f("N1(XLIPE)")["erros"][0]
    assert f("N1(A,B,C)")["erros"] and f("N1(cabo=X,cabo=Y)")["erros"]
    assert f("P50") == {"base": "P50", "texto": "P50", "params": {}, "faltando": [], "erros": [], "com_assinatura": False}
    assert f("P50(X)")["erros"] == ["P50 não tem variantes cadastradas"]


def test_resolver_sem_padrao_incompleto():
    sig = ap.montar_assinaturas(SEM_PADRAO)
    r = ap.resolver_pedido("N1", sig)
    assert r["faltando"] == ["CABO", "POSTE"] and r["texto"] == "N1" and not r["erros"]
    r = ap.resolver_pedido("N1(ALUM)", sig)
    assert r["faltando"] == ["POSTE"] and r["texto"] == "N1(CABO=ALUM)" and r["params"] == {"CABO": "ALUM"}
    r2 = ap.resolver_pedido(r["texto"], sig)                                   # forma incompleta é estável
    assert r2["texto"] == r["texto"] and r2["faltando"] == ["POSTE"]
    ok = ap.resolver_pedido("N1(poste=DT, cabo=ALUM)", sig)
    assert ok["texto"] == "N1(ALUM,DT)" and ok["faltando"] == []


def test_linha_vale_igualdade_simples():
    assert ap.linha_vale({}, {"CABO": "XLPE"}) and ap.linha_vale({"CABO": "XLPE"}, {"CABO": "XLPE", "POSTE": "DT"})
    assert not ap.linha_vale({"CABO": "ALUM"}, {"CABO": "XLPE"}) and not ap.linha_vale({"POSTE": "DT"}, {"CABO": "XLPE"})
    assert ap.linha_vale({"CABO": "XLPE", "POSTE": "CC"}, {"CABO": "XLPE", "POSTE": "CC"})


def test_dividir_tokens():
    assert ap.dividir_tokens("3-N1(XLPE, DT) *2-P50  1-N1") == ["3-N1(XLPE, DT)", "*2-P50", "1-N1"]
    assert ap.dividir_tokens("a (b c) d") == ["a", "(b c)", "d"] and ap.dividir_tokens("") == [] and ap.dividir_tokens("x((y z)") == ["x((y z)"]


# ── paridade Python x JS ──
pytestmark_node = pytest.mark.skipif(not o.NODE, reason="node não instalado")


@pytestmark_node
def test_paridade_analisar_e_nome_base():
    esperado = [[ap.analisar_ativo(t), ap.analisar_celula_base(t), ap.nome_base(t)] for t in TEXTOS + CELULAS]
    obtido = node("e.map(t=>[A.analisarAtivo(t), A.analisarCelulaBase(t), A.nomeBase(t)])", TEXTOS + CELULAS)
    assert obtido == esperado


@pytestmark_node
def test_paridade_assinaturas_e_pedidos():
    for base in (BASE, SEM_PADRAO, SEM_DECL, CELULAS):
        esperado_sig = ap.montar_assinaturas(base)
        assert node("A.montarAssinaturas(e)", base) == esperado_sig
        esperado = [ap.resolver_pedido(t, esperado_sig) for t in TEXTOS]
        assert node("e.t.map(x=>A.resolverPedido(x, e.s))", {"t": TEXTOS, "s": esperado_sig}) == esperado


@pytestmark_node
def test_paridade_tokens_e_linha_vale():
    textos = ["3-N1(XLPE, DT) *2-P50  1-N1", "", "a (b c) d", "x((y z)", "  ", "N1(A,B)"]
    assert node("e.map(A.dividirTokens)", textos) == [ap.dividir_tokens(t) for t in textos]
    casos = [[{}, {"A": "1"}], [{"A": "1"}, {"A": "1", "B": "2"}], [{"A": "2"}, {"A": "1"}], [{"B": "1"}, {}]]
    assert node("e.map(c=>A.linhaVale(c[0], c[1]))", casos) == [ap.linha_vale(a, b) for a, b in casos]


@pytestmark_node
def test_assinaturas_do_projeto_filtra_por_projeto():
    linhas = [{"ativo": "N1(cabo=XLPE)", "projeto": "PERNAMBUCO"}, {"ativo": "N1(cabo=ALUM)", "projeto": "RONDONIA"}, {"ativo": "N1(poste=DT)", "projeto": ""}]
    s = node("A.assinaturasDoProjeto(e, 'pernambuco')", linhas)
    assert s["N1"]["valores"] == {"CABO": ["XLPE"], "POSTE": ["DT"]}


# ── validação da base ──
def L(ativo="", comp="", cod="", proj="PERNAMBUCO", lv="A"):
    return {"ativo": ativo, "componente": comp, "codigo": cod, "projeto": proj, "lv": lv}


def cadastro_ok():
    return [L("N1=N1(cabo=XLPE, poste=DT)", "N1", "c1"), L("", "N1", "c2"),
            L("N1(cabo=XLPE)", "N1.XLPE", "x1"), L("", "N1.XLPE", "x2"),
            L("N1(cabo=ALUM)", "N1.ALUM", "a1"), L("", "N1.ALUM", "a2"),
            L("N1(poste=DT)", "N1.DT", "d1"), L("N1(poste=CC)", "N1.CC", "k1"),
            L("P50", "P", "p1")]


def test_validar_base_cadastro_correto_e_base_sem_parametros():
    r = ap.validar_base(cadastro_ok())
    assert r == {"erros": [], "avisos": []}
    assert ap.validar_base([L("P50", "P", "p1"), L("", "P", "p2"), L("DT9/150", "D", "d1")]) == {"erros": [], "avisos": []}


def test_validar_base_erros():
    def erros(linhas):
        return " | ".join(ap.validar_base(linhas)["erros"])
    base = cadastro_ok()
    assert "mais de uma declaração" in erros(base + [L("N1=N1(cabo=XLPE, poste=DT)", "N1X", "z1")])
    assert "primeira linha do ativo" in erros([L("N1", "N1", "c1")] + base[:1] + base[2:])
    assert "não está na declaração" in erros(base + [L("N1(tubo=X)", "N1.T", "t1")])
    assert "padrão" in erros([L("N1=N1(cabo=ZZZ, poste=DT)", "N1", "c1")] + base[1:])
    assert "levam nome" in erros(base + [L("N1(XLPE)", "N1.Z", "z1")])
    assert "inválido" in erros(base + [L("N1(cabo=XL-PE)", "N1.Z", "z1")])
    assert "COMPONENTE N1.XLPE é usado por grupos diferentes" in erros(base + [L("N1(cabo=COBRE)", "N1.XLPE", "z1")])
    assert "COMPONENTE N1 é usado" in erros(base + [L("N1(cabo=COBRE)", "N1", "z1")])


def test_validar_base_aviso_de_codigo_repetido_em_grupos_que_valem_juntos():
    base = cadastro_ok() + [L("", "N1.XLPE", "d1")]                      # d1 também está em N1.DT (poste=DT) — XLPE e DT valem juntos
    av = ap.validar_base(base)["avisos"]
    assert len(av) == 1 and "D1" in av[0] and "N1.XLPE" in av[0] and "N1.DT" in av[0]
    exclusivos = cadastro_ok() + [L("", "N1.ALUM", "x1")]               # x1 em XLPE e ALUM (mesmo parâmetro, valores diferentes): nunca juntos
    assert ap.validar_base(exclusivos)["avisos"] == []
    lv = cadastro_ok()[:2] + [L("N1(poste=CC)", "N1.CC", "m1", lv="LM"), L("N1(poste=DT)", "N1.DT", "m1", lv="LV")]
    assert ap.validar_base(lv)["avisos"] == []


def test_validar_base_escopo_por_projeto():
    outra = [dict(l, projeto="RONDONIA") for l in cadastro_ok()]
    assert ap.validar_base(cadastro_ok() + outra) == {"erros": [], "avisos": []}                 # mesma declaração em projetos diferentes


# ═══════════════ TASK-080: operadores nos grupos, componente automático ═══════════════
BASE80 = ["N1=N1(cabo=XLPE, poste=DT)", "N1", "N1(cabo=XLPE)", "N1(cabo=ALUM|COBRE)", "N1(cabo=ALUM)", "N1(cabo=COBRE)", "N1(poste=D%)", "N1(poste=C%)",
          "N1(cabo!=XLPE, poste=D%)"]
TEXTOS80 = TEXTOS + ["N1(ALUM,D12)", "N1(COBRE,CC)", "N1(XLPE,CC)", "N1(ALUM,XX)", "N1(poste=D%)", "N1(cabo!=XLPE)", "N1(A|B)", "N1(ALUM,D12,Z)", "N1(poste=cc, cabo=cobre)"]
CELULAS80 = CELULAS + ["N1(poste=D%)", "N1(cabo=ALUM|COBRE)", "N1(cabo!=XLPE, poste=D%)", "N1(cabo!=ALUM|COBRE)", "N1(poste=%CC%)", "N1(poste=%)", "N1(cabo=A||B)",
                       "N1(cabo=A B)", "N1=N1(cabo=ALUM|COBRE)", "N1=N1(cabo=D%)", "N1=N1(cabo!=X)", "N1(cabo!=)", "N1(!=X)", "N1(poste=D%,poste=C%)", "N1(cabo=X%Y%Z)"]


def test_analisar_operadores_so_na_base():
    a = ap.analisar_celula_base("N1(poste=D%, cabo!=ALUM|COBRE)")
    assert a["erro"] is None
    assert a["params"] == [{"nome": "POSTE", "op": "=", "valor": "D%", "alts": ["D%"]}, {"nome": "CABO", "op": "!=", "valor": "ALUM|COBRE", "alts": ["ALUM", "COBRE"]}]
    for pedido in ["N1(D%)", "N1(poste=D%)", "N1(A|B)", "N1(cabo!=X)"]:
        assert "só valem nos grupos da base" in ap.analisar_ativo(pedido)["erro"], pedido
    for decl in ["N1=N1(cabo=ALUM|COBRE)", "N1=N1(cabo=D%)", "N1=N1(cabo!=X)"]:
        assert ap.analisar_celula_base(decl)["erro"], decl
    for ruim in ["N1(cabo=A||B)", "N1(cabo!=)", "N1(!=X)", "N1(poste=D%,poste=C%)"]:
        assert ap.analisar_celula_base(ruim)["erro"], ruim
    assert ap.analisar_celula_base("N1(poste=%CC%)")["erro"] is None and ap.analisar_celula_base("N1(poste=%)")["erro"] is None


def test_casar_padrao_e_restricao():
    assert [ap.casar_padrao(p, v) for p, v in [("D%", "DT"), ("D%", "D"), ("D%", "XD"), ("%T", "DT"), ("%CC%", "ACCD"), ("D%T", "DAAT"), ("d%", "dt"),
                                              ("A%B", "A_B"), ("A.B", "AXB"), ("D%", "")]] == [True, True, False, True, True, True, True, True, False, False]
    r = {"op": "!=", "alts": ["ALUM", "COBRE"]}
    assert ap.restricao_vale(r, "XLPE") and not ap.restricao_vale(r, "alum") and ap.restricao_vale({"op": "=", "alts": ["A", "B%"]}, "BX")
    assert ap.restricao_vale("DT", "dt") and not ap.restricao_vale("DT", "DC")


def test_linha_vale_com_restricoes():
    g = ap.restricoes_do_grupo(ap.analisar_celula_base("N1(poste=D%, cabo!=XLPE)"))
    assert ap.linha_vale(g, {"CABO": "ALUM", "POSTE": "D12"}) and not ap.linha_vale(g, {"CABO": "XLPE", "POSTE": "D12"})
    assert not ap.linha_vale(g, {"CABO": "ALUM", "POSTE": "C1"}) and not ap.linha_vale(g, {"CABO": "ALUM"})
    assert ap.restricoes_do_grupo(ap.analisar_celula_base("N1")) == {} and ap.linha_vale({}, {})


def test_assinatura_com_operadores_e_valores_aceitos():
    s = ap.montar_assinaturas(BASE80)["N1"]
    assert s["nomes"] == ["CABO", "POSTE"] and s["padroes"] == {"CABO": "XLPE", "POSTE": "DT"}
    assert s["valores"] == {"CABO": ["XLPE", "ALUM", "COBRE"], "POSTE": ["DT"]} and s["mascaras"] == {"POSTE": ["D%", "C%"]}
    assert s["fechados"] == ["CABO", "POSTE"]
    sig = {"N1": s}
    f = lambda t: ap.resolver_pedido(t, sig)
    assert f("N1(ALUM,D12)")["texto"] == "N1(ALUM,D12)" and not f("N1(ALUM,D12)")["erros"]       # D12 casa D%
    assert f("N1(COBRE,CC)")["erros"] == [] and f("N1(ALUM,XX)")["erros"] and "D%" in f("N1(ALUM,XX)")["erros"][0]
    so_exceto = ap.montar_assinaturas(["N1(cabo!=XLPE)"])
    assert so_exceto["N1"]["fechados"] == [] and ap.resolver_pedido("N1(QUALQUER)", so_exceto)["erros"] == []      # só "!=": não limita o valor


def test_compatibilidade_de_restricoes():
    c = lambda a, b: ap._compativeis(a, b)
    eq = lambda *alts: {"op": "=", "alts": list(alts)}
    ne = lambda *alts: {"op": "!=", "alts": list(alts)}
    assert c(eq("DT"), eq("D%")) and not c(eq("DT"), eq("C%")) and c(eq("D%"), eq("%T")) and not c(eq("A"), eq("B"))
    assert c(eq("A", "B"), eq("B")) and c(ne("X"), ne("Y")) and c(eq("A"), ne("B")) and not c(eq("A"), ne("A")) and not c(eq("A", "B"), ne("A", "B"))
    assert c(eq("D%"), ne("DT")) and c(eq("A", "B"), ne("A"))
    assert not c(eq("D%"), eq("C%")) and not c(eq("%T"), eq("%C")) and c(eq("D%"), eq("%T")) and c(eq("D%"), eq("DT%")) and not c(eq("DT%"), eq("DC%"))


# ── componente automático ──
def Lc(ativo="", comp="", cod="", proj="PERNAMBUCO", lv="A"):
    return {"ativo": ativo, "componente": comp, "codigo": cod, "projeto": proj, "lv": lv}


def test_preencher_componentes_e_heranca():
    linhas = [Lc("N1=N1(cabo=XLPE, poste=DT)", "", "c1"), Lc("", "", "c2"), Lc("", "", "c3"),
              Lc("N1(cabo=XLPE)", "", "x1"), Lc("", "", "x2"),
              Lc("N1(cabo=ALUM|COBRE)", "", "n1"), Lc("N1(poste=D%, cabo!=XLPE)", "", "pd1"), Lc("", "", "pd2"),
              Lc("N1(cabo=ALUM)", "N1.MANUAL", "a1"), Lc("", "", "a2"),
              Lc("P50", "", "p1"), Lc("", "", "p2")]
    novas, n = ap.preencher_componentes(linhas)
    assert [l["componente"] for l in novas] == ["N1", "N1", "N1", "N1.CABO=XLPE", "N1.CABO=XLPE", "N1.CABO=ALUM|COBRE", "N1.POSTE=D%,CABO!=XLPE",
                                                "N1.POSTE=D%,CABO!=XLPE", "N1.MANUAL", "N1.MANUAL", "", ""]
    assert n == 9 and linhas[0]["componente"] == ""                                              # não altera a lista original
    novas2, n2 = ap.preencher_componentes(novas)                                                # idempotente
    assert n2 == 0 and novas2 == novas


def test_preencher_componentes_nao_mexe_em_ativo_sem_parametros_nem_cruza_projetos():
    sem = [Lc("P50", "", "p1"), Lc("", "", "p2"), Lc("DT9/150", "", "d1")]
    assert ap.preencher_componentes(sem) == (sem, 0)
    linhas = [Lc("N1(cabo=XLPE)", "", "x1", proj="RONDONIA"), Lc("N1", "", "c1", proj="PERNAMBUCO"), Lc("", "", "x2", proj="RONDONIA")]
    novas, n = ap.preencher_componentes(linhas)
    assert [l["componente"] for l in novas] == ["N1.CABO=XLPE", "", "N1.CABO=XLPE"] and n == 2         # N1 comum de outro projeto: base sem parâmetros lá
    so_decl = [Lc("N1=N1(cabo=XLPE)", "", ""), Lc("N1(cabo=XLPE)", "", "x1"), Lc("", "", "x2")]
    novas, n = ap.preencher_componentes(so_decl)
    assert [l["componente"] for l in novas] == ["", "N1.CABO=XLPE", "N1.CABO=XLPE"]                 # linha só de declaração não ganha componente
    assert ap.validar_base(novas) == {"erros": [], "avisos": []}


def test_validar_base_com_operadores():
    linhas = [Lc("N1=N1(cabo=XLPE, poste=DT)", "N1", "c1"), Lc("N1(cabo=XLPE)", "N1.X", "x1"), Lc("N1(poste=D%)", "N1.D", "d1"), Lc("N1(cabo!=XLPE)", "N1.NX", "n1")]
    assert ap.validar_base(linhas) == {"erros": [], "avisos": []}
    sem_padrao_valido = [Lc("N1=N1(cabo=ZZ%)", "N1", "c1")]
    assert ap.validar_base(sem_padrao_valido)["erros"]                                           # padrão com curinga: erro de sintaxe
    padrao_sem_grupo = [Lc("N1=N1(cabo=ALUM)", "N1", "c1"), Lc("N1(cabo=XLPE)", "N1.X", "x1")]
    assert any("padrão CABO=ALUM" in e for e in ap.validar_base(padrao_sem_grupo)["erros"])
    padrao_por_curinga = [Lc("N1=N1(poste=DT)", "N1", "c1"), Lc("N1(poste=D%)", "N1.D", "d1")]
    assert ap.validar_base(padrao_por_curinga)["erros"] == []                                    # o padrão DT é aceito pelo grupo D%
    aviso = [Lc("N1(poste=D%)", "N1.D", "k1"), Lc("N1(poste=%T)", "N1.T", "k1")]
    assert len(ap.validar_base(aviso)["avisos"]) == 1                                            # D% e %T podem valer juntos
    nenhum = [Lc("N1(poste=D%)", "N1.D", "k1"), Lc("N1(poste=C%)", "N1.C", "k1"), Lc("N1(cabo=A)", "N1.A", "k2"), Lc("N1(cabo!=A)", "N1.NA", "k2")]
    assert ap.validar_base(nenhum)["avisos"] == []


@pytestmark_node
def test_paridade_operadores_e_pedidos():
    for celulas in (BASE80, CELULAS80):
        assert node("A.montarAssinaturas(e)", celulas) == ap.montar_assinaturas(celulas)
    sig = ap.montar_assinaturas(BASE80)
    assert node("e.t.map(x=>A.resolverPedido(x, e.s))", {"t": TEXTOS80, "s": sig}) == [ap.resolver_pedido(t, sig) for t in TEXTOS80]
    assert node("e.map(t=>[A.analisarAtivo(t), A.analisarCelulaBase(t), A.analisarAtivo(t, true)])", TEXTOS80 + CELULAS80) == \
        [[ap.analisar_ativo(t), ap.analisar_celula_base(t), ap.analisar_ativo(t, True)] for t in TEXTOS80 + CELULAS80]
    pares = [["D%", "DT"], ["D%", "D"], ["%T", "DT"], ["%CC%", "ACCD"], ["a.b", "AXB"], ["A%B", "A_B"], ["D%", ""], ["X", "x"], ["A+%", "A+Z"]]
    assert node("e.map(p=>A.casarPadrao(p[0], p[1]))", pares) == [ap.casar_padrao(a, b) for a, b in pares]
    restr = [[{"op": "!=", "alts": ["ALUM", "COBRE"]}, "XLPE"], [{"op": "=", "alts": ["A", "B%"]}, "BX"], ["DT", "dt"]]
    assert node("e.map(p=>A.restricaoVale(p[0], p[1]))", restr) == [ap.restricao_vale(a, b) for a, b in restr]
    gr = ap.restricoes_do_grupo(ap.analisar_celula_base("N1(poste=D%, cabo!=XLPE)"))
    assert node("A.restricoesDoGrupo(A.analisarCelulaBase('N1(poste=D%, cabo!=XLPE)'))", None) == gr
    ped = [{"CABO": "ALUM", "POSTE": "D12"}, {"CABO": "XLPE", "POSTE": "D12"}, {"CABO": "ALUM"}]
    assert node("e.p.map(x=>A.linhaVale(e.g, x))", {"p": ped, "g": gr}) == [ap.linha_vale(gr, x) for x in ped]
