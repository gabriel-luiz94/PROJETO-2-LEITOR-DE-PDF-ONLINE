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
    assert a["params"] == [{"nome": "CABO", "valor": "XLPE"}, {"nome": "POSTE", "valor": "DT"}]
    assert ap.analisar_celula_base("N1=N1(cabo, poste)")["params"] == [{"nome": "CABO", "valor": None}, {"nome": "POSTE", "valor": None}]
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


def test_linha_vale():
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
