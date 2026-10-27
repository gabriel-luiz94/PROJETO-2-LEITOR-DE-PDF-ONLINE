"""TASK-096 — estrutura com parâmetros repetidos (N4 = dois `cabo`) preenchidos pelos cabos vinculados. Python = JS."""
import json
import os
import shutil
import subprocess

import pytest

import config
from services import ativo_parametros as ap
from services import anotacao_outros as ao
from services import vinculacao_regras as vr

NODE = shutil.which("node")
CFG = {"cabos": [{"padrao": "^A\\d+P$", "valor_cabo": "P"}, {"padrao": "^C\\d+$", "valor_cabo": "C"}, {"padrao": "^S\\d+/\\d+$", "valor_cabo": "A"}],
       "estruturas": [{"padrao": "^N4$", "qtd_cabos": 2, "parametros": "cabo=ativo, cabo=ativo, tipocabo=valor_cabo, tipocabo=valor_cabo"},
                      {"padrao": "^N1$", "qtd_cabos": 2, "parametro": "cabo"}]}


def primeiro(t):
    return (t or "").split(" ")[0].upper()


def codigos(a):
    return [p.strip().upper() for p in (a or "").split(",")[1:] if p.strip()]


def codigo(a, row=None):
    c = codigos(a)
    return c[0] if c else (a or "").split(",")[0].strip().upper()


def cab(a, v=None, f=None):
    d = {"entidade": "CABO", "operacao": "I", "ativo": a}
    if v is not None:
        d["vinculoEstruturas"] = v
    if f:
        d["estruturasFicticias"] = f
    return d


def est(a, fict=None):
    d = {"entidade": "ESTRUTURA", "operacao": "I", "ativo": a}
    if fict:
        d["cabosFicticios"] = fict
    return d


# ── ativo_parametros: nome repetido ──
SIG = ap.montar_assinaturas(["N4=N4(cabo, cabo)", "N4(cabo=%A35P%, cabo=%C16%)", "N4(cabo=%A35P%, cabo=%A35P%)"])


def test_nome_repetido_na_declaracao_nos_grupos_e_no_pedido():
    assert SIG["N4"]["nomes"] == ["CABO", "CABO#2"]
    r = ap.resolver_pedido("N4(cabo=C16,cabo=A35P)", SIG)
    assert r["erros"] == [] and r["texto"] == "N4(C16,A35P)"                           # a ordem não importa para ser cadastrado
    r = ap.resolver_pedido("N4(cabo=S1/0,cabo=C16)", SIG)
    assert r["erros"] and "S1/0" in r["erros"][0]                                      # "/" é aceito no valor; só não está cadastrado
    r = ap.resolver_pedido("N4(cabo=C16)", SIG)
    assert r["faltando"] == ["CABO#2"] and r["texto"] == "N4(CABO=C16)"                # só um cabo: pendente; na tela o nome volta a ser CABO


def test_grupo_com_nome_repetido_casa_sem_depender_da_ordem():
    g = ap.restricoes_do_grupo(ap.analisar_celula_base("N4(cabo=%A35P%, cabo=%C16%)"))
    assert ap.linha_vale(g, {"CABO": "C16", "CABO#2": "A35P"}) and ap.linha_vale(g, {"CABO": "A35P", "CABO#2": "C16"})
    assert not ap.linha_vale(g, {"CABO": "C16", "CABO#2": "C16"})                       # precisa de um de cada
    assert not ap.linha_vale(g, {"CABO": "A35P"})


# ── vínculo: um valor por cabo ──
def test_n4_com_dois_cabos_diferentes_recebe_os_dois():
    a = vr.analisar([cab("A35P", [0]), cab("C16", [0])], [est("N4")], CFG, primeiro, codigo, codigos)
    e = a["estruturas"][0]
    assert e["parametros"] == {"CABO": ["A35P", "C16"], "TIPOCABO": ["P", "C"]} and e["estado"] == "gr"


def test_cabos_iguais_repetem_o_valor_e_um_cabo_so_preenche_o_primeiro():
    a = vr.analisar([cab("A35P", [0]), cab("A35P", [0])], [est("N4")], CFG, primeiro, codigo, codigos)
    assert a["estruturas"][0]["parametros"] == {"CABO": ["A35P", "A35P"], "TIPOCABO": ["P", "P"]}
    a = vr.analisar([cab("S1/0", [0])], [est("N4")], CFG, primeiro, codigo, codigos)
    e = a["estruturas"][0]
    assert e["parametros"] == {"CABO": ["S1/0"], "TIPOCABO": ["A"]} and e["estado"] == "pu"          # 1 de 2: pendente


def test_cabo_ficticio_tambem_preenche_um_parametro():
    a = vr.analisar([cab("A35P", [0])], [est("N4", ["C16"])], CFG, primeiro, codigo, codigos)
    assert a["estruturas"][0]["parametros"]["CABO"] == ["A35P", "C16"]


def test_parametros_por_estrutura_nao_mistura_n1_e_n4_da_mesma_linha():
    outros = [est("9400D, N1, N4")]
    cabos = [cab("C16", [0.01]), cab("A35P", [0.01]), cab("A35P", [0])]
    por = vr.parametros_por_estrutura(cabos, outros, CFG, primeiro, codigo, codigos)
    assert por == {0: {"N4": {"CABO": ["C16", "A35P"], "TIPOCABO": ["C", "P"]}, "N1": {"CABO": "P"}}}


def test_interpretar_linha_aplica_a_cada_item_so_o_dele():
    sig = ap.montar_assinaturas(["N4=N4(cabo, cabo)", "N1=N1(cabo)", "N1(cabo=%P%)", "N4(cabo=%A35P%, cabo=%C16%)"])
    ctx = {"assinaturas": sig, "parametros_estrutura": {"N4": {"CABO": ["C16", "A35P"]}, "N1": {"CABO": "P"}}}
    r = ao.interpretar_linha("9400D, N1, N4", "I", None, ctx)
    nomes = [i["ativo"] for i in r["itens"]]
    assert "N1(CABO=P)" in nomes and "N4(CABO=C16,CABO=A35P)" in nomes


def test_validar_config_parametros():
    assert vr.validar_config({"estruturas": [{"padrao": "N4", "parametros": "cabo=ativo, cabo=ativo"}]}) == []
    assert vr.validar_config({"estruturas": [{"padrao": "N4", "parametros": "cabo=xyz"}]})


@pytest.mark.skipif(not NODE, reason="node indisponível")
def test_paridade_js():
    js = """
    const VR = require(process.argv[1]), AP = require(process.argv[2]), AO = require(process.argv[3]); const p = t => (t || '').split(' ')[0].toUpperCase();
    const cods = a => (a || '').split(',').map(x => x.trim().toUpperCase()).slice(1).filter(Boolean);
    const cod = (a, row) => { const c = cods(a); return c.length ? c[0] : (a || '').split(',')[0].trim().toUpperCase(); };
    const cfg = %s;
    const cab = (a, v, f) => { const d = { entidade: 'CABO', operacao: 'I', ativo: a }; if (v) d.vinculoEstruturas = v; if (f) d.estruturasFicticias = f; return d; };
    const est = (a, f) => Object.assign({ entidade: 'ESTRUTURA', operacao: 'I', ativo: a }, f ? { cabosFicticios: f } : {});
    const sig = AP.montarAssinaturas(['N4=N4(cabo, cabo)', 'N4(cabo=%%A35P%%, cabo=%%C16%%)', 'N4(cabo=%%A35P%%, cabo=%%A35P%%)']);
    const g = AP.restricoesDoGrupo(AP.analisarCelulaBase('N4(cabo=%%A35P%%, cabo=%%C16%%)'));
    const outros = [est('9400D, N1, N4')], cabos = [cab('C16', [0.01]), cab('A35P', [0.01]), cab('A35P', [0])];
    const sig2 = AP.montarAssinaturas(['N4=N4(cabo, cabo)', 'N1=N1(cabo)', 'N1(cabo=%%P%%)', 'N4(cabo=%%A35P%%, cabo=%%C16%%)']);
    console.log(JSON.stringify({ sig: sig.N4.nomes, r1: AP.resolverPedido('N4(cabo=C16,cabo=A35P)', sig), r2: AP.resolverPedido('N4(cabo=S1/0,cabo=C16)', sig), r3: AP.resolverPedido('N4(cabo=C16)', sig),
        lv: [AP.linhaVale(g, { CABO: 'C16', 'CABO#2': 'A35P' }), AP.linhaVale(g, { CABO: 'C16', 'CABO#2': 'C16' }), AP.linhaVale(g, { CABO: 'A35P' })],
        a: VR.analisar([cab('A35P', [0]), cab('C16', [0])], [est('N4')], cfg, p, cod, cods).estruturas, a2: VR.analisar([cab('S1/0', [0])], [est('N4')], cfg, p, cod, cods).estruturas,
        a3: VR.analisar([cab('A35P', [0])], [est('N4', ['C16'])], cfg, p, cod, cods).estruturas, por: VR.parametrosPorEstrutura(cabos, outros, cfg, p, cod, cods),
        it: AO.interpretarLinha('9400D, N1, N4', 'I', null, { assinaturas: sig2, parametros_estrutura: { N4: { CABO: ['C16', 'A35P'] }, N1: { CABO: 'P' } } }).itens.map(i => i.ativo),
        v: [VR.validarConfig({ estruturas: [{ padrao: 'N4', parametros: 'cabo=ativo, cabo=ativo' }] }), VR.validarConfig({ estruturas: [{ padrao: 'N4', parametros: 'cabo=xyz' }] }).length > 0] }));
    """ % json.dumps(CFG).replace("%", "%%")
    s = json.loads(subprocess.run([NODE, "-e", js.replace("%%%%", "%%") if False else js, os.path.join(config.STATIC_DIR, "vinculacao_regras.js"), os.path.join(config.STATIC_DIR, "ativo_parametros.js"),
                                   os.path.join(config.STATIC_DIR, "anotacao_outros.js")], capture_output=True, text=True, check=True).stdout)
    assert s["sig"] == SIG["N4"]["nomes"]
    for chave, pedido in (("r1", "N4(cabo=C16,cabo=A35P)"), ("r2", "N4(cabo=S1/0,cabo=C16)"), ("r3", "N4(cabo=C16)")):
        assert s[chave] == ap.resolver_pedido(pedido, SIG)
    g = ap.restricoes_do_grupo(ap.analisar_celula_base("N4(cabo=%A35P%, cabo=%C16%)"))
    assert s["lv"] == [ap.linha_vale(g, {"CABO": "C16", "CABO#2": "A35P"}), ap.linha_vale(g, {"CABO": "C16", "CABO#2": "C16"}), ap.linha_vale(g, {"CABO": "A35P"})]
    assert {int(k): v for k, v in s["a"].items()} == vr.analisar([cab("A35P", [0]), cab("C16", [0])], [est("N4")], CFG, primeiro, codigo, codigos)["estruturas"]
    assert {int(k): v for k, v in s["a2"].items()} == vr.analisar([cab("S1/0", [0])], [est("N4")], CFG, primeiro, codigo, codigos)["estruturas"]
    assert {int(k): v for k, v in s["a3"].items()} == vr.analisar([cab("A35P", [0])], [est("N4", ["C16"])], CFG, primeiro, codigo, codigos)["estruturas"]
    cabos = [cab("C16", [0.01]), cab("A35P", [0.01]), cab("A35P", [0])]
    assert {int(k): v for k, v in s["por"].items()} == vr.parametros_por_estrutura(cabos, [est("9400D, N1, N4")], CFG, primeiro, codigo, codigos)
    sig2 = ap.montar_assinaturas(["N4=N4(cabo, cabo)", "N1=N1(cabo)", "N1(cabo=%P%)", "N4(cabo=%A35P%, cabo=%C16%)"])
    r = ao.interpretar_linha("9400D, N1, N4", "I", None, {"assinaturas": sig2, "parametros_estrutura": {"N4": {"CABO": ["C16", "A35P"]}, "N1": {"CABO": "P"}}})
    assert s["it"] == [i["ativo"] for i in r["itens"]]
    assert s["v"] == [[], True]
