"""TASK-101 — `cabo` = descrição completa do cabo (sem o comprimento); `tipocabo` = coluna VALOR DE CABO da Tabela de Cabos. Python = JS."""
import json
import os
import shutil
import subprocess

import pytest

import config
from services import ativo_parametros as ap
from services import vinculacao_regras as vr
from services.autonomo.montagem import prefixo_cabo
from services.autonomo.vinculacao import codigo_estrutura, codigos_estrutura

NODE = shutil.which("node")
SEED = json.load(open(os.path.join(config.BASE_DIR if hasattr(config, "BASE_DIR") else ".", "data", "vinculacao_pernambuco_seed.json"), encoding="utf-8"))
TEXTOS = ["3xS04-7,97kV 45 m", "3XS1/0-13.8KV 35 M", "3xA35P(7,94)-13,8kV 45 m", "3xC16(16) 30M", "C 3x1x95(95)+20kV 60 m", "RAMA25 12", "  3x s04 - 7,97 kv   45,5 m "]


def cab(t, op="I", v=None):
    d = {"entidade": "CABO", "operacao": op, "ativo": t, "textoOriginal": t}
    if v is not None:
        d["vinculoEstruturas"] = v
    return d


def est(a):
    return {"entidade": "ESTRUTURA", "operacao": "I", "ativo": a}


def analisa(cabos, outros, cfg=SEED):
    return vr.analisar(cabos, outros, cfg, prefixo_cabo, codigo_estrutura, codigos_estrutura)


def test_descricao_completa_sem_o_comprimento():
    assert vr.descricao_cabo("3xS04-7,97kV 45 m") == "3XS04-7.97KV"
    assert vr.descricao_cabo("3XS1/0-13.8KV 35 M") == "3XS1/0-13.8KV"
    assert vr.descricao_cabo("3xA35P(7,94)-13,8kV 45 m") == "3XA35P_7.94_-13.8KV"          # parênteses viram "_" (não cabem dentro de um parâmetro)
    assert vr.descricao_cabo("3xC16(16) 30M") == "3XC16_16_"
    assert vr.descricao_cabo("RAMA25 12") == "RAMA25"
    assert vr.descricao_cabo("", "c16") == "C16"


def test_cabo_e_tipocabo_sao_coisas_diferentes_na_estrutura():
    cfg = {"cabos": [{"padrao": "^S\\d+$", "valor_cabo": "C"}, {"padrao": "^A\\d+P$", "valor_cabo": "P"}],           # S04 cadastrado como COBRE
           "estruturas": [{"padrao": "^N4$", "qtd_cabos": 2, "parametros": "cabo, cabo, tipocabo, tipocabo"}, {"padrao": "^N1$", "qtd_cabos": 1, "parametro": "tipocabo"}]}
    a = analisa([cab("3xS04-7,97kV 45 m", "I", [0]), cab("3xA35P(7,94)-13,8kV 45 m", "I", [0])], [est("N4")], cfg)
    assert a["estruturas"][0]["parametros"] == {"CABO": ["3XS04-7.97KV", "3XA35P_7.94_-13.8KV"], "TIPOCABO": ["C", "P"]}
    a = analisa([cab("3xS04-7,97kV 45 m", "I", [0])], [est("N1")], cfg)
    assert a["estruturas"][0]["parametros"] is None and a["estruturas"][0]["valor_cabo"] == "C" and a["estruturas"][0]["parametro"] == "TIPOCABO"
    por = vr.parametros_por_estrutura([cab("3xS04-7,97kV 45 m", "I", [0])], [est("N1")], cfg, prefixo_cabo, codigo_estrutura, codigos_estrutura)
    assert por == {0: {"N1": {"TIPOCABO": "C"}}}


def test_fonte_explicita_vale_mais_que_o_nome():
    cfg = {"cabos": [{"padrao": "^S\\d+$", "valor_cabo": "C"}], "estruturas": [{"padrao": "^N1$", "qtd_cabos": 1, "parametros": "cabo=ativo, tipocabo=descricao"}]}
    a = analisa([cab("3xS04-7,97kV 45 m", "I", [0])], [est("N1")], cfg)
    assert a["estruturas"][0]["parametros"] == {"CABO": ["S04"], "TIPOCABO": ["3XS04-7.97KV"]}


def test_pedido_e_base_aceitam_a_descricao_do_cabo():
    sig = ap.montar_assinaturas(["N1=N1(tipocabo, cabo)", "N1(tipocabo=C, cabo=%S04%)", "N1(tipocabo=A, cabo=%7,97KV%)"])
    r = ap.resolver_pedido("N1(tipocabo=C,cabo=3XS04-7.97KV)", sig)
    assert r["erros"] == [] and r["texto"] == "N1(C,3XS04-7.97KV)"
    assert ap.analisar_celula_base("N1(tipocabo=A, cabo=%7,97KV%)")["params"][1]["alts"] == ["%7.97KV%"]        # vírgula decimal da tensão não separa parâmetros


def test_seed_pernambuco_usa_tipocabo_e_cabo():
    assert vr.validar_config(SEED) == []
    n1 = next(e for e in SEED["estruturas"] if e["padrao"] == "^N1$")
    n4 = next(e for e in SEED["estruturas"] if "N4" in e["padrao"])
    assert n1["parametro"] == "tipocabo" and n4["parametros"] == "cabo, cabo"


@pytest.mark.skipif(not NODE, reason="node indisponível")
def test_paridade_js():
    js = """
    const VR = require(process.argv[1]), AP = require(process.argv[2]);
    const textos = %s, cfg = %s;
    const sig = AP.montarAssinaturas(['N1=N1(tipocabo, cabo)', 'N1(tipocabo=C, cabo=%%S04%%)', 'N1(tipocabo=A, cabo=%%7,97KV%%)']);
    console.log(JSON.stringify({ d: textos.map(t => VR.descricaoCabo(t, '')), x: VR.extraCabo({ textoOriginal: textos[0] }).descricao,
        r: AP.resolverPedido('N1(tipocabo=C,cabo=3XS04-7.97KV)', sig), g: AP.analisarCelulaBase('N1(tipocabo=A, cabo=%%7,97KV%%)').params,
        n: VR.normalizarParametros('cabo, cabo, tipocabo=ativo, x=descricao') }));
    """ % (json.dumps(TEXTOS), json.dumps(SEED))
    s = json.loads(subprocess.run([NODE, "-e", js, os.path.join(config.STATIC_DIR, "vinculacao_regras.js"), os.path.join(config.STATIC_DIR, "ativo_parametros.js")], capture_output=True, text=True, check=True).stdout)
    assert s["d"] == [vr.descricao_cabo(t) for t in TEXTOS]
    sig = ap.montar_assinaturas(["N1=N1(tipocabo, cabo)", "N1(tipocabo=C, cabo=%S04%)", "N1(tipocabo=A, cabo=%7,97KV%)"])
    assert s["r"] == ap.resolver_pedido("N1(tipocabo=C,cabo=3XS04-7.97KV)", sig)
    assert s["g"] == ap.analisar_celula_base("N1(tipocabo=A, cabo=%7,97KV%)")["params"]
    assert s["n"] == vr.normalizar_parametros("cabo, cabo, tipocabo=ativo, x=descricao")
