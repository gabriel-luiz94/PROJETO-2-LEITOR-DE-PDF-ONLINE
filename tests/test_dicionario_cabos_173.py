"""TASK-173 — dicionário de cabos por projeto (notação, nível, seção, ordem do "maior"). Paridade Python = JS."""
import json
import os
import shutil
import subprocess

import pytest

import config
from services import anotacao_cabos as ac
from services import dicionario_cabos as dc
from services import vinculacao_regras as vr

NODE = shutil.which("node")
DIC081 = json.load(open(os.path.join(config.BASE_DIR, "data", "pacotes", "081", "dicionario_cabos.json"), encoding="utf-8"))
OUTRO = [{"padrao": r"^ALU-(\d+)$", "familia": "ALUMINIO", "valor_cabo": "A", "secao": "$1", "nivel": "MT", "fases_padrao": 3, "neutro": "NEU{valor}", "notacao_solta": True},
         {"padrao": r"^CU-(\d+)$", "familia": "COBRE", "valor_cabo": "C", "secao": "$1", "nivel": "BT", "notacao_solta": True, "apelidos": [{"padrao": r"^(\d{2})$", "codigo": "CU-$1"}]}]
TEXTOS = ["3xA35P(7,94)-13,8kV 45 m", "3xC16(16) 30M", "3xS50P(35) 40", "1xS1/0-7,9kV 20", "RAMA25 12", "PB35 5", "16 10 m", "(S04) 8", "C 3x1x95(95)+20kV 12",
          "3xS04-13,8kV 100", "CAA2 ABC 35 m", "2xS336 12", "xyz 5", "", "3xCAZ7,94 5", "S02 4", "1xAA4/0 9"]
CODIGOS = ["C16", "A185P", "S35P", "S04", "S02", "S1/0", "S2/0", "AA4/0", "CAA2", "S336", "RAMA25", "PB35", "CAZ7,94", "ALU-50", "CU-16", "", "ZZ"]


def _node(expr, **dados):
    js = ("const DC=require('./static/dicionario_cabos.js'),AC=require('./static/anotacao_cabos.js'),VR=require('./static/vinculacao_regras.js');"
          f"const D={json.dumps(dados)};console.log(JSON.stringify({expr}));")
    r = subprocess.run([NODE, "-e", js], cwd=config.BASE_DIR, capture_output=True, text=True, timeout=60)
    assert r.returncode == 0, r.stderr
    return json.loads(r.stdout)


def itens(texto, dic=None):
    r = ac.interpretar(texto, dic)
    return [(i["ativo"], i["fases"], i["comprimento"], i["neutro"]) for i in r["itens"]], r["tensao"], r["erros"]


def test_dicionario_081_da_o_mesmo_resultado_de_sem_dicionario():
    for t in TEXTOS:
        assert itens(t, DIC081) == itens(t), t
        assert ac.eh_notacao(t, DIC081) == ac.eh_notacao(t), t
    for c in CODIGOS:
        assert vr.secao_cabo(c, DIC081) == vr.secao_cabo(c), c


def test_dicionario_081_nao_muda_nivel_nem_valor_cabo():
    for t in TEXTOS:
        assert vr.derivar_cabo(t, DIC081) == vr.derivar_cabo(t), t
    cfg = vr.normalizar_config({"dicionario_cabos": DIC081})
    assert cfg["ordem_maior"] == ["A", "C", "P"] or cfg["ordem_maior"][0] in ("A", "C")      # sem ordem própria usa a do dicionário (valores distintos)
    assert vr.normalizar_config({"dicionario_cabos": DIC081, "ordem_maior": ["C", "P", "A"]})["ordem_maior"] == ["C", "P", "A"]


def test_outra_convencao_so_pelo_dicionario():
    assert itens("ALU-50 30 m", OUTRO)[0] == [("ALU-50", 3, 30.0, False)]                            # fases_padrao
    assert itens("3xALU-50(NEU7,9)-13,8kV 45 m", OUTRO)[2] == []
    assert itens("16 10 m", OUTRO)[0] == [("CU-16", 1, 10.0, False)]                                  # apelido com código
    assert itens("ALU-50 30 m")[0] == []                                                              # sem dicionário não é notação
    assert itens("2xALU-50(7,94) 10", OUTRO)[0] == [("ALU-50", 2, 10.0, False), ("NEU7,94", 1, 10.0, True)]   # neutro pelo modelo
    assert vr.secao_cabo("ALU-50", OUTRO) == 50.0 and vr.secao_cabo("CU-16", OUTRO) == 16.0
    info = vr.info_cabo("ALU-50", "I", vr.normalizar_config({"dicionario_cabos": OUTRO}))
    assert info["valor_cabo"] == "A" and info["secao"] == 50.0


def test_nivel_sem_tensao_vem_do_dicionario_e_tensao_manda():
    assert vr.derivar_cabo("ALU-50 30 m", OUTRO)["nivel"] == "MT"
    assert vr.derivar_cabo("ALU-50-0,4kV 30 m", OUTRO)["nivel"] == "BT"                                # tensão no texto continua mandando
    assert vr.derivar_cabo("CU-16 10", OUTRO)["nivel"] == "BT"
    assert vr.derivar_cabo("ALU-50 30 m")["nivel"] == "BT"                                            # sem dicionário: como hoje
    assert vr.extra_cabo({"textoOriginal": "ALU-50 30 m"}, OUTRO)["nivel"] == "MT"


def test_maior_primeiro_pela_secao_do_dicionario():
    cfg = vr.normalizar_config({"dicionario_cabos": [{"padrao": "^X(\\d)$", "secao": 100}, {"padrao": "^Y$", "secao": 5}]})
    assert vr.info_cabo("X1", "I", cfg)["secao"] == 100.0 and vr.info_cabo("Y", "I", cfg)["secao"] == 5.0
    assert vr.info_cabo("C16", "I", cfg)["secao"] == 16.0                                              # sem linha: regra fixa


def test_validacao_do_dicionario():
    assert dc.validar([]) == [] and dc.validar(None) == []
    erros = dc.validar([{"padrao": "("}, {"padrao": "^A$", "nivel": "XX", "fases_padrao": 0, "secao": "abc", "neutro": "CAZ"}, {"valor_cabo": "A"}, 5])
    assert len(erros) >= 6 and any("regex" in e for e in erros) and any("neutro" in e for e in erros)
    assert vr.validar_config({"dicionario_cabos": [{"padrao": "("}]})


def test_testar_um_texto_mostra_a_linha_usada():
    r = dc.testar(OUTRO, "3xALU-50(7,94)-13,8kV 45 m")
    assert r["codigo"] == "ALU-50" and r["fases"] == 3 and r["nivel"] == "MT" and r["secao"] == 50.0 and r["valor_cabo"] == "A" and r["linha"] == 0
    assert r["neutro"] == "NEU7,94"
    assert dc.testar(OUTRO, "??")["codigo"] is None


@pytest.mark.skipif(not NODE, reason="sem node")
def test_paridade_com_o_javascript():
    for dic in (None, DIC081, OUTRO):
        for t in TEXTOS + ["ALU-50 30 m", "3xALU-50(7,94)-13,8kV 45 m", "16 10 m", "CU-16 10"]:
            py = ac.interpretar(t, dic)
            js = _node("AC.interpretar(D.t, D.d)", t=t, d=dic)
            assert py == js, (t, py, js)
            assert vr.derivar_cabo(t, dic) == _node("VR.derivarCabo(D.t, D.d)", t=t, d=dic), t
        for c in CODIGOS:
            assert vr.secao_cabo(c, dic) == _node("VR.infoCabo(D.c, 'I', VR.normalizarConfig({dicionario_cabos: D.d})).secao", c=c, d=dic), c
    assert dc.validar([{"padrao": "("}, {"padrao": "^A$", "nivel": "XX", "fases_padrao": 0, "secao": "abc", "neutro": "CAZ"}]) == _node("DC.validar(D.d)", d=[{"padrao": "("}, {"padrao": "^A$", "nivel": "XX", "fases_padrao": 0, "secao": "abc", "neutro": "CAZ"}])


def test_aba_e_rota_do_dicionario_ligadas():
    ui = open(os.path.join(config.STATIC_DIR, "vinculacao_ui.js"), encoding="utf-8").read()
    assert "abaDicionario" in ui and "'Dicionário de cabos'" in ui and "Testar um texto" in ui
    html = open(os.path.join(config.STATIC_DIR, "resumo.html"), encoding="utf-8").read()
    assert html.index("/static/dicionario_cabos.js") < html.index("/static/anotacao_cabos.js") < html.index("/static/vinculacao_regras.js")
