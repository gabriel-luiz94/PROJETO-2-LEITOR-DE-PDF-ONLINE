"""TASK-083 / TASK-082 — notação do LABEL na tabela Cabos e regras do Leitor de Pernambuco."""
import json
import os
import random
import shutil
import sqlite3
import subprocess

import pytest

import config
import database
from services import anotacao_cabos as ac
from services.autonomo import montagem as mt
from services.autonomo import totalizadora as tt
from services.validacao_planilhas import validar_planilhas
from tests import oraculo_tela as o

NODE = shutil.which("node")


def itens(texto):
    r = ac.interpretar(texto)
    return [(i["ativo"], i["fases"], i["comprimento"], i["neutro"]) for i in r["itens"]], r["tensao"], r["erros"]


# ── interpretação ──
def test_neutro_com_virgula_vira_caz():
    assert itens("3xA35P(7,94)-13,8kV 45 m") == ([("A35P", 3, 45.0, False), ("CAZ7,94", 1, 45.0, True)], "13,8", [])


def test_neutro_igual_herda_o_cabo_e_secao_diferente_troca_a_secao():
    assert itens("3xC16(16) 30M")[0] == [("C16", 3, 30.0, False), ("C16", 1, 30.0, True)]
    assert itens("3xS50P(35) 40")[0] == [("S50P", 3, 40.0, False), ("S35P", 1, 40.0, True)]
    assert itens("1xS95P(50) 40")[0][1] == ("S50P", 1, 40.0, True)


def test_fase_implicita_tensao_ignorada_e_comprimentos():
    assert itens("1xS1/0-7,9kV 20")[0] == [("S1/0", 1, 20.0, False)]
    assert itens("RAMA25 12")[0] == [("RAMA25", 1, 12.0, False)]
    assert itens("PB35 12,5 m")[0] == [("PB35", 1, 12.5, False)]
    assert itens("3xW2,25 15 m")[0] == [("W2,25", 3, 15.0, False)]
    for forma in ["45 m", "45m", "45M", "45"]:
        assert itens(f"3xC16 {forma}")[0][0][2] == 45.0


def test_soltos_numero_e_potencia():
    assert itens("(S04) 8 m")[0] == [("S04", 1, 8.0, True)]
    assert itens("[C16] 10")[0] == [("C16", 1, 10.0, True)]
    assert itens("16 9 m")[0] == [("C16", 1, 9.0, False)]
    assert itens("C 3x1x95(95)+20kV 60 m") == ([("C3X1X95(95)", 1, 60.0, False)], "20", [])


def test_tensao_em_qualquer_caixa_e_ignorada():
    for t in ["3xC35-13,8kV 45 m", "3xC35-13,8KV 45 m", "3xC35-13,8kv 45m", "3xC35 - 13,8 kV 45"]:
        assert itens(t)[0] == [("C35", 3, 45.0, False)], t


def test_formato_antigo_nao_e_notacao():
    for t in ["CAA2 ABC 35 m", "CAA2 3 45", "M3X16 20 m", "CAA 2 ABC 35", "P50", "CU 25 AB 10 m"]:
        assert not ac.eh_notacao(t), t
        assert ac.interpretar(t)["formato"] == "antigo"


def test_erros():
    assert ac.interpretar("3xA35P(7,94)")["erros"]
    assert ac.interpretar("3x(  ) 45 m")["erros"]
    assert ac.interpretar("3xC16() 45 m")["erros"]


def test_fases_e_prefixo():
    assert ac.fases_da_linha("3xA35P(7,94) 45 m") == 3 and ac.fases_da_linha("CAA2 ABC 35") == 1
    assert ac.prefixo_e_comprimento("3xA35P(7,94) 45 m") == ("A35P", 45.0)
    assert mt.calcular_qtd_ativos("3xA35P(7,94)-13,8kV 45 m") == 3 and mt.prefixo_cabo("RAMA25 12") == "RAMA25"
    assert not mt.linha_standalone("PB35 12")


# ── Totalizadora (Python = JS real) ──
BASE_T = [{"ativo": a, "desc_ativo": a, "componente": a, "projeto": "P1", "codigo": f"c{i}", "desc_codigo": "d"} for i, a in
          enumerate(["A35P", "CAZ7,94", "C16", "S1/0", "RAMA25", "S50P", "S35P"])]
LINHAS = ["3xA35P(7,94)-13,8kV 45 m", "3xC16(16) 30M", "1xS1/0-7,9kV 20", "RAMA25 12", "(S04) 8 m", "16 9 m", "3xS50P(35) 40", "CAA2 ABC 35 m",
          "C 3x1x95(95)+20kV 60 m", "3xA35P(7,94)", "3xW2,25 15"]


@pytest.mark.skipif(not NODE, reason="node indisponível")
def test_totalizadora_com_notacao_igual_ao_js():
    cabos = [{"entidade": "CABO", "operacao": op, "ativo": t} for op in ("I", "R") for t in LINHAS]
    mt.recalcular_qtd_ativos(cabos)
    js = o.totalizadora_e_payload(cabos, [], [], BASE_T, "P1")
    tot = tt.montar_totalizadora(cabos, [], [], BASE_T, "P1")
    assert tt.normalizar_para_json(tot) == js["totalizadora"]
    assert tt.payload_calculo(tot) == js["payload"]
    pl = tt.payload_calculo(tt.montar_totalizadora(cabos[:1], [], [], BASE_T, "P1"))["cabos"]
    assert pl == [{"entidade": "CABO", "operacao": "I", "ativo": "A35P 1 45"}] * 3 + [{"entidade": "CABO", "operacao": "I", "ativo": "CAZ7,94 1 45"}]


# ── paridade do módulo (inclui texto estranho) ──
_JS = """
const fs = require('fs'); const A = require(%(ac)s);
const t = JSON.parse(fs.readFileSync(0, 'utf8'));
console.log(JSON.stringify(t.map(x => [A.ehNotacao(x), A.interpretar(x), A.fasesDaLinha(x), A.prefixoEComprimento(x)])));
"""


@pytest.mark.skipif(not NODE, reason="node indisponível")
def test_paridade_python_js_do_modulo():
    rnd = random.Random(83)
    pecas = ["3x", "1X", "2x", "C ", "", "A35P", "C16", "S1/0", "RAMA25", "PB35", "W2,25", "(7,94)", "(16)", "(S04)", "[C16]", "-13,8kV", "+20kV", "-7,9kV",
             " 45 m", " 45m", " 45M", " 12,5", " 8", "16", "x", "(", ")", " ﻿", " ", "  "]
    textos = LINHAS + ["".join(rnd.choice(pecas) for _ in range(rnd.randint(1, 5))) for _ in range(600)]
    r = subprocess.run([NODE, "-e", _JS % {"ac": json.dumps(os.path.join(config.STATIC_DIR, "anotacao_cabos.js"))}], input=json.dumps(textos),
                       capture_output=True, text=True, timeout=60)
    assert r.returncode == 0, r.stderr
    js = json.loads(r.stdout)
    for t, (nota, interp, fases, pc) in zip(textos, js):
        assert ac.eh_notacao(t) == nota, repr(t)
        assert ac.interpretar(t) == interp, repr(t)
        assert ac.fases_da_linha(t) == fases, repr(t)
        esp = ac.prefixo_e_comprimento(t)
        assert (list(esp) if esp else None) == pc, repr(t)


# ── validação ──
def test_validacao_camada1_de_cabos():
    def cods(t):
        return {a["regra_id"] for a in validar_planilhas([{"ativo": t, "operacao": "I"}], [])}
    assert cods("3xA35P(7,94)-13,8kV 45 m") == set() and cods("CAA2 ABC 35 m") == set()
    assert "C1-CABO-NOTACAO" in cods("3xA35P(7,94)")


# ── regras do Leitor de Pernambuco (motor real em Node) ──
_PE = """
const E = require(%(motor)s); const fs = require('fs');
const e = JSON.parse(fs.readFileSync(0, 'utf8'));
e.itens.forEach((it, i) => { it.index = i; it.allItems = e.itens; });
console.log(JSON.stringify(e.itens.map(it => E.processarEClassificar(it, e.proc, e.cls))));
"""


def _leitor_pe(itens_):
    d = os.path.join(config.BASE_DIR if hasattr(config, "BASE_DIR") else ".", "data")
    proc = json.load(open(os.path.join(d, "regras_leitor_processamento_pernambuco_seed.json"), encoding="utf-8"))
    cls = json.load(open(os.path.join(d, "regras_leitor_classificacao_pernambuco_seed.json"), encoding="utf-8"))
    r = subprocess.run([NODE, "-e", _PE % {"motor": json.dumps(os.path.join(config.STATIC_DIR, "regras_leitor_engine.js"))}],
                       input=json.dumps({"itens": [dict({"pagina": 1}, **i) for i in itens_], "proc": proc, "cls": cls}), capture_output=True, text=True, timeout=60)
    assert r.returncode == 0, r.stderr
    return json.loads(r.stdout)


def T(texto, cor="#000000", layer=""):
    return {"texto": texto, "cor": cor, "layer": layer}


@pytest.mark.skipif(not NODE, reason="node indisponível")
def test_leitor_pernambuco_postes_cabos_e_operacoes():
    itens_ = [T("11300D", layer="01_INSTALANDO"), T("7150FV", "#000001", "01_REMOVER"), T("10300R", layer="01_REAPLICANDO"), T("9300F"), T("12400D", "#000001"),
              T("3xA35P(7,94)-13,8kV", layer="01_INSTALANDO"), T("45 m", layer="01_INSTALANDO"),
              T("3xC16(16) 30M", layer="01_INSTALANDO"), T("1xS1/0-7,9kV", "#000001"), T("12,5m", "#000001"),
              T("C 3x1x95(95)+20kV"), T("60 m"), T("RAMA25"), T("20 m"), T("(S04)"), T("8 m"), T("16"), T("9 m"), T("3xW2,25"), T("15 m")]
    r = _leitor_pe(itens_)
    got = [(x["operacao"], x["entidade"], x["ativo"]) for x in r]
    assert got[:5] == [("I", "POSTE", "11300D"), ("R", "POSTE", "7150FV"), ("M", "POSTE", "10300R"), ("I", "POSTE", "9300F"), ("R", "POSTE", "12400D")]
    assert got[5] == ("I", "CABO", "3xA35P(7,94)-13,8kV 45 m") and got[6][0] == "0"
    assert got[7] == ("I", "CABO", "3xC16(16) 30 m")
    assert got[8] == ("R", "CABO", "1xS1/0-7,9kV 12,5 m") and got[9][0] == "0"
    assert got[10] == ("I", "CABO", "C3x1x95(95)+20kV 60 m")
    assert got[12] == ("I", "CABO", "RAMA25 20 m") and got[14] == ("I", "CABO", "(S04) 8 m") and got[16] == ("I", "CABO", "16 9 m") and got[18] == ("I", "CABO", "3xW2,25 15 m")
    # o que o Leitor entrega vira a lista certa na tabela Cabos
    assert [i[:2] for i in itens(got[5][2])[0]] == [("A35P", 3), ("CAZ7,94", 1)]


@pytest.mark.skipif(not NODE, reason="node indisponível")
def test_leitor_pernambuco_ignora_ruido_e_cores_desconhecidas():
    r = _leitor_pe([T("VALDECI", "#0f0f0f", "TXT_KML"), T("Rua Arca de Noe", "#c0c0c0", "TXT_RUAS"), T("TESTE", "#0f0f0f", "01_INSTALANDO"), T("POSTE 1", layer="01_INSTALANDO"),
                    T("qualquer", "#123456"), T("11300D", "#0f0f0f"), T("11300D", "#0f0f0f", "TXT_KML")])
    assert [x["operacao"] for x in r][:2] == ["0", "0"]
    assert [x["entidade"] for x in r][2:4] == ["0", "0"]          # sem ativo reconhecido: a tela descarta (entidade 0)
    assert r[4]["operacao"] == "0" and r[5]["operacao"] == "0" and r[6]["operacao"] == "0"


@pytest.mark.skipif(not NODE, reason="node indisponível")
def test_cor_exata_e_vizinhanca_para_frente():
    E = os.path.join(config.STATIC_DIR, "regras_leitor_engine.js")
    cod = f"""const E = require({json.dumps(E)});
    console.log(JSON.stringify([E.corHexCasa(['#000000'], null, '#000000'), E.corHexCasa(['#000000'], null, '#000001'), E.corHexCasa(null, ['#000001'], '#000001'), E.corHexCasa(null, null, '#abc')]));"""
    r = subprocess.run([NODE, "-e", cod], capture_output=True, text=True, timeout=30)
    assert json.loads(r.stdout) == [True, False, False, True]
    # "DEPOIS" não enxerga o comprimento do cabo anterior
    itens_ = [T("45 m", layer="01_INSTALANDO"), T("1xC16", layer="01_INSTALANDO"), T("30 m", layer="01_INSTALANDO")]
    assert [x["ativo"] for x in _leitor_pe(itens_)][1] == "1xC16 30 m"


def test_seed_pernambuco_081_e_regras_validas(tmp_path, monkeypatch):
    from routers.regras_leitor import validar_regras
    d = os.path.join(config.STATIC_DIR, "..", "data")
    proc = json.load(open(os.path.join(d, "regras_leitor_processamento_pernambuco_seed.json"), encoding="utf-8"))
    cls = json.load(open(os.path.join(d, "regras_leitor_classificacao_pernambuco_seed.json"), encoding="utf-8"))
    assert validar_regras("processamento", proc) == [] and validar_regras("classificacao", cls) == []
    monkeypatch.setattr(database, "DB_PATH", str(tmp_path / "t.db"))
    database.init_db()
    con = sqlite3.connect(tmp_path / "t.db")
    for t in ("regras_leitor_processamento", "regras_leitor_classificacao"):
        assert con.execute(f"SELECT COUNT(*) FROM {t} WHERE projeto_codigo = '081'").fetchone()[0] == 1
    con.execute("UPDATE regras_leitor_classificacao SET regras_json = '[]' WHERE projeto_codigo = '081'"); con.commit(); con.close()
    database.init_db()                      # o seed não sobrescreve o que o admin editou
    con = sqlite3.connect(tmp_path / "t.db")
    assert con.execute("SELECT regras_json FROM regras_leitor_classificacao WHERE projeto_codigo = '081'").fetchone()[0] == "[]"
