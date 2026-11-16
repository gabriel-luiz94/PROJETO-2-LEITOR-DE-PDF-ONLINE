"""TASK-174 — convenção de postes por projeto (tipos, formatos, compacto, faixas). Paridade Python = JS; sem tabela = Pernambuco."""
import json
import shutil
import subprocess

import pytest

import config
from services import anotacao_outros as ao
from services import convencao_postes as cp
from services import orcamento_calc as calc
from services import vinculacao_regras as vr

NODE = shutil.which("node")
CONV_P = {"tipos": [{"codigo": "P", "sufixo": "P", "nome": "poste"}], "formatos": ["{tipo}{altura}-{esforco}"], "altura": [7, 16], "esforco_digitos": [3, 4], "compacto": "{tipo}{altura}-{esforco}"}
CONV_B = {"tipos": [{"codigo": "B", "sufixo": "B"}, {"codigo": "CC", "sufixo": "C"}], "formatos": ["{tipo}{altura}/{esforco}", "{tipo}-{altura}-{esforco}", "{altura}{esforco}{sufixo}"],
          "altura": [8, 14], "esforco_digitos": [3, 3], "compacto": "{altura}{esforco}{sufixo}"}
NOMES = ["DT11/600", "R12/1000", "121000R", "11600D", "7150FV", "FV7/150", "F12/1000", "CV10/300", "RC9/200", "N3", "T3", "11600X", "123D", "xx", "", "DT11/600X", "101000F", "DT011/600"]
EXTRA = ["P11-400", "B11/600", "CC-11-300", "CC-11-300X", "11600B", "10300C", "P11/400", "p11-400"]


def L(ativo, comp="", cod="", fi=1, proj="XYZ"):
    return {"ativo": ativo, "desc_ativo": ativo and "D", "componente": comp, "projeto": proj, "mdo": "MATERIAL", "codigo": cod, "desc_codigo": f"d {cod}", "fator_i": fi, "fator_r": 0.0,
            "filtro": "", "origem": "OUTROS", "lv": "A"}


def test_padrao_de_pernambuco_e_identico_ao_sem_tabela():
    for n in NOMES:
        assert cp.compacto_do_poste(n, None) == cp.compacto_do_poste(n, cp.PADRAO) == ao.poste_para_compacto(n), n
        assert ao.eh_poste(n) == cp.eh_poste(n) == ao.eh_poste(n, {})
    assert ao.compacto_para_poste("11600D", {"DT1/1600", "DT11/600"})["ativo"] == "DT11/600"
    assert ao.compacto_para_poste("123D")["erro"] and ao.compacto_para_poste("N3") is None


def test_outro_formato_de_poste():
    assert cp.testar("P11-400", CONV_P) == {"poste": True, "tipo": "P", "ativo": "P11-400", "compacto": "P11-400", "erro": None}
    assert not cp.eh_poste("DT11/600", CONV_P) and not cp.eh_poste("N3", CONV_P) and not cp.eh_poste("P11/400", CONV_P)
    assert cp.testar("CC-11-300", CONV_B)["ativo"] == "CC11/300" and cp.testar("CC-11-300", CONV_B)["compacto"] == "11300C"
    assert cp.testar("B11/600", CONV_B)["compacto"] == "11600B" and cp.testar("11600B", CONV_B)["ativo"] == "B11/600"
    assert cp.testar("10300C", CONV_B)["ativo"] == "CC10/300"
    assert cp.testar("11600B", CONV_B, {"B116/00", "B1/1600"})["ativo"] == "B11/600"          # esforço de 3 dígitos fixo nesta convenção


def test_ambiguidade_resolvida_pela_base_como_hoje():
    conv = {"tipos": [{"codigo": "DT", "sufixo": "D"}], "formatos": ["{tipo}{altura}/{esforco}", "{altura}{esforco}{sufixo}"], "altura": [7, 16], "esforco_digitos": [3, 4], "compacto": "{altura}{esforco}{sufixo}"}
    assert cp.testar("11600D", conv, {"DT1/1600"})["ativo"] == "DT1/1600"
    assert cp.testar("123D", conv)["erro"]


def test_linha_de_outros_com_poste_do_projeto_passa_o_parametro_poste():
    base = [L("P11-400", "P11-400", "p400"), L("N3=N3(poste=P11-400)", "N3", "n3a"), L("", "N3", "n3b", fi=2), L("N3(poste=P11-400)", "N3.P", "n3esp"), L("T3", "T3", "t3")]
    ctx = ao.contexto_da_base(base, "XYZ", CONV_P)
    assert "P11-400" in ctx["postes"]
    r = ao.interpretar_linha("P11-400, 2 N3, T3", "I", calc._classico_outros, ctx)
    assert [(i["ativo"], i["qtd"]) for i in r["itens"]] == [("P11-400", 1.0), ("N3(POSTE=P11-400)", 2.0), ("T3", 1.0)]
    assert ao.codigo_estrutura("P11-400, N3", calc._classico_outros, ctx) == "N3"
    # sem a tabela do projeto, "P11-400" não é poste (vira estrutura)
    assert ao.codigo_estrutura("P11-400, N3", calc._classico_outros, ao.contexto_da_base(base, "XYZ")) == "P11-400"
    # M: poste ignorado
    r = ao.interpretar_linha("P11-400, N3", "M", calc._classico_outros, ctx)
    assert [i["ativo"] for i in r["itens"]] == ["N3(POSTE=P11-400)"]


def test_calculo_usa_a_convencao_do_projeto():
    base = [L("P11-400", "P11-400", "p400"), L("N3=N3(poste=P11-400)", "N3", "n3a"), L("", "N3", "n3b", fi=2), L("T3", "T3", "t3")]
    pedido = [{"ativo": "P11-400, N3, T3", "operacao": "I", "entidade": "0"}]
    com = calc.processar_calculo([], pedido, "XYZ", base, convencao_postes=CONV_P)
    assert {r["codigo"]: r["total"] for r in com["resultado"]} == {"p400": 1.0, "n3a": 1.0, "n3b": 2.0, "t3": 1.0}, com


def test_validacao():
    assert cp.validar(None) == [] and cp.validar(CONV_P) == [] and cp.validar(CONV_B) == []
    erros = cp.validar({"tipos": [], "formatos": ["{altura}"], "altura": [9, 3]})
    assert len(erros) >= 3
    assert cp.validar({"tipos": [{"codigo": "A"}], "formatos": ["{tipo}{altura}{altura}{esforco}"]})
    assert vr.validar_config({"postes": {"tipos": []}}) and vr.validar_config({"postes": CONV_P}) == []
    assert vr.normalizar_config({"postes": CONV_P})["postes"] == CONV_P and "postes" not in vr.normalizar_config({})


def _node(expr, **dados):
    js = f"const CP=require('./static/convencao_postes.js'),AO=require('./static/anotacao_outros.js'),VR=require('./static/vinculacao_regras.js');const D={json.dumps(dados)};console.log(JSON.stringify({expr}));"
    r = subprocess.run([NODE, "-e", js], cwd=config.BASE_DIR, capture_output=True, text=True, timeout=60)
    assert r.returncode == 0, r.stderr
    return json.loads(r.stdout)


@pytest.mark.skipif(not NODE, reason="sem node")
def test_paridade_com_o_javascript():
    for conv in (None, cp.PADRAO, CONV_P, CONV_B):
        for n in NOMES + EXTRA:
            assert cp.compacto_do_poste(n, conv) == _node("CP.compactoDoPoste(D.n, D.c)", n=n, c=conv), (n, conv)
            assert cp.testar(n, conv, {"DT1/1600", "B1/1600"}) == _node("CP.testar(D.n, D.c, new Set(['DT1/1600','B1/1600']))", n=n, c=conv), (n, conv)
    assert cp.validar({"tipos": [], "formatos": ["{altura}"], "altura": [9, 3]}) == _node("CP.validar(D.c)", c={"tipos": [], "formatos": ["{altura}"], "altura": [9, 3]})
    base = [L("P11-400", "P11-400", "p400"), L("N3=N3(poste=P11-400)", "N3", "n3a")]
    for t in ("P11-400, 2 N3, T3", "P11-400, N3"):
        py = ao.interpretar_linha(t, "I", calc._classico_outros, ao.contexto_da_base(base, "XYZ", CONV_P))
        js = _node("(()=>{const ctx=AO.contextoDaBase(D.b,'XYZ',D.c);return AO.interpretarLinha(D.t,'I',x=>x.split(/\\s+/).filter(Boolean).map(k=>({ativo:k.replace(/^\\d+-/,''),qtd:1})),ctx).itens.map(i=>i.ativo)})()", b=base, c=CONV_P, t=t)
        assert [i["ativo"] for i in py["itens"]] == js, (t, js)


def test_ligacoes_na_tela():
    import os
    html = open(os.path.join(config.STATIC_DIR, "resumo.html"), encoding="utf-8").read()
    assert html.index("/static/convencao_postes.js") < html.index("/static/anotacao_outros.js")
    ui = open(os.path.join(config.STATIC_DIR, "vinculacao_ui.js"), encoding="utf-8").read()
    assert "abaPostes" in ui and "Testar um nome" in ui
    js = open(os.path.join(config.STATIC_DIR, "resumo.js"), encoding="utf-8").read()
    assert "window.__vinculacaoConfig.postes" in js
