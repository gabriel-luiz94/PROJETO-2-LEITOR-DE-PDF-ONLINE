"""TASK-123: a Totalizadora avalia o NOME do ativo separado dos PARÂMETROS (static/ativo_avaliar.js, parte pura em Node)."""
import json
import os
import shutil
import subprocess

import pytest

import config

NODE = shutil.which("node")
pytestmark = pytest.mark.skipif(not NODE, reason="node indisponível")
MODULO = os.path.join(config.STATIC_DIR, "ativo_avaliar.js")


def _l(ativo, comp="", cod="", desc="", dcod="", origem="", proj="PERNAMBUCO"):
    return {"ativo": ativo, "desc_ativo": desc, "componente": comp, "codigo": cod, "desc_codigo": dcod, "projeto": proj, "origem": origem}


BASE = [
    _l("FLABIT=FLABIT(poste=D,cabobt=35)"),                                              # só declaração
    _l("FLABIT", "FLABIT", "3417025", "FLABIT", "ARMACAO"),
    _l("FLABIT(poste=%D)", "FLABIT.D", "3480305", "FLABIT POSTE D", "PARAFUSO"),
    _l("FLABIT(cabobt=35)", "FLABIT.C35", "3416055", "FLABIT 35", "CINTA"),
    _l("IT_2=IT_2(poste=D), IT-2"), _l("IT_2", "IT_2", "9001", "IT 2", "X"), _l("IT_2(poste=%D)", "IT_2.D", "9002", "IT 2 D", "Y"),
    _l("CFU", "CFU", "200", "CHAVE", "CHAVE FUSIVEL", "OUTROS"),
    _l("P50", "C1", "100", "CABO", "CABO P50", "CABOS"),
]


def _av(textos, origem="OUTROS", projeto="PERNAMBUCO"):
    codigo = f"const A=require({json.dumps(MODULO)});const e=JSON.parse(require('fs').readFileSync(0,'utf8'));const idx=A.construirIndice(e.base,e.proj);console.log(JSON.stringify(e.t.map(x=>A.avaliarAtivo(x,e.o,idx))));"
    r = subprocess.run([NODE, "-e", codigo], input=json.dumps({"base": BASE, "t": textos, "o": origem, "proj": projeto}), capture_output=True, text=True, timeout=30)
    assert r.returncode == 0, r.stderr
    return json.loads(r.stdout)


def test_ativo_com_parametros_e_avaliado_pelo_nome_e_depois_pelos_parametros():
    r = _av(["FLABIT", "FLABIT(D,35)", "flabit (poste=D, cabobt=35)", "FLABIT(poste=%D)"])
    assert r[0]["naoEncontrado"] is False and r[0]["desc"] == "FLABIT" and r[0]["erroParametros"] == ""
    assert r[1]["naoEncontrado"] is False and r[1]["desc"] == "FLABIT" and r[1]["erroParametros"] == ""         # antes: vermelho e sem descrição
    assert r[2]["naoEncontrado"] is False and r[2]["erroParametros"] == ""
    assert r[3]["erroParametros"] != ""                                                                          # "%D" é condição de grupo, não valor de pedido


def test_erros_verdadeiros_de_parametros_ficam_separados_do_nome():
    r = _av(["FLABIT(XX,35)", "FLABIT(D,35,7)", "FLABIT(poste=D,cabot=35)", "CFU(A)", "FLABIT(D,99)"])
    for x in r:
        assert x["naoEncontrado"] is False                                                                       # o nome existe em todos
    assert "POSTE=XX não cadastrado" in r[0]["erroParametros"]                                                  # fora do escopo cadastrado
    assert r[1]["erroParametros"] != "" and "FLABIT" in r[1]["erroParametros"]                                   # valores a mais
    assert "CABOT" in r[2]["erroParametros"] and "desconhecido" in r[2]["erroParametros"]                      # parâmetro não declarado
    assert "ativo sem parâmetros declarados" in r[3]["erroParametros"] and r[3]["nome"] == "CFU"
    assert "CABOBT=99 não cadastrado" in r[4]["erroParametros"]


def test_nome_inexistente_origem_projeto_apelido_componente_codigo_e_descricao():
    r = _av(["XYZ(D)", "P50", "IT-2(D)", "FLABIT.D", "3480305", "CHAVE", "FLABIT POSTE D"])
    assert r[0]["naoEncontrado"] is True and r[0]["desc"] == ""                                                 # nome inexistente: vermelho
    assert r[1]["naoEncontrado"] is True                                                                         # P50 só existe para CABOS
    assert r[2]["naoEncontrado"] is False and r[2]["nome"] == "IT_2" and r[2]["erroParametros"] == ""           # apelido com parâmetros
    assert r[3]["naoEncontrado"] is False and r[4]["naoEncontrado"] is False and r[4]["desc"] == "PARAFUSO"     # componente e código são nomes válidos
    assert r[5]["naoEncontrado"] is True and r[6]["naoEncontrado"] is True                                       # a descrição NÃO resolve nome
    assert _av(["P50"], origem="CABOS")[0]["naoEncontrado"] is False
    assert _av(["FLABIT"], projeto="RONDONIA")[0]["naoEncontrado"] is True                                       # outro projeto


def test_ligacao_na_tela():
    html = open(os.path.join(config.STATIC_DIR, "resumo.html"), encoding="utf-8").read()
    assert html.index("/static/ativo_parametros.js") < html.index("/static/ativo_avaliar.js") < html.index("/static/resumo.js")
    js = open(os.path.join(config.STATIC_DIR, "resumo.js"), encoding="utf-8").read()
    i = js.index("window.updateTotRow = async function")
    corpo = js[i:js.index("window.excluirTotRow")]
    assert "avaliarTot(" in corpo and "orcamentoBaseData.find" not in corpo                                      # a edição não compara mais o texto inteiro
    assert "avisoParametros" in js and "rgba(210, 153, 34" in js


def test_geracao_da_totalizadora_ambar_para_parametro_e_vermelho_para_nome_igual_no_js_e_no_python():
    """O JS real da tela (syncTotalizadora) e o porte em Python (modo autônomo) classificam igual: nome que existe + parâmetro ruim = âmbar."""
    from services.autonomo import totalizadora as tt
    from tests import oraculo_tela as o

    def l(ativo, comp, cod, desc="D"):
        return {"ativo": ativo, "desc_ativo": desc if ativo else "", "componente": comp, "projeto": "P1", "codigo": cod, "desc_codigo": "d"}
    base = [l("N1=N1(cabo=A)", "N1", "BASE"), l("N1(cabo=A)", "N1.A", "FIOA"), l("N1(cabo=P)", "N1.P", "FIOP"), l("CFU", "CFU", "200")]
    outros = [{"entidade": "ESTRUTURA", "operacao": "I", "ativo": t} for t in ("1-N1(A)", "1-N1(ZZ)", "1-NAOEXISTE(A)", "1-CFU(A)", "1-CFU")]
    js = o.totalizadora_e_payload([], outros, [], base, "P1")
    py = tt.normalizar_para_json(tt.montar_totalizadora([], outros, [], base, "P1"))
    assert py == js["totalizadora"]
    estado = {r["ativo"]: (r["naoEncontrado"], bool(r.get("avisoParametros"))) for r in py}
    assert estado["N1(A)"] == (False, False)
    assert estado["N1(ZZ)"] == (False, True) or estado.get("N1(ZZ)".upper()) == (False, True)       # nome existe, parâmetro fora do escopo: âmbar
    assert estado["NAOEXISTE(A)"] == (True, False)                                                   # nome inexistente: vermelho
    assert estado["CFU(A)"] == (False, True) and estado["CFU"] == (False, False)                     # ativo sem parâmetros declarados: âmbar
