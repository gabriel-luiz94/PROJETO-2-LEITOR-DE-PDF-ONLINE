"""TASK-165 — parâmetros de obra (OBRA.*): normalização, derivação de OBRA.REGIAO, resolver_pedido, cálculo, validação e payload."""
import pytest

from services import ativo_parametros as ap
from services import obra_params as op
from services import orcamento_calc as calc


def L(ativo, comp="", cod="", fi=1.0, proj="PERNAMBUCO", **k):
    r = {"ativo": ativo, "desc_ativo": "d", "componente": comp, "projeto": proj, "mdo": "MATERIAL", "codigo": cod, "desc_codigo": f"d {cod}",
         "fator_i": fi, "fator_r": 0.0, "filtro": "", "origem": "OUTROS", "lv": "A", "arred": "NORMAL", "formula_i": "", "formula_r": ""}
    r.update(k)
    return r


BASE = ap.preencher_componentes([
    L("FLBIM=FLBIM(poste=D,obra.regiao=CENTRO)"), L("FLBIM", "FLBIM", "100", 1), L("FLBIM(poste=D)", "FLBIM.D", "101", 1),
    L("FLBIM(obra.regiao=CENTRO|CENTRO_OESTE|OESTE)", "FLBIM.ATERR", "2412001", 2), L("FLBIM(obra.regiao=SUL|METROPOLITANO|NORTE|RECIFE)"),
    L("FLBIT=FLBIT(poste=D,cabobt=35,obra.regiao=CENTRO)"), L("FLBIT(obra.regiao=CENTRO|CENTRO_OESTE|OESTE)", "FLBIT.ATERR", "2412000", 4), L("FLBIT(obra.regiao=SUL|METROPOLITANO|NORTE|RECIFE)"),
])[0]
SIG = ap.assinaturas_do_projeto(BASE, "PERNAMBUCO")


def total(pedido, obra=None):
    r = calc.processar_calculo([], [{"ativo": pedido, "operacao": "I", "entidade": "0"}], "PERNAMBUCO", BASE, obra=obra)
    return {x["codigo"]: x["total"] for x in r["resultado"]}, r["nao_encontrados"], r.get("avisos") or []


# ── normalização e derivação ──
def test_normalizacao():
    assert op.normalizar("Centro Oeste") == "CENTRO_OESTE"
    assert op.normalizar("34,5kV") == "34,5KV"
    assert op.normalizar("  São  José ") == "SAO_JOSE"


@pytest.mark.parametrize("obra,esperado", [
    ({"utep": "Oeste"}, "OESTE"), ({"utep": "CENTRO OESTE"}, "CENTRO_OESTE"),
    ({"utd": "JAB GUARARAPES"}, "SUL"), ({"municipio": "Caruaru"}, "CENTRO"), ({"utep": "xyz", "municipio": "Caruaru"}, "CENTRO"), ({}, ""),
])
def test_regiao_derivada(obra, esperado):
    assert op.regiao_de(obra, "PERNAMBUCO") == esperado


def test_projeto_sem_tabela_nao_deriva():
    assert op.regiao_de({"municipio": "Caruaru"}, "OUTRO") == ""


def test_montar_so_o_que_a_obra_tem():
    v = op.montar({"municipio": "Caruaru", "contexto": "34,5kV"}, "PERNAMBUCO")
    assert v == {"OBRA.MUNICIPIO": "CARUARU", "OBRA.CONTEXTO": "34,5KV", "OBRA.REGIAO": "CENTRO"}


# ── resolver_pedido ──
def test_obra_manda_no_parametro():
    r = ap.resolver_pedido("FLBIM", SIG, {"OBRA.REGIAO": "SUL"})
    assert r["params"]["OBRA.REGIAO"] == "SUL" and not r["erros"] and r["avisos"] == []
    assert r["texto"] == "FLBIM(D)" and r["obra"]["OBRA.REGIAO"] == {"valor": "SUL", "origem": "obra"}


def test_pedido_com_obra_x_e_erro():
    r = ap.resolver_pedido("FLBIM(obra.regiao=SUL)", SIG, {})
    assert any("vem do cabeçalho da obra" in e for e in r["erros"])


def test_ligacao_posicional_pula_obra():
    r = ap.resolver_pedido("FLBIT(D,35)", SIG, {"OBRA.REGIAO": "SUL"})
    assert r["params"]["POSTE"] == "D" and r["params"]["CABOBT"] == "35" and not r["erros"]


def test_ausente_usa_padrao_com_aviso_uma_vez():
    r = ap.resolver_pedido("FLBIM", SIG, {})
    assert r["params"]["OBRA.REGIAO"] == "CENTRO" and r["avisos"] == ["OBRA.REGIAO não informado na obra; usando CENTRO"]
    assert ap.resolver_pedido("FLBIM", SIG)["avisos"] == []                       # chamador sem obra: padrão, sem aviso


def test_valor_nao_cadastrado_diz_a_origem():
    r = ap.resolver_pedido("FLBIM", SIG, {"OBRA.REGIAO": "MARTE"})
    assert any("OBRA.REGIAO=MARTE (da obra) não cadastrado em FLBIM" in e for e in r["erros"])


# ── cálculo ──
def test_regional_sul_sem_aterramento():
    t, nao, _ = total("FLBIM", {"utep": "SUL"})
    assert "2412001" not in t and not nao


def test_regional_centro_oeste_com_aterramento():
    t, nao, av = total("FLBIM", {"utep": "CENTRO OESTE"})
    assert t["2412001"] == 2 and not nao and not [a for a in av if "OBRA." in a]


def test_municipio_pela_tabela():
    t, _, _ = total("FLBIM", {"municipio": "CARUARU"})
    assert "2412001" in t


def test_sem_cabecalho_usa_padrao_e_avisa():
    t, _, av = total("FLBIM", {})
    assert "2412001" in t and av.count("OBRA.REGIAO não informado na obra; usando CENTRO") == 1
    t2, _, av2 = total("FLBIM")                                                   # payload antigo
    assert t2 == t and av2 == av


def test_pedido_com_obra_x_vira_nao_encontrado():
    _, nao, _ = total("FLBIM(obra.regiao=SUL)", {"utep": "SUL"})
    assert any("vem do cabeçalho da obra" in n for n in nao)


def test_flbit_posicional():
    t, nao, _ = total("FLBIT(D,35)", {"utep": "OESTE"})
    assert t.get("2412000") == 4 and not nao


# ── validação da base ──
def test_validador_recusa_obra_desconhecido():
    r = ap.validar_base([L("X=X(obra.foo=A)", "c", "1")])
    assert any("não é parâmetro de obra" in e for e in r["erros"])
    assert not [e for e in ap.validar_base(BASE)["erros"] if "OBRA" in e]


# ── API ──
def test_payload_aceita_obra_e_sem_obra():
    from models import OrcamentoRequest
    assert OrcamentoRequest(cabos=[], outros=[]).obra is None
    assert OrcamentoRequest(cabos=[], outros=[], obra={"utep": "SUL"}).obra == {"utep": "SUL"}


# ── .obra guarda o contexto ──
def test_obra_guarda_e_preserva_contexto():
    import json
    from services import projeto_obra as po
    t1 = po.montar({"nome": "P", "contexto": "34,5kV"}, None, "229")
    assert json.loads(t1)["projeto_trabalho"]["contexto"] == "34,5kV" and po.ler(t1, "229")["contexto"] == "34,5kV"
    t2 = po.montar({"nome": "P2"}, json.loads(t1), "229")                        # parcial sem contexto preserva o do arquivo
    assert json.loads(t2)["projeto_trabalho"]["contexto"] == "34,5kV"
    t3 = po.montar({"contexto": ""}, json.loads(t2), "229")                      # vazio = nenhum
    assert "contexto" not in json.loads(t3)["projeto_trabalho"]
    assert len(json.loads(po.montar({"nome": "P", "contexto": "x" * 200}, None, "229"))["projeto_trabalho"]["contexto"]) == 80


def test_obra_antiga_sem_contexto_abre():
    import json
    from services import projeto_obra as po
    antigo = json.loads(po.montar({"nome": "P"}, None, "229"))
    assert "contexto" not in antigo["projeto_trabalho"]
    assert po.ler(json.dumps(antigo), "229")["contexto"] == ""


def test_endpoint_regionais():
    from fastapi.testclient import TestClient
    import app
    from middleware.auth_middleware import create_jwt_token
    c = TestClient(app.app)
    h = {"Authorization": "Bearer " + create_jwt_token("1", "t@t", False)}
    r = c.get("/api/orcamento/regionais?projeto=PERNAMBUCO", headers=h)
    assert r.status_code == 200 and "CENTRO_OESTE" in r.json()["regioes"]
    assert c.get("/api/orcamento/regionais?projeto=OUTRO", headers=h).json() == {}


# ── paridade Python × JS ──
import json as _json
import os as _os
import subprocess as _sp

import config as _config
from tests import oraculo_tela as _o

_JS = _os.path.join(_config.STATIC_DIR, "ativo_parametros.js")
_JSO = _os.path.join(_config.STATIC_DIR, "obra_params.js")


@pytest.mark.skipif(not _o.NODE, reason="node não instalado")
def test_paridade_resolver_pedido_e_regiao():
    casos = [("FLBIM", {}), ("FLBIM", {"OBRA.REGIAO": "SUL"}), ("FLBIM", None), ("FLBIM(obra.regiao=SUL)", {}), ("FLBIT(D,35)", {"OBRA.REGIAO": "OESTE"}),
             ("FLBIM", {"OBRA.REGIAO": "MARTE"}), ("FLBIT(D,35,CENTRO)", {})]
    obras = [{"utep": "Centro Oeste"}, {"utd": "JAB GUARARAPES"}, {"municipio": "Caruaru"}, {}, {"municipio": "Recife", "bairro": "Boa Viagem"}]
    codigo = (f"const AP=require({_json.dumps(_JS)}),OB=require({_json.dumps(_JSO)});const e=JSON.parse(require('fs').readFileSync(0,'utf8'));"
              "const sig=AP.assinaturasDoProjeto(e.base,'PERNAMBUCO');"
              "console.log(JSON.stringify({p:e.casos.map(([t,o])=>AP.resolverPedido(t,sig,o)),r:e.obras.map(o=>OB.montar(o,e.tabela))}));")
    r = _sp.run([_o.NODE, "-e", codigo], input=_json.dumps({"base": BASE, "casos": casos, "obras": obras, "tabela": op.tabela_regional("PERNAMBUCO")}), capture_output=True, text=True, timeout=30)
    assert r.returncode == 0, r.stderr
    js_ = _json.loads(r.stdout)
    for (t, o), got in zip(casos, js_["p"]):
        assert got == ap.resolver_pedido(t, SIG, o), (t, o)
    for ob, got in zip(obras, js_["r"]):
        assert got == op.montar(ob, "PERNAMBUCO"), ob
