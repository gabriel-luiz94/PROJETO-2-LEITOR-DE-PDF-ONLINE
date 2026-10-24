"""TASK-077 — motor (static/modal_montavel_engine.js, Node) e API/validador (routers/modais_montaveis.py) dos modais montáveis."""
import json
import os
import subprocess

import pytest
from fastapi.testclient import TestClient

import config
import database
from middleware.auth_middleware import create_jwt_token
from routers import modais_montaveis as mm
from tests import oraculo_tela as o

MOTOR = os.path.join(config.STATIC_DIR, "modal_montavel_engine.js")
node = pytest.mark.skipif(not o.NODE, reason="node não instalado")

RAMAIS = {"id": "ramais", "nome": "Ramais colados", "tipo": "interpretativo", "blocos": [
    {"tipo": "texto", "texto": "Cole as linhas do Excel."},
    {"tipo": "grade", "id": "entrada", "colunas": [{"id": "texto", "rotulo": "TEXTO"}]},
    {"tipo": "interpretar", "id": "saida", "fonte": "entrada", "regras": [
        {"quando": {"coluna": "texto", "regex": r"TROCAR\s+(\d+)\s+RS M AC"},
         "gera": [{"ativo": "MAC", "qtd": "{1}*20", "operacao": "I"}, {"ativo": "MAC", "qtd": "{1}*15", "operacao": "R"}]},
        {"quando": {"regex": r"CP[- ]?REDE"}, "gera": [{"ativo": "CPREDE"}]}]},
    {"tipo": "destino", "id": "dest", "padrao": "outros"},
    {"tipo": "acao_adicionar", "fonte": "saida", "destino": "dest", "rotulo": "Adicionar ativos"}]}

TITULOS = {"id": "titulos", "nome": "Títulos de contrato", "tipo": "gerador_texto", "blocos": [
    {"tipo": "campo", "id": "regiao", "rotulo": "Região", "entrada": "select", "opcoes": ["NORTE", "SUL"]},
    {"tipo": "campo", "id": "obs", "rotulo": "Observação", "entrada": "texto", "visivel_se": {"campo": "regiao", "op": "igual", "valor": "SUL"}},
    {"tipo": "grade", "id": "itens", "colunas": [{"id": "codigo", "rotulo": "CÓDIGO"}, {"id": "qtd", "rotulo": "QTD"}]},
    {"tipo": "tabela_interna", "id": "precos", "colunas": ["codigo", "titulo", "contrato"],
     "linhas": [["MAC", "Mufla AC", "C-100"], ["TAM", "Terminal AM", "C-200"]]},
    {"tipo": "saida_texto", "id": "txt", "fonte": "itens", "rotulo": "Texto",
     "modelo": "{{codigo}} x{{qtd}} - {{buscar(precos, codigo, codigo, titulo)}} ({{buscar(precos, codigo, codigo, contrato)}}) {{regiao}}"}]}


def js(expr, **entrada):
    codigo = f"const M=require({json.dumps(MOTOR)});const e=JSON.parse(require('fs').readFileSync(0,'utf8'));console.log(JSON.stringify({expr}));"
    r = subprocess.run([o.NODE, "-e", codigo], input=json.dumps(entrada), capture_output=True, text=True, timeout=30)
    assert r.returncode == 0, r.stderr
    return json.loads(r.stdout)


# ── motor ──
@node
def test_calcular_so_aritmetica():
    assert js("M.calcular('2+3*4')") == 14
    assert js("M.calcular('(2+3)*4/2')") == 10
    assert js("M.calcular('-3+5')") == 2
    for ruim in ["process.exit()", "1/0", "2**3", "alert(1)", "", "1+"]:
        assert js("M.calcular(e.x)", x=ruim) is None


@node
def test_expressao_com_capturas():
    assert js("M.avaliarExpressao('{1}*20', ['x','3'])") == 60
    assert js("M.avaliarExpressao('{1}*20', ['x'])") == 0
    assert js("M.avaliarExpressao(5, [])") == 5


@node
def test_interpretar_ramais_soma_e_agrupa():
    linhas = [{"texto": "TROCAR 2 RS M AC"}, {"texto": "TROCAR 1 RS M AC"}, {"texto": "CP-REDE no poste"}, {"texto": "nada"}]
    r = js("M.interpretar(e.b, e.l)", b=RAMAIS["blocos"][2], l=linhas)
    assert r == [{"ativo": "MAC", "qtd": 60, "operacao": "I"}, {"ativo": "MAC", "qtd": 45, "operacao": "R"},
                 {"ativo": "CPREDE", "qtd": 1, "operacao": "I"}]


@node
def test_interpretar_sem_agrupar_e_parar_false():
    b = {"agrupar": False, "regras": [{"quando": {"regex": "A(\\d)"}, "parar": False, "gera": [{"ativo": "X{1}"}]},
                                       {"quando": {"regex": "A"}, "gera": [{"ativo": "Y"}]}]}
    r = js("M.interpretar(e.b, e.l)", b=b, l=[{"t": "A5"}])
    assert [(x["ativo"], x["qtd"]) for x in r] == [("X5", 1), ("Y", 1)]


@node
def test_primeira_regra_que_casa_vale_por_padrao():
    b = {"regras": [{"quando": {"regex": "A"}, "gera": [{"ativo": "PRIMEIRA"}]}, {"quando": {"regex": "A"}, "gera": [{"ativo": "SEGUNDA"}]}]}
    assert [x["ativo"] for x in js("M.interpretar(e.b, e.l)", b=b, l=[{"t": "A"}])] == ["PRIMEIRA"]


@node
def test_modelo_com_busca_e_linhas():
    t = {"precos": {"colunas": ["codigo", "titulo"], "linhas": [["MAC", "Mufla"]]}}
    assert js("M.modelo(e.m, {linha:{codigo:'mac'}, valores:{r:'SUL'}, tabelas:e.t})",
              m="{{codigo}} - {{buscar(precos, codigo, codigo, titulo)}} - {{r}} - {{nao_existe}}", t=t) == "mac - Mufla - SUL - "
    assert js("M.modelo('{{buscar(precos, codigo, \"zzz\", titulo)}}|', {tabelas:e.t})", t=t) == "|"


@node
def test_gerar_texto_por_linha_ignora_linhas_vazias():
    b = TITULOS["blocos"][4]
    tabelas = {"precos": {"colunas": ["codigo", "titulo", "contrato"], "linhas": [["MAC", "Mufla AC", "C-100"]]}}
    linhas = [{"codigo": "MAC", "qtd": "2"}, {"codigo": "", "qtd": ""}, {"codigo": "XX", "qtd": "1"}]
    r = js("M.gerarTexto(e.b, e.v, e.t, e.l)", b=b, v={"regiao": "SUL"}, t=tabelas, l=linhas)
    assert r == "MAC x2 - Mufla AC (C-100) SUL\nXX x1 -  () SUL"


@node
def test_visivel_se():
    b = {"visivel_se": {"campo": "a", "op": "igual", "valor": "1"}}
    assert js("M.visivel(e.b, {a:'1'})", b=b) is True and js("M.visivel(e.b, {a:'2'})", b=b) is False
    assert js("M.visivel({}, {})") is True
    assert js("M.visivel({visivel_se:{campo:'a',op:'em',valor:['x','y']}}, {a:'y'})") is True
    assert js("M.visivel({visivel_se:{campo:'a',op:'vazio'}}, {})") is True


@node
def test_linhas_da_matriz():
    r = js("M.linhasDaMatriz(e.c, e.m)", c=TITULOS["blocos"][2]["colunas"], m=[["MAC", " 2 "], ["TAM"]])
    assert r == [{"codigo": "MAC", "qtd": "2"}, {"codigo": "TAM", "qtd": ""}]


# ── validador ──
def test_exemplos_validos():
    assert mm.validar({"versao": 1, "modais": [RAMAIS, TITULOS]}) == []


def _com(bloco_mod):
    return {"versao": 1, "modais": [bloco_mod]}


def test_validador_recusa_o_perigoso_e_o_quebrado():
    import copy
    m = copy.deepcopy(RAMAIS)
    m["blocos"][2]["regras"][0]["gera"][0]["qtd"] = "__import__('os')"
    assert any("qtd" in e for e in mm.validar(_com(m)))
    m = copy.deepcopy(RAMAIS); m["blocos"][2]["regras"][0]["quando"]["regex"] = "("
    assert any("regex" in e for e in mm.validar(_com(m)))
    m = copy.deepcopy(RAMAIS); m["blocos"][2]["fonte"] = "naoexiste"
    assert any("fonte" in e for e in mm.validar(_com(m)))
    m = copy.deepcopy(RAMAIS); m["blocos"][4]["destino"] = "x"
    assert any("destino" in e for e in mm.validar(_com(m)))
    m = copy.deepcopy(RAMAIS); m["tipo"] = "outro"
    assert any("tipo" in e for e in mm.validar(_com(m)))
    m = copy.deepcopy(TITULOS); m["blocos"][1]["visivel_se"]["campo"] = "fantasma"
    assert any("visivel_se" in e for e in mm.validar(_com(m)))
    m = copy.deepcopy(TITULOS); m["blocos"][3]["linhas"][0] = ["so", "dois"]
    assert any("linhas" in e for e in mm.validar(_com(m)))
    assert mm.validar({"versao": 2, "modais": []}) and mm.validar([])
    assert any("repetido" in e for e in mm.validar({"versao": 1, "modais": [RAMAIS, RAMAIS]}))


def test_expressao_segura():
    for ok in [3, 2.5, "{1}*20", "(1+{2})/2", "-4"]:
        assert mm.expressao_segura(ok), ok
    for ruim in [True, "", "1+", "2**3", "abs(1)", "a+1", "1;2", None, "x" * 200]:
        assert not mm.expressao_segura(ruim), ruim


# ── API ──
@pytest.fixture
def client(tmp_path, monkeypatch):
    monkeypatch.setattr(database, "DB_PATH", str(tmp_path / "t.db"))
    database.init_db()
    import app as appmod
    monkeypatch.setattr(mm, "get_supabase", lambda: None)
    return TestClient(appmod.app)


def hdr(role):
    return {"Authorization": f"Bearer {create_jwt_token('u', 'u@x.com', role)}"}


def test_api_vazio_gravar_ler_isolado_por_projeto(client):
    assert client.get("/api/modais-montaveis?projeto_codigo=229", headers=hdr("operador")).json()["dados"] == {"versao": 1, "modais": []}
    dados = {"versao": 1, "modais": [RAMAIS]}
    r = client.put("/api/modais-montaveis", json={"projeto_codigo": "229", "dados": dados}, headers=hdr("admin"))
    assert r.status_code == 200
    assert client.get("/api/modais-montaveis?projeto_codigo=229", headers=hdr("operador")).json()["dados"] == dados
    assert client.get("/api/modais-montaveis?projeto_codigo=027", headers=hdr("operador")).json()["dados"]["modais"] == []


def test_api_so_admin_grava_e_valida(client):
    corpo = {"projeto_codigo": "229", "dados": {"versao": 1, "modais": [RAMAIS]}}
    assert client.put("/api/modais-montaveis", json=corpo, headers=hdr("operador")).status_code == 403
    assert client.post("/api/modais-montaveis/validar", json=corpo, headers=hdr("operador")).status_code == 403
    assert client.post("/api/modais-montaveis/validar", json=corpo, headers=hdr("admin")).json() == {"erros": []}
    ruim = {"projeto_codigo": "229", "dados": {"versao": 1, "modais": [{"id": "x"}]}}
    r = client.put("/api/modais-montaveis", json=ruim, headers=hdr("admin"))
    assert r.status_code == 400 and r.json()["detail"]["erros"]
    assert client.get("/api/modais-montaveis?projeto_codigo=229", headers=hdr("admin")).json()["dados"]["modais"] == []
