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


# ── E3: condições compostas, regras condicionais e [campo] nas quantidades ──
@node
def test_condicao_composta_e_ou_nao():
    c = {"todos": [{"campo": "a", "op": "igual", "valor": "1"}, {"algum": [{"campo": "b", "op": "preenchido"}, {"campo": "c", "op": "maior", "valor": 5}]}]}
    assert js("M.avaliarCondicao(e.c, e.v)", c=c, v={"a": "1", "b": "", "c": "7"}) is True
    assert js("M.avaliarCondicao(e.c, e.v)", c=c, v={"a": "1", "b": "", "c": "3"}) is False
    assert js("M.avaliarCondicao(e.c, e.v)", c=c, v={"a": "2", "b": "x", "c": "9"}) is False
    assert js("M.avaliarCondicao({nao: e.c}, e.v)", c=c, v={"a": "2"}) is True
    assert js("M.avaliarCondicao(null, {})") is True


@node
def test_operadores_numericos_e_contem():
    f = lambda op, valor, v: js("M.avaliarCondicao({campo:'x',op:e.op,valor:e.val}, {x:e.v})", op=op, val=valor, v=v)
    assert f("maior", 5, "5,5") is True and f("maior", 5, "5") is False
    assert f("maior_igual", 5, "5") is True and f("menor", 5, "4.9") is True and f("menor_igual", 5, "6") is False
    assert f("maior", 5, "abc") is False and f("menor", 5, "") is False          # não numérico nunca satisfaz
    assert f("contem", "mac", "Mufla MAC 3") is True and f("contem", "zz", "abc") is False


@node
def test_visivel_usa_a_mesma_linguagem():
    b = {"visivel_se": {"todos": [{"campo": "a", "op": "igual", "valor": "1"}, {"nao": {"campo": "b", "op": "vazio"}}]}}
    assert js("M.visivel(e.b, {a:'1', b:'x'})", b=b) is True and js("M.visivel(e.b, {a:'1', b:''})", b=b) is False


@node
def test_campo_na_quantidade_e_no_ativo():
    assert js("M.avaliarExpressao('{1}*[fator]', ['x','3'], {fator:'2,5'})") == 7.5
    assert js("M.avaliarExpressao('[fator]+1', [], {fator:'abc'})") == 1        # não numérico = 0
    assert js("M.avaliarExpressao('[nao_existe]*2', [], {})") is None
    b = {"regras": [{"quando": {"regex": "(\\d+)"}, "gera": [{"ativo": "P{1}-[regiao]", "qtd": "[fator]"}]}]}
    r = js("M.interpretar(e.b, e.l, e.v)", b=b, l=[{"t": "poste 12"}], v={"regiao": "SUL", "fator": "3"})
    assert r == [{"ativo": "P12-SUL", "qtd": 3, "operacao": "I"}]


@node
def test_regra_com_se_so_vale_quando_a_condicao_dos_campos_e_verdadeira():
    b = {"regras": [{"se": {"campo": "corr", "op": "igual", "valor": "sim"}, "quando": {"regex": "RS"}, "gera": [{"ativo": "COM_CORROSAO"}]},
                    {"quando": {"regex": "RS"}, "gera": [{"ativo": "NORMAL"}]}]}
    sim = js("M.interpretar(e.b, e.l, {corr:'sim'})", b=b, l=[{"t": "RS"}])
    nao = js("M.interpretar(e.b, e.l, {corr:'nao'})", b=b, l=[{"t": "RS"}])
    assert [x["ativo"] for x in sim] == ["COM_CORROSAO"] and [x["ativo"] for x in nao] == ["NORMAL"]


def test_validador_condicoes_compostas_e_referencias():
    import copy
    base = copy.deepcopy(TITULOS)
    base["blocos"][1]["visivel_se"] = {"todos": [{"campo": "regiao", "op": "igual", "valor": "SUL"}, {"nao": {"campo": "regiao", "op": "vazio"}}]}
    assert mm.validar(_com(base)) == []
    base["blocos"][1]["visivel_se"] = {"algum": [{"campo": "fantasma", "op": "vazio"}]}
    assert any("fantasma" in e for e in mm.validar(_com(base)))
    base["blocos"][1]["visivel_se"] = {"algum": []}
    assert any("algum" in e for e in mm.validar(_com(base)))
    base["blocos"][1]["visivel_se"] = {"campo": "regiao", "op": "maior", "valor": "abc"}
    assert any("número" in e for e in mm.validar(_com(base)))
    fundo = {"campo": "regiao", "op": "vazio"}
    for _ in range(6):
        fundo = {"nao": fundo}
    base["blocos"][1]["visivel_se"] = fundo
    assert any("aninhada" in e for e in mm.validar(_com(base)))


def test_validador_regra_com_se_e_campo_na_qtd():
    import copy
    m = copy.deepcopy(RAMAIS)
    m["blocos"].insert(0, {"tipo": "campo", "id": "fator", "rotulo": "Fator", "entrada": "numero"})
    m["blocos"][3]["regras"][0]["gera"][0]["qtd"] = "{1}*[fator]"
    m["blocos"][3]["regras"][0]["se"] = {"campo": "fator", "op": "maior", "valor": 0}
    assert mm.validar(_com(m)) == []
    m["blocos"][3]["regras"][0]["gera"][0]["qtd"] = "{1}*[inexistente]"
    assert any("inexistente" in e for e in mm.validar(_com(m)))
    m["blocos"][3]["regras"][0]["gera"][0]["qtd"] = "{1}"
    m["blocos"][3]["regras"][0]["se"] = {"campo": "nada", "op": "vazio"}
    assert any("se.campo" in e for e in mm.validar(_com(m)))
    m["blocos"][3]["regras"][0]["se"] = None
    m["blocos"][3]["regras"][0]["gera"][0]["ativo"] = "P[fantasma]"
    assert any("fantasma" in e for e in mm.validar(_com(m)))


def test_expressao_segura_aceita_campo():
    assert mm.expressao_segura("{1}*[fator]") and mm.expressao_segura("[a]+[b-2]/2")
    assert not mm.expressao_segura("[a b]") and not mm.expressao_segura("[a]**2")


# ── E4: colagem, ocorrências, normalização, linhas não reconhecidas, carregar da tabela ──
@node
def test_ajustar_colagem_modos():
    f = lambda col, m: js("M.ajustarColagem({colagem: e.c}, e.m)", c=col, m=m)
    assert f({"modo": "colunas"}, [["a", "b"]]) == [["a", "b"]]
    assert f({"modo": "linha"}, [["a", "b c"], ["d"]]) == [["a b c"], ["d"]]
    assert f({"modo": "dividir", "separador": ";"}, [["x; y ;z"]]) == [["x", "y", "z"]]
    assert f({"modo": "dividir"}, [["x;y"]]) == [["x", "y"]]                       # separador padrão
    assert f({"modo": "colunas", "pular_cabecalho": True}, [["H1", "H2"], ["a", "b"]]) == [["a", "b"]]
    assert js("M.ajustarColagem({}, [['a']])") == [["a"]]


@node
def test_regra_global_gera_um_item_por_ocorrencia():
    b = {"regras": [{"quando": {"regex": "P(\\d+)", "global": True}, "gera": [{"ativo": "P{1}"}]}]}
    assert [x["ativo"] for x in js("M.interpretar(e.b, e.l, {})", b=b, l=[{"t": "P1, P2 e P3"}])] == ["P1", "P2", "P3"]
    sem = {"regras": [{"quando": {"regex": "P(\\d+)"}, "gera": [{"ativo": "P{1}"}]}]}
    assert [x["ativo"] for x in js("M.interpretar(e.b, e.l, {})", b=sem, l=[{"t": "P1, P2 e P3"}])] == ["P1"]
    vazio = {"regras": [{"quando": {"regex": "x*", "global": True}, "gera": [{"ativo": "A"}]}]}
    assert js("M.interpretar(e.b, e.l, {})", b=vazio, l=[{"t": "abc"}])           # regex que casa vazio não trava


@node
def test_normalizar_e_nao_reconhecidas():
    b = {"normalizar": "maiusculo", "regras": [{"quando": {"regex": "mac"}, "gera": [{"ativo": "mac-1"}]}]}
    d = js("M.diagnosticar(e.b, e.l, {})", b=b, l=[{"t": "MAC"}, {"t": "zzz"}, {"t": "   "}])
    assert [x["ativo"] for x in d["itens"]] == ["MAC-1"]
    assert d["naoCasadas"] == [{"linha": 1, "texto": "zzz"}]                         # linha em branco não conta
    assert js("M.interpretar({normalizar:'minusculo',regras:[{quando:{regex:'A'},gera:[{ativo:'XyZ'}]}]}, [{t:'a'}], {})")[0]["ativo"] == "xyz"


def test_validador_e4():
    import copy
    m = copy.deepcopy(RAMAIS)
    m["blocos"][1]["colagem"] = {"modo": "dividir", "separador": ";", "pular_cabecalho": True}
    m["blocos"][1]["carregar_de"] = "outros"
    m["blocos"][2]["normalizar"] = "maiusculo"
    m["blocos"][2]["regras"][0]["quando"]["global"] = True
    assert mm.validar(_com(m)) == []
    for caminho, valor, trecho in [(("blocos", 1, "colagem"), {"modo": "x"}, "colagem"), (("blocos", 1, "carregar_de"), "tudo", "carregar_de"),
                                   (("blocos", 2, "normalizar"), "capitalizar", "normalizar")]:
        r = copy.deepcopy(m); alvo = r
        for c in caminho[:-1]:
            alvo = alvo[c]
        alvo[caminho[-1]] = valor
        assert any(trecho in e for e in mm.validar(_com(r))), caminho
    r = copy.deepcopy(m); r["blocos"][2]["regras"][0]["quando"]["global"] = "sim"
    assert any("global" in e for e in mm.validar(_com(r)))
    r = copy.deepcopy(m); r["blocos"][1]["colagem"]["separador"] = "x" * 11
    assert any("separador" in e for e in mm.validar(_com(r)))
