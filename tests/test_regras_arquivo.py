"""TASK-061 — regras do leitor e ajustes em arquivo JSON: validação no servidor, lógica pura do módulo JS (via Node) e ida e volta."""
import json
import subprocess

import pytest
from fastapi.testclient import TestClient

import app as appmod
import database
import routers.regras_leitor as rl
from middleware.auth_middleware import create_jwt_token
from tests import oraculo_tela

with open("data/regras_leitor_processamento_seed.json", encoding="utf-8") as f:
    PROC = json.load(f)
with open("data/regras_leitor_classificacao_seed.json", encoding="utf-8") as f:
    CLS = json.load(f)


@pytest.fixture
def client(tmp_path, monkeypatch):
    monkeypatch.setattr(database, "DB_PATH", str(tmp_path / "t.db"))
    database.init_db()
    return TestClient(appmod.app)


def cab(role="admin", uid="adm"):
    return {"Authorization": f"Bearer {create_jwt_token(uid, f'{uid}@x.com', role)}"}


# ── servidor: validar a importação sem gravar ───────────────────────────────
def test_sementes_reais_passam_na_validacao_de_importacao():
    assert rl.validar_importacao("processamento", PROC) == [] and rl.validar_importacao("classificacao", CLS) == []


@pytest.mark.parametrize("tabela,regras,trecho", [
    ("classificacao", "x", "lista de regras"),
    ("classificacao", [{"ordem": 1}] * 1001, "Regras demais"),
    ("classificacao", [{"ordem": 1, "ativo_regex": "a" * 501}], "passa de 500"),
    ("processamento", [{"fase": 3, "ordem": 1, "texto_regex": "a" * 501}], "passa de 500"),
    ("processamento", [{"fase": 3, "ordem": 1, "vizinhanca": {"regex": "a" * 501}}], "vizinhanca.regex"),
    ("classificacao", [{"ordem": "x"}], "ordem"),
    ("classificacao", [{"ordem": 1, "ativo_regex": "("}], "regex inválida"),
    ("processamento", [{"fase": 9, "ordem": 1}], "fase"),
    ("classificacao", [1], "objeto"),
])
def test_validacao_recusa_com_mensagem(tabela, regras, trecho):
    assert any(trecho in e for e in rl.validar_importacao(tabela, regras))


def test_rotas_validar_nao_gravam_e_exigem_admin(client):
    antes = client.get("/api/regras-leitor/classificacao?projeto_codigo=229", headers=cab()).json()["regras"]
    r = client.post("/api/regras-leitor/classificacao/validar", json={"projeto_codigo": "229", "regras": CLS}, headers=cab())
    assert r.status_code == 200 and r.json() == {"erros": [], "total": len(CLS)}
    r = client.post("/api/regras-leitor/processamento/validar", json={"projeto_codigo": "229", "regras": [{"fase": 3, "ordem": 1, "texto_regex": "("}]}, headers=cab())
    assert r.status_code == 200 and r.json()["erros"]
    assert client.get("/api/regras-leitor/classificacao?projeto_codigo=229", headers=cab()).json()["regras"] == antes          # nada mudou
    assert client.post("/api/regras-leitor/classificacao/validar", json={"projeto_codigo": "229", "regras": CLS}, headers=cab("operador", "op")).status_code == 403
    assert client.post("/api/regras-leitor/classificacao/validar", json={"projeto_codigo": "229", "regras": CLS}).status_code == 401


def test_ajustes_validar_importacao(client):
    ok = {"id": "A1", "nome": "troca", "ativa": True, "acoes": [{"acao": "substituir", "tabela": "outros", "de": "A", "para": "B"}]}
    com_extras = {**ok, "id": "A2", "origem": "projeto", "oculta": False, "frases": ["x"]}                 # campos de tela não atrapalham
    r = client.post("/api/validacao/ajustes/validar", json={"projeto_codigo": "229", "ajustes": [ok, com_extras]}, headers=cab())
    assert r.status_code == 200 and r.json() == {"erros": [], "total": 2}
    ruim = client.post("/api/validacao/ajustes/validar", json={"projeto_codigo": "229", "ajustes": [{"id": "x", "nome": "", "ativa": True, "acoes": []}]}, headers=cab())
    assert ruim.json()["erros"]
    assert client.post("/api/validacao/ajustes/validar", json={"projeto_codigo": "229", "ajustes": [ok] * 1001}, headers=cab()).json()["erros"][0].startswith("Ajustes demais")
    assert client.post("/api/validacao/ajustes/validar", json={"projeto_codigo": "229", "ajustes": [ok]}, headers=cab("operador", "op")).status_code == 403


def test_ida_e_volta_pela_rota_normal_de_salvar_mantem_o_motor_igual(client):
    """Exporta (lista como está), importa (substituir) e grava pelo Salvar de sempre: o motor entrega o mesmo resultado e a versão anterior vai ao histórico."""
    from services.autonomo.leitor_js import LeitorJS
    pytest.importorskip("quickjs")
    atual = client.get("/api/regras-leitor/processamento?projeto_codigo=229", headers=cab()).json()["regras"]
    assert atual
    r = client.post("/api/regras-leitor/processamento", json={"projeto_codigo": "229", "regras": json.loads(json.dumps(atual))}, headers=cab())
    assert r.status_code == 200
    depois = client.get("/api/regras-leitor/processamento?projeto_codigo=229", headers=cab()).json()["regras"]
    assert depois == atual
    assert client.get("/api/regras-leitor/processamento/historico?projeto_codigo=229", headers=cab()).json()["historico"]
    lt = LeitorJS()
    try:
        itens = [{"pagina": 1, "texto": t, "cor": c, "layer": ""} for t in ("AFASTADOR", "DT11/300", "INST. 01 - U4", "10 METROS") for c in ("#ff0000", "#808080", "#000000")]
        cls = client.get("/api/regras-leitor/classificacao?projeto_codigo=229", headers=cab()).json()["regras"]
        assert lt.processar_lote(itens, atual, cls, "extracao") == lt.processar_lote(itens, depois, cls, "extracao")
    finally:
        lt.fechar()


# ── módulo JS (lógica pura, executada no Node) ──────────────────────────────
def js(corpo, entrada=None):
    codigo = ("const A = require('./static/regras_arquivo.js'); const e = JSON.parse(require('fs').readFileSync(0, 'utf8') || 'null'); "
              f"const sai = (() => {{ {corpo} }})(); console.log(JSON.stringify(sai));")
    r = subprocess.run([oraculo_tela.NODE, "-e", codigo], input=json.dumps(entrada), capture_output=True, text=True, timeout=60)
    assert r.returncode == 0, r.stderr
    return json.loads(r.stdout)


pytestmark_node = pytest.mark.skipif(not oraculo_tela.NODE, reason="node não instalado")


@pytestmark_node
def test_interpretar_aceita_envelope_e_lista_pura_e_recusa_o_resto():
    ok = js("return A.interpretar(JSON.stringify(A.montarEnvelope('regras','229','classificacao',e)), 'regras', 'classificacao')", CLS)
    assert ok["itens"] == CLS and ok["projeto"] == "229" and ok["envelope"] is True
    pura = js("return A.interpretar(JSON.stringify(e), 'regras', 'classificacao')", CLS)
    assert pura["itens"] == CLS and pura["projeto"] is None and pura["envelope"] is False

    def erro(texto, tipo="regras", tabela="classificacao"):
        return js("try { A.interpretar(e.t, e.tipo, e.tabela); return 'sem erro'; } catch (x) { return x.mensagens || String(x); }", {"t": texto, "tipo": tipo, "tabela": tabela})
    assert "vazio" in erro("")[0] and "JSON válido" in erro("{x")[0]
    assert "outro cadastro" in erro(json.dumps({"formato": "leitor-ajustes", "versao": 1, "ajustes": []}))[0]
    assert "desconhecido" in erro(json.dumps({"formato": "x", "versao": 1}))[0]
    assert "mais nova" in erro(json.dumps({"formato": "leitor-regras", "versao": 2, "tabela": "classificacao", "regras": [{}]}))[0]
    assert "Versão" in erro(json.dumps({"formato": "leitor-regras", "versao": "1", "regras": [{}]}))[0]
    assert "tabela de processamento" in erro(json.dumps({"formato": "leitor-regras", "versao": 1, "tabela": "processamento", "regras": [{}]}))[0]
    assert "nenhuma regra" in erro("[]")[0] and "lista de regras" in erro("42")[0] and "não tem a lista" in erro(json.dumps({"formato": "leitor-regras", "versao": 1}))[0]
    assert "Itens demais" in erro(json.dumps([{}] * 1001))[0]
    assert "objeto" in erro(json.dumps([1]))[0] and "passa de 500" in erro(json.dumps([{"ativo_regex": "a" * 501}]))[0]
    assert "grande demais" in erro("[" + " " * (1024 * 1024) + "]")[0]


@pytestmark_node
def test_resumo_e_aplicar_regras_do_leitor_acrescentar_e_substituir():
    atual = [{"fase": 3, "ordem": 10, "modo": "DEFINIR", "operacao": "I"}, {"fase": 1, "ordem": 20, "modo": "DEFINIR", "texto_regex": "A", "ativo_template": "X"}]
    imp = [{"fase": 3, "ordem": 99, "modo": "DEFINIR", "operacao": "I"},                       # igual à existente (a ordem não conta)
           {"fase": 3, "ordem": 5, "modo": "DEFINIR", "operacao": "R"}, {"fase": 1, "ordem": 7, "modo": "DEFINIR", "texto_regex": "B", "ativo_template": "Y"}]
    corpo = "return {r: A.resumir(e.a, e.i, e.m, {}), n: A.aplicar(e.a, e.i, e.m, {porFase: true})}"
    ac = js(corpo, {"a": atual, "i": imp, "m": "acrescentar"})
    assert ac["r"] == {"novas": 2, "atualizadas": 0, "iguais": 1, "removidas": 0}
    assert ac["n"][:2] == atual and [(x["fase"], x["ordem"]) for x in ac["n"][2:]] == [(3, 20), (1, 30)]    # renumeradas depois da última, por fase, na ordem do arquivo
    assert js("return A.aplicar(e.a, e.i, 'acrescentar', {porFase: true})", {"a": atual, "i": imp + imp}) == js("return A.aplicar(e.a, e.i, 'acrescentar', {porFase: true})", {"a": atual, "i": imp})   # repetidas no arquivo entram uma vez
    sub = js(corpo, {"a": atual, "i": imp, "m": "substituir"})
    assert sub["r"] == {"novas": 2, "atualizadas": 0, "iguais": 1, "removidas": 1} and sub["n"] == imp
    classif = js("return A.aplicar(e.a, e.i, 'acrescentar', {})", {"a": CLS, "i": [{"ordem": 1, "ativo_regex": "^Z$", "entidade": "X"}]})
    assert classif[-1]["ordem"] == max(r["ordem"] for r in CLS) + 10 and len(classif) == len(CLS) + 1


@pytestmark_node
def test_resumo_e_aplicar_ajustes_por_id():
    def aj(i, de="A"):
        return {"id": i, "nome": i, "ativa": True, "acoes": [{"acao": "substituir", "tabela": "outros", "de": de, "para": "B"}]}
    atual = [{**aj("A1"), "origem": "projeto", "frases": ["x"], "_aberta": True}, aj("A2")]
    imp = [aj("A1"), aj("A2", de="Z"), aj("A3")]
    opc = "{chave: 'id', limpar: r => { const c = {}; Object.keys(r).forEach(k => { if (!['origem','oculta','frases'].includes(k) && !k.startsWith('_')) c[k] = r[k]; }); return c; }}"
    corpo = f"const o = {opc}; return {{r: A.resumir(e.a, e.i, e.m, o), n: A.aplicar(e.a, e.i, e.m, o).map(x => x.id)}}"
    ac = js(corpo, {"a": atual, "i": imp, "m": "acrescentar"})
    assert ac["r"] == {"novas": 1, "atualizadas": 1, "iguais": 1, "removidas": 0} and ac["n"] == ["A1", "A2", "A3"]
    sub = js(corpo, {"a": atual + [aj("A9")], "i": imp[:2], "m": "substituir"})
    assert sub["r"]["removidas"] == 1 and sub["n"] == ["A1", "A2"]
    mantido = js(f"const o = {opc}; return A.aplicar(e.a, e.i, 'acrescentar', o)[1].acoes[0].de", {"a": atual, "i": imp})
    assert mantido == "Z"                                                                     # mesmo id, conteúdo diferente: o do arquivo vence


@pytestmark_node
def test_nome_de_arquivo_seguro_e_envelope():
    assert js("return [A.nomeDeArquivo('regras','229','classificacao'), A.nomeDeArquivo('ajustes','../x y'), A.nomeDeArquivo('regras','','processamento')]") == \
        ["regras-leitor-classificacao-229.json", "ajustes-x_y.json", "regras-leitor-processamento-projeto.json"]
    env = js("return A.montarEnvelope('ajustes','229',null,e)", [{"id": "A"}])
    assert env["formato"] == "leitor-ajustes" and env["versao"] == 1 and env["ajustes"] == [{"id": "A"}] and "tabela" not in env and "@" not in json.dumps(env)
