"""TASK-248 — cava em rocha (SDEMU3007II) igual à aba CAVA_ROCHA da ICP: os 5.687 pontos da grade e os 7 casos conferidos no LibreOffice,
Python e JS com a mesma saída, parte `cava_rocha` do pacote (só o projeto que a tem mostra o botão) e a linha lançada em Outros."""
import json
import os
import re
import shutil
import subprocess

import pytest
from fastapi.testclient import TestClient

import config
import database
from middleware.auth_middleware import create_jwt_token
from services import cava_rocha as cr
from services import pacote_projeto as pp

RAIZ = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
ANEXOS = os.path.join(RAIZ, ".ai", "tasks", "anexos", "TASK-248")
NODE = shutil.which("node")


def _json(caminho):
    with open(caminho, encoding="utf-8") as f:
        return json.load(f)


CASOS = _json(os.path.join(ANEXOS, "casos_cava_icp.json"))
DADOS = _json(os.path.join(ANEXOS, "cava_rocha_pernambuco.json"))
TABELA = [(l["localidade"], l["indice"]) for l in DADOS["localidades"]]


def test_grade_e_casos_da_icp_em_python():
    assert len(CASOS["grade"]) == 5687 and len(CASOS["casos"]) == 7
    for g in CASOS["grade"]:
        assert cr.calcular(g["localidade"], {"SDEMU1004II": g["postes_estais"]}, TABELA)["cavas"] == g["cavas_icp"], g
    for c in CASOS["casos"]:
        assert cr.calcular(c["localidade"], c, TABELA)["cavas"] == c["SDEMU3007II_icp"], c


def test_exemplos_da_task():
    def q(loc, a, b, c):
        return cr.calcular(loc, {"SDEMU1004II": a, "SDEMU3002II": b, "SDEMU1005II": c}, TABELA)
    assert q("AGRESTINA", 10, 0, 0) == {"f": 10, "cavas": 6}
    assert q("BETANIA", 3, 2, 1)["cavas"] == 6                       # 7,3 → 7, limitado a F = 6
    assert q("PALMARES (PAL)", 25, 3, 0)["cavas"] == 2 and q("RIBEIRAO", 40, 5, 2)["cavas"] == 4 and q("CUSTODIA", 7, 1, 0)["cavas"] == 8
    assert q("NAO EXISTE", 10, 0, 0)["cavas"] == 0
    assert cr.calcular("AGRESTINA", {"SDEMU1003II": 10}, TABELA)["cavas"] == 0      # poste de 7/8 m não conta


@pytest.mark.skipif(not NODE, reason="node indisponível")
def test_grade_e_casos_em_js_igual_ao_python():
    script = """
const C = require(process.argv[1]); const fs = require('fs');
const k = JSON.parse(fs.readFileSync(process.argv[2], 'utf8')), d = JSON.parse(fs.readFileSync(process.argv[3], 'utf8')), t = C.tabelaDe(d);
const xs = [0.05, 0.15, 0.25, 2.675, 6.45, -0.15, 1e-7, 123.45, 7.35, 0.65 * 10];
process.stdout.write(JSON.stringify({
  grade: k.grade.map(g => C.calcular(g.localidade, { SDEMU1004II: g.postes_estais }, t).cavas),
  casos: k.casos.map(c => C.calcular(c.localidade, c, t).cavas),
  mround: xs.map(x => C.mround(x, 0.1)), floor: xs.map(x => C.floor1(x)),
  loc: [C.resolverLocalidade({ bairro: 'agrestina' }, d), C.resolverLocalidade({ municipio: 'Betânia' }, d), C.resolverLocalidade({ bairro: 'x' }, d)],
  linhas: [C.linhaParaOutros(6, d), C.linhaParaOutros(0, d), C.ehLinhaCava('6-CAVA_ROCHA', d), C.ehLinhaCava('1-CAVA_ROCHA_X', d), C.ehLinhaCava('12600D, N3', d)],
  qtd: C.qtdDoResultado([{ codigo: 'SDEMU1004II', operacao: 'I', total: 3, desc_codigo: 'POSTE' }, { codigo: 'SDEMU1004II', operacao: 'R', total: 9 }, { codigo: 'SDEMU3002II', operacao: 'I', total: 2 }], d).qtd
}));
"""
    r = subprocess.run([NODE, "-e", script, os.path.join(RAIZ, "static", "cava_rocha.js"), os.path.join(ANEXOS, "casos_cava_icp.json"), os.path.join(ANEXOS, "cava_rocha_pernambuco.json")],
                       capture_output=True, text=True, check=True)
    js = json.loads(r.stdout)
    assert js["grade"] == [g["cavas_icp"] for g in CASOS["grade"]]
    assert js["casos"] == [c["SDEMU3007II_icp"] for c in CASOS["casos"]]
    xs = [0.05, 0.15, 0.25, 2.675, 6.45, -0.15, 1e-7, 123.45, 7.35, 0.65 * 10]
    assert js["mround"] == [cr.mround(x, 0.1) for x in xs] and js["floor"] == [cr.floor1(x) for x in xs]
    assert js["loc"] == [cr.resolver_localidade({"bairro": "agrestina"}, DADOS), cr.resolver_localidade({"municipio": "Betânia"}, DADOS), cr.resolver_localidade({"bairro": "x"}, DADOS)]
    assert js["linhas"] == [cr.linha_para_outros(6, DADOS), "", True, False, False] and js["linhas"][0] == "6-CAVA_ROCHA"
    assert js["qtd"] == {"SDEMU1004II": 3, "SDEMU3002II": 2}
    assert cr.qtd_do_resultado([{"codigo": "SDEMU1004II", "operacao": "I", "total": 3}, {"codigo": "SDEMU1004II", "operacao": "R", "total": 9}, {"codigo": "SDEMU3002II", "operacao": "I", "total": 2}], DADOS) == js["qtd"]


def test_localidade_e_linha():
    assert cr.resolver_localidade({"bairro": "Agrestina", "municipio": "OUTRA"}, DADOS)["indice"] == 0.65
    assert cr.resolver_localidade({}, DADOS)["localidade"] == ""
    assert cr.linha_para_outros(6, DADOS) == "6-CAVA_ROCHA" and cr.linha_para_outros(0, DADOS) == ""
    assert cr.eh_linha_cava("6-CAVA_ROCHA", DADOS) and cr.eh_linha_cava("CAVA_ROCHA", DADOS) and not cr.eh_linha_cava("1-TRANSP(faixa=8T)", DADOS)


def test_localidades_batem_com_os_bairros_do_transporte():
    transp = _json(os.path.join(RAIZ, "data", "pacotes", "081", "transporte.json"))
    bairros = {cr._norm(l.get("bairro")) for l in transp["municipio_utd"]}
    assert len(DADOS["localidades"]) == 121 and all(cr._norm(l["localidade"]) in bairros for l in DADOS["localidades"])


def test_parte_invalida_e_recusada():
    assert cr.validar(DADOS) == []
    assert cr.validar({"formato": "x"}) and cr.validar([]) and cr.validar(dict(DADOS, localidades=[{"localidade": "A"}]))


def test_ativo_na_base():
    rows = [{"ativo": "CAVA_ROCHA", "projeto": "PERNAMBUCO"}, {"ativo": "N3", "projeto": ""}]
    assert cr.ativo_na_base(rows, "CAVA_ROCHA", "PERNAMBUCO") and not cr.ativo_na_base(rows, "CAVA_ROCHA", "PARAIBA")
    assert cr.ativo_na_base([{"ativo": "CAVA_ROCHA", "projeto": ""}], "CAVA_ROCHA", "PARAIBA")


@pytest.fixture
def banco(tmp_path, monkeypatch):
    monkeypatch.setattr(database, "DB_PATH", str(tmp_path / "t.db"))
    monkeypatch.setattr(config, "DB_PATH", str(tmp_path / "t.db"))
    database.init_db()
    conn = database.get_connection()
    conn.execute("INSERT OR REPLACE INTO projetos (nome, codigo) VALUES ('PERNAMBUCO', '081')")
    conn.execute("INSERT OR REPLACE INTO projetos (nome, codigo) VALUES ('TESTE', '900')")
    conn.commit()
    conn.close()


def test_parte_no_pacote_081_exportar_e_importar(banco):
    assert "cava_rocha" in pp.PARTES and "transporte" in pp.PARTES
    assert cr.dados_do_projeto("PERNAMBUCO")["ativo"] == "CAVA_ROCHA" and len(cr.dados_do_projeto("081")["localidades"]) == 121
    assert cr.dados_do_projeto("TESTE") == {} and cr.dados_do_projeto("") == {}
    pacote = pp.ler_pacote(pp.exportar("081", ["cava_rocha"]))
    assert set(pacote["partes"]) == {"cava_rocha"} and pacote["partes"]["cava_rocha"] == DADOS
    previa = pp.previa("900", pacote)["partes"]
    assert previa[0]["situacao"] == "novo" and previa[0]["avisos"] == []
    pp.aplicar("900", pacote)
    assert cr.dados_do_projeto("TESTE") == DADOS


def test_rota_so_devolve_para_o_projeto_que_tem_a_parte(banco):
    import app as appmod
    cli = TestClient(appmod.app)
    cli.headers["Authorization"] = "Bearer " + create_jwt_token("u", "u@example.test", "operador")
    d = cli.get("/api/orcamento/cava_rocha", params={"projeto": "PERNAMBUCO"}).json()
    assert d["ativo"] == "CAVA_ROCHA" and d["servico"] == "SDEMU3007II" and isinstance(d["ativo_na_base"], bool)
    assert cli.get("/api/orcamento/cava_rocha", params={"projeto": "TESTE"}).json() == {}


def test_botao_ao_lado_do_mapa_e_sem_nome_de_projeto_no_codigo():
    html = open(os.path.join(RAIZ, "static", "resumo.html"), encoding="utf-8").read()
    i = html.index(">MAPA</button>")
    assert 'id="btn-cava-rocha"' in html[i:i + 400] and 'data-modal-id="cava"' in html[i:i + 400]
    assert "/static/cava_rocha.js" in html
    for arq in ("services/cava_rocha.py", "static/cava_rocha.js"):
        t = open(os.path.join(RAIZ, arq), encoding="utf-8").read().upper()
        assert "PERNAMBUCO" not in t and not re.search(r"\b081\b", t), arq
