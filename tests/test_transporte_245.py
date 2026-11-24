"""TASK-245 — transporte de material (TRANSP / TRANSP_KM) igual à aba TRANSPORT da ICP: 151 casos conferidos no LibreOffice, Python e JS
com a mesma saída, parte `transporte` do pacote (só o projeto que a tem mostra o modal) e o texto das linhas lançadas em Outros."""
import json
import os
import shutil
import subprocess

import pytest
from fastapi.testclient import TestClient

import config
import database
from middleware.auth_middleware import create_jwt_token
from services import pacote_projeto as pp
from services import transporte as tr

RAIZ = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
ANEXOS = os.path.join(RAIZ, ".ai", "tasks", "anexos", "TASK-245")
NODE = shutil.which("node")


def _json(nome):
    with open(os.path.join(ANEXOS, nome), encoding="utf-8") as f:
        return json.load(f)


CASOS = _json("casos_icp.json")["casos"]
DADOS = _json("transporte_pernambuco.json")
PESOS = [(p["codigo"], p["descricao"], p["kg"]) for p in DADOS["pesos"]]


def _km(c):
    return c["utd_km"] if c["municipio_informado"] else None


def test_151_casos_da_icp_em_python():
    assert len(CASOS) == 151
    for c in CASOS:
        r = tr.calcular(c["qtd_odi"], c["qtd_extra"], _km(c), PESOS)
        assert [list(s) for s in r["servicos"]] == c["saida_icp"], c
        assert r["toneladas"] == pytest.approx(c["toneladas"])


@pytest.mark.skipif(not NODE, reason="node indisponível")
def test_151_casos_da_icp_em_js_igual_ao_python():
    script = """
const T = require(process.argv[1]); const fs = require('fs');
const casos = JSON.parse(fs.readFileSync(process.argv[2], 'utf8')).casos, dados = JSON.parse(fs.readFileSync(process.argv[3], 'utf8'));
const out = casos.map(c => T.calcular(c.qtd_odi, c.qtd_extra, c.municipio_informado ? c.utd_km : null, T.pesosDe(dados)).servicos);
const arred = [0.185, 18.485, 1.005, 2.675, -0.185, 1537.185, 0.125, 1e-7, 123456789.125].map(x => T.arred(x, 2));
process.stdout.write(JSON.stringify({ out, arred, utd: T.resolverUtd({ bairro: '7 RO' }, dados), linhas: T.linhasParaOutros([['SDETU5007SI', 12.6], ['SDETU5008SI', 1537.19], ['#VALUE!', '#VALUE!']], dados) }));
"""
    r = subprocess.run([NODE, "-e", script, os.path.join(RAIZ, "static", "transporte.js"), os.path.join(ANEXOS, "casos_icp.json"), os.path.join(ANEXOS, "transporte_pernambuco.json")],
                       capture_output=True, text=True, check=True)
    js = json.loads(r.stdout)
    assert js["out"] == [c["saida_icp"] for c in CASOS]
    assert js["arred"] == [tr.arred(x, 2) for x in [0.185, 18.485, 1.005, 2.675, -0.185, 1537.185, 0.125, 1e-7, 123456789.125]]
    assert js["utd"] == tr.resolver_utd({"bairro": "7 RO"}, DADOS)
    assert js["linhas"] == tr.linhas_para_outros([("SDETU5007SI", 12.6), ("SDETU5008SI", 1537.19), ("#VALUE!", "#VALUE!")], DADOS)


def test_arredondamento_como_o_excel():
    assert tr.arred(0.185) == 0.19 and tr.arred(18.485) == 18.49 and tr.arred(-0.185) == -0.19


def test_exemplos_da_task_viram_as_linhas_de_outros():
    def linhas(t, km):
        kg = DADOS["pesos"][0]["kg"]
        r = tr.calcular({DADOS["pesos"][0]["codigo"]: t * 1000 / kg}, {}, km, PESOS)
        return [x["texto"] for x in tr.linhas_para_outros(r["servicos"], DADOS)]
    assert linhas(0.3, 42) == ["0.3-TRANSP(faixa=0.6T)"]
    assert linhas(0.3, 52) == ["0.3-TRANSP(faixa=0.6T)", "0.6-TRANSP_KM(faixa=0.6T)"]
    assert linhas(12, 102) == ["12-TRANSP(faixa=14T)", "624-TRANSP_KM(faixa=14T)"]
    assert linhas(30, 102) == ["5-TRANSP(faixa=8T)", "260-TRANSP_KM(faixa=8T)", "25-TRANSP(faixa=25T)", "1300-TRANSP_KM(faixa=25T)"]
    assert [x["texto"] for x in tr.linhas_para_outros([("SDETU5007SI", 12.6), ("SDETU5008SI", 1537.19)], DADOS)] == ["12.6-TRANSP(faixa=14T)", "1537.19-TRANSP_KM(faixa=14T)"]
    r = tr.linhas_para_outros(tr.calcular({DADOS["pesos"][0]["codigo"]: 50.0001 * 1000 / DADOS["pesos"][0]["kg"]}, {}, 52, PESOS)["servicos"], DADOS)
    assert [x["texto"] for x in r if not x["erro"]] == ["50-TRANSP(faixa=25T)"] and sum(x["erro"] for x in r) == 2


def test_utd_da_obra_bairro_e_municipio():
    assert tr.resolver_utd({"utd": "Paulista"}, DADOS) == {"utd": "PAULISTA", "km": 102, "origem": "UTD da obra"}
    assert tr.resolver_utd({"municipio": "olinda"}, DADOS)["km"] == 42
    assert tr.resolver_utd({}, DADOS) == {"utd": "", "km": None, "origem": ""}
    assert tr.resolver_utd({"municipio": "CIDADE QUE NAO EXISTE"}, DADOS)["utd"] == ""


def test_linha_de_transporte_reconhecida():
    assert tr.eh_linha_transporte("37.6-TRANSP(faixa=25T)", DADOS) and tr.eh_linha_transporte("4024.6-TRANSP_KM(faixa=25T)", DADOS)
    assert not tr.eh_linha_transporte("1-TRANSPORTE", DADOS) and not tr.eh_linha_transporte("12600D, N3", DADOS)


def test_parte_invalida_e_recusada():
    assert tr.validar(DADOS) == []
    assert tr.validar({"formato": "x"}) and tr.validar([]) and tr.validar(dict(DADOS, pesos=[{"codigo": "1"}]))


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
    assert "transporte" in pp.PARTES
    assert tr.dados_do_projeto("PERNAMBUCO")["formato"] == "transporte-material" and tr.dados_do_projeto("081")["km_por_utd"]["OLINDA"] == 42
    assert tr.dados_do_projeto("TESTE") == {} and tr.dados_do_projeto("") == {}
    pacote = pp.ler_pacote(pp.exportar("081", ["transporte"]))
    assert set(pacote["partes"]) == {"transporte"} and pacote["partes"]["transporte"] == DADOS
    previa = pp.previa("900", pacote)["partes"]
    assert previa[0]["situacao"] == "novo" and previa[0]["avisos"] == []
    pp.aplicar("900", pacote)
    assert tr.dados_do_projeto("TESTE") == DADOS


def test_rota_so_devolve_para_o_projeto_que_tem_a_parte(banco):
    import app as appmod
    cli = TestClient(appmod.app)
    cli.headers["Authorization"] = "Bearer " + create_jwt_token("u", "u@example.test", "operador")
    assert cli.get("/api/orcamento/transporte", params={"projeto": "PERNAMBUCO"}).json()["ativos"] == {"ate_50km": "TRANSP", "acima_50km": "TRANSP_KM"}
    assert cli.get("/api/orcamento/transporte", params={"projeto": "TESTE"}).json() == {}


def test_sem_nome_de_projeto_no_codigo():
    for arq in ("services/transporte.py", "static/transporte.js"):
        with open(os.path.join(RAIZ, arq), encoding="utf-8") as f:
            t = f.read().upper()
        assert "PERNAMBUCO" not in t and "081" not in t, arq
