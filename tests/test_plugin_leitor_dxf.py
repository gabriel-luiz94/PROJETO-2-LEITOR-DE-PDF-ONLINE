"""Paridade do plugin com a extração e o motor do Leitor DXF."""
import io
import json
import shutil
import subprocess

import ezdxf
import pytest
from fastapi.testclient import TestClient

import app as appmod
import database
from middleware.auth_middleware import create_jwt_token
from routers import regras_leitor


def test_plugin_usa_a_mesma_extracao_e_as_mesmas_regras_do_leitor(tmp_path, monkeypatch):
    node = shutil.which("node")
    if not node:
        pytest.skip("Node indisponível para o oráculo do Leitor")
    monkeypatch.setattr(database, "DB_PATH", str(tmp_path / "paridade.db"))
    database.init_db()
    processo = [{"fase": 1, "ordem": 1, "texto_regex": r"^(.+)$", "ativo_template": "{1}"}, {
        "fase": 1, "ordem": 10, "texto_regex": r"^(CE2|A35P|B093116)$",
        "ativo_template": "{1} {vizinhanca} m",
        "vizinhanca": {"regex": r"^(\d+) m$", "geometrica": True, "raio": 12, "posicao": "ABAIXO",
                      "janela": 16, "direcao": "DEPOIS"},
    }, {"fase": 3, "ordem": 10, "operacao": "I"}]
    classes = [
        {"ordem": 10, "ativo_regex": r"^A35P \d+ m$", "entidade": "CABO", "parar": True},
        {"ordem": 20, "ativo_regex": r"^CE2$", "entidade": "ESTRUTURA", "parar": True},
    ]
    monkeypatch.setattr(regras_leitor, "_get_regras", lambda tabela, codigo: processo if tabela == "processamento" else classes)
    doc = ezdxf.new()
    for texto, x, y in [("CE2", 200, 100), ("B093116", 300, 99), ("8 m", 100, 98),
                        ("A35P", 0, 90), ("8 m", 100, 89), ("25 m", 0, 88)]:
        doc.modelspace().add_text(texto, dxfattribs={"insert": (x, y)})
    arquivo = io.StringIO()
    doc.write(arquivo)
    dxf = arquivo.getvalue().encode("utf-8")
    cliente = TestClient(appmod.app)
    caminho = "/api/plugin/ler-dxf?projeto=027"
    assert cliente.post(caminho, content=dxf).status_code == 401
    headers = {"Authorization": "Bearer " + create_jwt_token("teste", "teste@example.com", "operador")}
    resposta = cliente.post(caminho, content=dxf, headers=headers)
    assert resposta.status_code == 200, resposta.text
    extracao = cliente.post("/upload?projeto=027", files={"file": ("paridade.dxf", dxf)})
    assert extracao.status_code == 200
    itens = extracao.json()["data"]
    script = """
const L = require('./static/ler_projeto.js');
const M = require('./static/regras_leitor_engine.js');
const o = JSON.parse(require('fs').readFileSync(0, 'utf8'));
L.processar(o.itens, o.proc, o.cls, M).then(r => console.log(JSON.stringify([r.cabos,r.outros])));
"""
    programa = subprocess.run([node, "-e", script], check=True, text=True, capture_output=True,
                               input=json.dumps({"itens": itens, "proc": processo, "cls": classes}), timeout=15)
    cabos, outros = json.loads(programa.stdout)
    linhas = resposta.json()["linhas"]
    def triplas(rows):
        return [(l["entidade"], l["operacao"], l["ativo"]) for l in rows]
    assert triplas([l for l in linhas if l["entidade"] == "CABO"]) == triplas(cabos) == [("CABO", "I", "A35P 25 m")]
    assert triplas([l for l in linhas if l["entidade"] not in ("0", "CABO", "RAMAIS")]) == triplas(outros) == [
        ("ESTRUTURA", "I", "CE2")]
    assert resposta.json()["regras_leitor"]
    assert cliente.post(caminho, content=b"arquivo incorreto", headers=headers).status_code == 400


def test_regiao_aplica_as_regras_depois_de_excluir_vizinhos_externos():
    if not shutil.which("node"):
        pytest.skip("Node indisponível")
    database.init_db()
    doc = ezdxf.new()
    for texto, x, y in [("N3", 0, 0), ("N1", 20, 20), ("U4", 1, 0)]:
        doc.modelspace().add_text(texto, dxfattribs={"insert": (x, y)})
    arquivo = io.StringIO()
    doc.write(arquivo)
    headers = {"Authorization": "Bearer " + create_jwt_token("cad", "cad@example.com", "operador")}
    cliente = TestClient(appmod.app)
    rota = "/api/plugin/ler-dxf"
    params = {"projeto": "027", "regiao": json.dumps([[-1, -1], [1, -1], [1, 1], [-1, 1]])}
    resposta = cliente.post(rota, params=params, content=arquivo.getvalue().encode(), headers=headers)
    assert resposta.status_code == 200, resposta.text
    assert resposta.json()["regiao"] is True
    assert {r["texto"] for r in resposta.json()["linhas"]} == {"N3", "U4"}
    params["regiao"] = "[[1, 2]]"
    resposta = cliente.post(rota, params=params, content=arquivo.getvalue().encode(), headers=headers)
    assert resposta.status_code == 400
