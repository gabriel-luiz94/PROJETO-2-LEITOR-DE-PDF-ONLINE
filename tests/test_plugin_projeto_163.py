"""TASK-163 — integração do plugin do nanoCAD com o PROJETO DE TRABALHO (.obra): o plugin grava `<desenho>.obra` (com a seção `cad`) e o
OrçaCore abre esse projeto; gravar no programa preserva a `cad`; nada vai para o banco."""
import json
from pathlib import Path

import pytest
from fastapi.testclient import TestClient

import app as appmod
import config
import database
from middleware.auth_middleware import create_jwt_token
from services import janela_desktop, programa_local, projeto_desktop, projeto_obra

RAIZ = Path(__file__).resolve().parent.parent
EXEMPLO = json.loads((RAIZ / "plugin_nanocad" / "exemplos" / "Projeto_Rua_das_Flores.obra.json").read_text(encoding="utf-8"))


def projeto_do_plugin(pasta: Path, numero="1234567") -> Path:
    """O que o plugin grava: o envelope do .obra.json (com `cad`) + `projeto_trabalho` {nome, numero_obra}."""
    env = dict(EXEMPLO)
    env["projeto_trabalho"] = {"nome": "Rua das Flores", "numero_obra": numero, "sap": {}}
    p = pasta / "Rua das Flores.obra"
    p.write_text(json.dumps(env, ensure_ascii=False), encoding="utf-8")
    return p


@pytest.fixture
def client(tmp_path, monkeypatch):
    monkeypatch.setattr(database, "DB_PATH", str(tmp_path / "t.db"))
    database.init_db()
    conn = database.get_connection()
    conn.execute("INSERT OR IGNORE INTO projetos (nome, codigo) VALUES ('PERNAMBUCO', '081')")
    conn.commit()
    conn.close()
    monkeypatch.setattr(config, "APP_MODE", "desktop")
    return TestClient(appmod.app)


def cab():
    return {"Authorization": f"Bearer {create_jwt_token('u1', 'u1@x.com', 'operador')}"}


def test_programa_abre_o_projeto_do_plugin_e_preserva_a_cad_ao_gravar(tmp_path):
    caminho = projeto_do_plugin(tmp_path)
    r = projeto_desktop.abrir(str(caminho), "081")
    assert r["ok"] and r["projeto"]["nome"] == "Rua das Flores" and r["projeto"]["numero_obra"] == "1234567"
    assert len(r["projeto"]["dados"]["cabos"]) == 7
    # o programa grava (Resumo/SAP/resultado): a obra do plugin continua no arquivo
    g = projeto_desktop.gravar(str(caminho), {"sap": {"bairro": "Centro"}, "dados": r["projeto"]["dados"]}, r["assinatura"], projeto_codigo="081")
    assert g["ok"], g
    gravado = json.loads(caminho.read_text(encoding="utf-8"))
    assert gravado["cad"] == EXEMPLO["cad"]
    assert gravado["projeto_trabalho"]["sap"]["bairro"] == "Centro" and gravado["projeto_trabalho"]["numero_obra"] == "1234567"


def test_montar_sem_cad_continua_igual():
    texto = projeto_obra.montar({"nome": "X", "dados": {"cabos": [], "outros": []}}, None, "081")
    assert "cad" not in json.loads(texto)


def test_plugin_entrega_o_caminho_e_a_janela_abre_o_projeto(client, tmp_path, monkeypatch):
    urls = []
    monkeypatch.setattr(janela_desktop, "mostrar", lambda url, porta=None: urls.append(url) or True)
    caminho = projeto_do_plugin(tmp_path)
    r = client.post("/api/plugin/receber", json={"chave": "C:/obras/rua.dwg", "caminho": str(caminho)})     # sem token: é o plugin
    assert r.status_code == 200, r.text
    assert urls == [r.json()["url"]] and r.json()["url"].startswith("/resumo?plugin_pendente=")
    d = client.get(f"/api/plugin/pendente/{r.json()['id']}", headers=cab()).json()
    assert d["caminho"] == str(caminho.resolve()) and d["projeto"] == {"codigo": "081", "nome": "PERNAMBUCO"}
    assert "envelope" not in d
    assert client.get("/api/obras?projeto=081", headers=cab()).json() == []                                 # nada no banco


@pytest.mark.parametrize("ruim", ["C:/nao/existe.obra", "projeto.obra.json", "x.txt"])
def test_caminho_invalido_e_recusado(client, ruim, monkeypatch):
    monkeypatch.setattr(janela_desktop, "mostrar", lambda url, porta=None: True)
    assert client.post("/api/plugin/receber", json={"chave": "a", "caminho": ruim}).status_code == 400


def test_modelo_ou_arquivo_ilegivel_nao_abre(tmp_path):
    p = tmp_path / "ruim.obra"
    p.write_text("{nao é json", encoding="utf-8")
    assert programa_local.validar_projeto(str(p))[2]
    modelo = dict(EXEMPLO, tipo="modelo", parametros=[])
    p.write_text(json.dumps(modelo), encoding="utf-8")
    assert programa_local.validar_projeto(str(p))[2]


def test_programa_fechado_abre_com_o_projeto(tmp_path):
    caminho = projeto_do_plugin(tmp_path)
    assert programa_local.ler_argumento_projeto(["--abrir-projeto", str(caminho)]) == str(caminho)
    assert programa_local.ler_argumento_projeto([f"--abrir-projeto={caminho}"]) == str(caminho)
    c, codigo, erro = programa_local.validar_projeto(str(caminho))
    assert not erro and codigo == "081"
    i = programa_local.guardar_projeto("", c, codigo)
    assert programa_local.pegar(i)["caminho"] == c and caminho.exists()          # o arquivo do projeto NÃO é apagado
    app_py = (RAIZ / "app.py").read_text(encoding="utf-8")
    assert "programa_local.ler_argumento_projeto(sys.argv[1:])" in app_py


def test_telas_ligadas():
    js = (RAIZ / "static" / "projeto.js").read_text(encoding="utf-8")
    assert "async function abrirDoPlugin(caminho, projetoCodigo)" in js and "abrirDoPlugin," in js
    resumo = (RAIZ / "static" / "resumo.js").read_text(encoding="utf-8")
    assert "GridRProjeto.abrirDoPlugin(d.caminho" in resumo and "/api/plugin/obra" not in resumo


def test_contrato_o_programa_abre_o_obra_gerado_pelo_plugin(tmp_path):
    """plugin_nanocad/exemplos/Projeto_Rua_das_Flores.obra é gerado pelo plugin (ProjetoArquivoTests). O OrçaCore abre, grava sem mexer
    na `cad` e sem mudar operação/ativo das linhas (o plugin compara isso em cad.assinaturaLinhas para saber se o programa editou)."""
    origem = RAIZ / "plugin_nanocad" / "exemplos" / "Projeto_Rua_das_Flores.obra"
    gerado = json.loads(origem.read_text(encoding="utf-8"))
    assert gerado["projeto_trabalho"]["numero_obra"] == "1234567" and gerado["cad"]["assinaturaLinhas"]
    caminho = tmp_path / "Projeto_Rua_das_Flores.obra"
    caminho.write_text(origem.read_text(encoding="utf-8"), encoding="utf-8")
    r = projeto_desktop.abrir(str(caminho), "081")
    assert r["ok"], r
    assert r["projeto"]["numero_obra"] == "1234567" and not r["projeto"]["legado"]
    g = projeto_desktop.gravar(str(caminho), {"dados": r["projeto"]["dados"], "sap": {"utd": "X"}}, r["assinatura"], projeto_codigo="081")
    assert g["ok"], g
    depois = json.loads(caminho.read_text(encoding="utf-8"))
    assert depois["cad"] == gerado["cad"]
    linhas = lambda env: [(t, (l.get("operacao") or "").strip(), (l.get("ativo") or "").strip()) for t in ("cabos", "outros") for l in env["dados"][t]]
    assert linhas(depois) == linhas(gerado)
