"""TASK-117: configuração da barra de modais por projeto (API, funções puras em Node) e Limpar Tudo nos modais (estrutura da tela)."""
import json
import os
import re
import subprocess

import pytest
from fastapi.testclient import TestClient

import config
import database
from middleware.auth_middleware import create_jwt_token
from routers import modais_montaveis as mm
from tests import oraculo_tela as o
from tests.test_modal_montavel import RAMAIS, TITULOS

node = pytest.mark.skipif(not o.NODE, reason="node não instalado")
STATIC = config.STATIC_DIR


@pytest.fixture
def client(tmp_path, monkeypatch):
    monkeypatch.setattr(database, "DB_PATH", str(tmp_path / "t.db"))
    database.init_db()
    import app as appmod
    monkeypatch.setattr(mm, "get_supabase", lambda: None)
    return TestClient(appmod.app)


def hdr(role):
    return {"Authorization": f"Bearer {create_jwt_token('u', 'u@x.com', role)}"}


def _put_mm(client, projeto, modais):
    r = client.put("/api/modais-montaveis", json={"projeto_codigo": projeto, "dados": {"versao": 1, "modais": modais}}, headers=hdr("admin"))
    assert r.status_code == 200, r.text


# ── API ──
def test_barra_vazia_gravar_ler_isolado_por_projeto_e_so_admin_grava(client):
    assert client.get("/api/barra-modais?projeto_codigo=PE", headers=hdr("operador")).json()["dados"] == {"versao": 1, "ocultos": [], "ordem": []}
    dados = {"versao": 1, "ocultos": ["conexoes", "mm:ramais"], "ordem": ["ramais", "cabos"]}
    assert client.put("/api/barra-modais", json={"projeto_codigo": "PE", "dados": dados}, headers=hdr("operador")).status_code == 403
    assert client.put("/api/barra-modais", json={"projeto_codigo": "PE", "dados": dados}, headers=hdr("admin")).status_code == 200
    assert client.get("/api/barra-modais?projeto_codigo=PE", headers=hdr("operador")).json()["dados"] == dados
    assert client.get("/api/barra-modais?projeto_codigo=PB", headers=hdr("operador")).json()["dados"]["ocultos"] == []          # outro projeto não muda


def test_barra_valida_ids_e_duplicados(client):
    ruim = [{"ocultos": ["a b"]}, {"ocultos": "x"}, {"ordem": ["a", "a"]}, {"ocultos": ["x" * 61]}, {"ocultos": [str(i) for i in range(101)]}]
    for d in ruim:
        r = client.put("/api/barra-modais", json={"projeto_codigo": "PE", "dados": d}, headers=hdr("admin"))
        assert r.status_code == 400, d


def test_copiar_modais_montaveis_entre_projetos(client):
    _put_mm(client, "PB", [RAMAIS, TITULOS])
    _put_mm(client, "PE", [dict(RAMAIS, nome="Ramais do PE")])
    client.put("/api/barra-modais", json={"projeto_codigo": "PB", "dados": {"ocultos": ["mm:titulos", "conexoes"], "ordem": []}}, headers=hdr("admin"))
    corpo = {"origem": "PB", "destino": "PE", "ids": ["ramais", "titulos", "naoexiste"], "copiar_ocultos": True}
    assert client.post("/api/barra-modais/copiar", json=corpo, headers=hdr("operador")).status_code == 403
    r = client.post("/api/barra-modais/copiar", json=corpo, headers=hdr("admin"))
    assert r.status_code == 200, r.text
    assert {c["de"]: c["para"] for c in r.json()["copiados"]} == {"ramais": "ramais_2", "titulos": "titulos"}          # id repetido ganha sufixo
    ids = [m["id"] for m in client.get("/api/modais-montaveis?projeto_codigo=PE", headers=hdr("operador")).json()["dados"]["modais"]]
    assert ids == ["ramais", "ramais_2", "titulos"]
    cfg = client.get("/api/barra-modais?projeto_codigo=PE", headers=hdr("operador")).json()["dados"]
    assert cfg["ocultos"] == ["mm:titulos"]                                                                                  # só a ocultação do copiado; "conexoes" é do sistema
    # a origem não muda
    assert len(client.get("/api/modais-montaveis?projeto_codigo=PB", headers=hdr("operador")).json()["dados"]["modais"]) == 2


def test_copiar_recusa_mesmo_projeto_e_ids_inexistentes(client):
    _put_mm(client, "PB", [RAMAIS])
    assert client.post("/api/barra-modais/copiar", json={"origem": "PB", "destino": "PB", "ids": ["ramais"]}, headers=hdr("admin")).status_code == 400
    assert client.post("/api/barra-modais/copiar", json={"origem": "PB", "destino": "PE", "ids": ["x"]}, headers=hdr("admin")).status_code == 400


# ── funções puras (Node) ──
@node
def test_ordem_efetiva_e_mover():
    codigo = """const B=require(%s);
const padrao=['a','b','c','mm:x'];
console.log(JSON.stringify([B.ordemEfetiva([],padrao), B.ordemEfetiva(['c','zzz','a','a'],padrao), B.ordemEfetiva(null,padrao),
 B.mover(['a','b','c'],'b',-1), B.mover(['a','b','c'],'a',-1), B.mover(['a','b','c'],'c',1), B.mover(['a','b','c'],'q',1), B.SISTEMA.map(s=>s.id)]));""" % json.dumps(os.path.join(STATIC, "barra_modais.js"))
    r = subprocess.run([o.NODE, "-e", codigo], capture_output=True, text=True, timeout=30)
    assert r.returncode == 0, r.stderr
    e = json.loads(r.stdout)
    assert e[0] == ["a", "b", "c", "mm:x"] and e[1] == ["c", "a", "b", "mm:x"] and e[2] == ["a", "b", "c", "mm:x"]
    assert e[3] == ["b", "a", "c"] and e[4] == ["a", "b", "c"] and e[5] == ["a", "b", "c"] and e[6] == ["a", "b", "c"]
    assert e[7] == ["resumo-rede", "ajustar", "validar", "contexto", "regras", "postes", "cabos", "conexoes", "ramais", "vincular"]


# ── tela ──
def test_barra_do_html_tem_um_id_para_cada_modal_do_sistema():
    html = open(os.path.join(STATIC, "resumo.html"), encoding="utf-8").read()
    barra = html[html.index('id="barra-modais-itens"'):html.index('id="barra-modais-montaveis"')]
    ids = set(re.findall(r'data-modal-id="([^"]+)"', barra))
    js = open(os.path.join(STATIC, "barra_modais.js"), encoding="utf-8").read()
    sistema = set(re.findall(r"\{ id: '([a-z-]+)', nome:", js))
    assert ids == sistema and len(sistema) == 10
    assert 'id="cfg-modais-corpo"' in html and "/static/barra_modais.js" in html and "BarraModais.recarregar()" in html


def test_limpar_tudo_chama_os_modais_e_desfaz():
    js = open(os.path.join(STATIC, "resumo.js"), encoding="utf-8").read()
    i = js.index("btnLimparTabelas.addEventListener")
    corpo = js[i:js.index("GERADOR DE CÓDIGOS - CABOS")]
    assert "limparModaisDaObra()" in corpo and "coletarModaisDaObra()" in corpo
    assert "snap.modais" in js and "restaurarModaisDaObra(snap.modais)" in js
    assert "JSON.stringify([a && a.cabos, a && a.outros])" in js                         # `modais` não entra na comparação de estados
    gm = open(os.path.join(STATIC, "gerador_modal.js"), encoding="utf-8").read()
    assert "ctl.limpar = function" in gm
    mmui = open(os.path.join(STATIC, "modal_montavel_ui.js"), encoding="utf-8").read()
    assert "'data-modal-id': 'mm:' + m.id" in mmui and "excluir" in mmui
