"""TASK-031 fase C — API de controle do modo autônomo (só admin) e modo trabalhador do executável."""
import os
import sys
import time

import pytest
from fastapi.testclient import TestClient

import app as appmod
import database
from middleware.auth_middleware import create_jwt_token
from services.autonomo import config_autonomo as ca, pasta, pipeline as pl
from services.autonomo.leitor_js import LeitorJS
from tests.test_autonomo_pasta import dxf
from tests.test_autonomo_pipeline import RECEITA_REMOVE, contexto


@pytest.fixture(scope="module")
def leitor():
    lt = LeitorJS()
    yield lt
    lt.fechar()


@pytest.fixture
def api(tmp_path, monkeypatch, leitor):
    monkeypatch.setattr(database, "DB_PATH", str(tmp_path / "t.db"))
    database.init_db()
    conn = database.get_connection()
    conn.execute("INSERT OR IGNORE INTO projetos (nome, codigo) VALUES ('P1', 'P1')")
    conn.commit()
    conn.close()
    receitas = []
    monkeypatch.setattr(pl, "carregar_contexto", lambda proj, user: contexto(leitor, receitas))
    monkeypatch.setattr(pasta, "_vigia", pasta.Vigia())          # vigia isolada por teste
    c = TestClient(appmod.app)
    c.receitas, c.base = receitas, tmp_path / "auto"
    yield c
    pasta.obter_vigia().parar()


def cab(role="admin", uid="u1"):
    return {"Authorization": f"Bearer {create_jwt_token(uid, f'{role}@x.com', role)}"}


def configurar(api, **extra):
    r = api.put("/api/autonomo/config", json={"pasta_base": str(api.base), "user_id": "u1", "intervalo_s": 1, "estabilizacao_s": 0, **extra}, headers=cab())
    assert r.status_code == 200, r.text
    return r.json()


def test_so_admin_acessa(api):
    for metodo, url in [("get", "/api/autonomo/config"), ("put", "/api/autonomo/config"), ("post", "/api/autonomo/ligar"), ("post", "/api/autonomo/desligar"),
                        ("get", "/api/autonomo/status"), ("post", "/api/autonomo/varrer"), ("get", "/api/autonomo/execucoes"),
                        ("post", "/api/autonomo/execucoes/x/confirmar")]:
        assert getattr(api, metodo)(url).status_code == 401, url
        assert getattr(api, metodo)(url, headers=cab("operador")).status_code == 403, url


def test_config_ida_e_volta_com_validacao(api):
    inicial = api.get("/api/autonomo/config", headers=cab(uid="adm9")).json()
    assert inicial["config"]["ligado"] is False and inicial["usuario_atual"]["user_id"] == "adm9"
    r = configurar(api)
    assert r["config"]["user_id"] == "u1" and r["pastas"]["entrada"] == str(api.base / "entrada")
    assert api.get("/api/autonomo/config", headers=cab()).json()["config"]["intervalo_s"] == 1
    ruim = api.put("/api/autonomo/config", json={"intervalo_s": 0}, headers=cab())
    assert ruim.status_code == 400 and "intervalo" in ruim.json()["detail"]
    assert api.get("/api/autonomo/config", headers=cab()).json()["config"]["intervalo_s"] == 1     # a recusada não gravou


def test_ligar_exige_dono_cria_pastas_e_processa_sozinho(api):
    assert api.post("/api/autonomo/ligar", headers=cab()).status_code == 400          # sem usuário dono
    configurar(api)
    r = api.post("/api/autonomo/ligar", headers=cab())
    assert r.status_code == 200 and r.json()["status"]["ligado"] is True
    p = ca.pastas(ca.carregar())
    assert all(os.path.isdir(v) for v in p.values()) and ca.carregar()["ligado"] is True
    dxf(os.path.join(p["entrada"], "P1", "auto.dxf"), ["1-U4"])
    limite = time.time() + 15
    while time.time() < limite and not os.path.exists(os.path.join(p["processados"], "P1", "auto.dxf")):
        time.sleep(0.3)
    assert os.path.exists(os.path.join(p["processados"], "P1", "auto.dxf"))
    st = api.get("/api/autonomo/status", headers=cab()).json()
    assert st["configurado"] is True and st["status"]["processados"] == 1
    d = api.post("/api/autonomo/desligar", headers=cab())
    assert d.json()["status"]["ligado"] is False and ca.carregar()["ligado"] is False


def test_varrer_agora_e_historico(api):
    configurar(api)
    p = ca.garantir_pastas(ca.carregar())
    dxf(os.path.join(p["entrada"], "P1", "m.dxf"), ["1-U4"])
    assert api.post("/api/autonomo/varrer", headers=cab()).json()["resultados"] == []
    res = api.post("/api/autonomo/varrer", headers=cab()).json()["resultados"]
    assert res[0]["status"] == "ok"
    lista = api.get("/api/autonomo/execucoes", headers=cab()).json()["execucoes"]
    assert len(lista) == 1 and lista[0]["id"] == res[0]["execucao"] and "originais" not in lista[0]
    ex = api.get(f"/api/autonomo/execucoes/{lista[0]['id']}", headers=cab()).json()
    assert ex["status"] == "ok" and ex["pendencias"] == [] and "originais" not in ex and "diff" not in ex
    assert api.get("/api/autonomo/execucoes/naoexiste", headers=cab()).status_code == 404
    assert api.get("/api/autonomo/execucoes?status=erro", headers=cab()).json()["execucoes"] == []


def test_pendencias_sim_sim_para_todos_rejeitar_reverter_e_reprocessar(api):
    api.receitas.append(RECEITA_REMOVE)
    configurar(api)
    p = ca.garantir_pastas(ca.carregar())
    dxf(os.path.join(p["entrada"], "P1", "a.dxf"), ["1-U3 1-U4", "1-U3 1-CFU", "1-U4"])
    dxf(os.path.join(p["entrada"], "P1", "b.dxf"), ["1-U3 1-U4", "2-CFU"])
    api.post("/api/autonomo/varrer", headers=cab())
    res = {r["arquivo"]: r for r in api.post("/api/autonomo/varrer", headers=cab()).json()["resultados"]}
    assert {r["status"] for r in res.values()} == {"aguardando_confirmacao"}
    assert api.get("/api/autonomo/status", headers=cab()).json()["aguardando_confirmacao"] == 2
    a = api.get(f"/api/autonomo/execucoes/{res['a.dxf']['execucao']}", headers=cab()).json()
    assert [x["indice"] for x in a["pendencias"]] == [0, 1] and "Editar" in a["pendencias"][0]["descricao"]
    # "Sim" individual: continua aguardando a outra
    parcial = api.post(f"/api/autonomo/execucoes/{a['id']}/confirmar", json={"indices": [0]}, headers=cab()).json()
    assert parcial["status"] == "aguardando_confirmacao" and [x["indice"] for x in parcial["pendencias"]] == [1]
    # "Sim para todos" (indices nulo) conclui
    final = api.post(f"/api/autonomo/execucoes/{a['id']}/confirmar", json={}, headers=cab()).json()
    assert final["status"] in ("ok", "com_pendencias") and final["obra_id"] and final["pendencias"] == []
    # arquivo b: rejeitar tudo
    b = api.post(f"/api/autonomo/execucoes/{res['b.dxf']['execucao']}/rejeitar", json={}, headers=cab()).json()
    assert b["status"] == "com_pendencias"
    # reverter / reprocessar
    rev = api.post(f"/api/autonomo/execucoes/{a['id']}/reverter", headers=cab()).json()
    assert rev["status"] == "revertida"
    assert api.post(f"/api/autonomo/execucoes/{a['id']}/reverter", headers=cab()).status_code == 400
    novo = api.post(f"/api/autonomo/execucoes/{a['id']}/reprocessar", headers=cab()).json()
    assert novo["id"] != a["id"] and novo["status"] == "aguardando_confirmacao"
    # erros viram 400 com mensagem
    r = api.post(f"/api/autonomo/execucoes/{novo['id']}/confirmar", json={"indices": [99]}, headers=cab())
    assert r.status_code == 400 and "sem pendência" in r.json()["detail"]
    assert api.post("/api/autonomo/execucoes/naoexiste/confirmar", json={}, headers=cab()).status_code == 400


def test_modo_trabalhador_do_executavel_funciona_pelo_app_py(monkeypatch):
    """No .exe o trabalhador é `programa --trabalhador-leitor-js`; aqui o equivalente em dev: `python app.py --trabalhador-leitor-js`."""
    import services.autonomo.leitor_js as lj
    monkeypatch.setattr(lj, "_comando_trabalhador", lambda: [sys.executable, "app.py", "--trabalhador-leitor-js"])
    lt = LeitorJS()
    try:
        r = lt.processar_lote([{"pagina": 1, "texto": "AFASTADOR", "cor": "#ff0000", "layer": ""}],
                              [{"fase": 1, "ordem": 1, "modo": "DEFINIR", "texto_regex": "AFASTADOR", "ativo_template": "1-AF", "parar": True}], [], "extracao")
        assert r[0]["ativo"] == "1-AF"
    finally:
        lt.fechar()


def test_app_inicia_a_vigia_se_estava_ligada(tmp_path, monkeypatch):
    monkeypatch.setattr(database, "DB_PATH", str(tmp_path / "t.db"))
    database.init_db()
    vig = pasta.Vigia()
    monkeypatch.setattr(pasta, "_vigia", vig)
    ca.salvar({"pasta_base": str(tmp_path / "auto"), "user_id": "u1", "ligado": True})
    appmod._iniciar_modo_autonomo()
    try:
        assert vig.ativo()
    finally:
        appmod._parar_modo_autonomo()
    assert not vig.ativo()
    ca.salvar({"pasta_base": str(tmp_path / "auto"), "user_id": "u1", "ligado": False})
    appmod._iniciar_modo_autonomo()
    assert not vig.ativo()


def test_pagina_e_script_da_tela_de_controle(api):
    r = api.get("/autonomo")
    assert r.status_code == 200 and "Modo autônomo" in r.text and "/static/autonomo.js" in r.text
    assert "no-cache" in r.headers.get("cache-control", "").lower() or "no-store" in r.headers.get("cache-control", "").lower()
    js = api.get("/static/autonomo.js")
    assert js.status_code == 200 and "application/javascript" in js.headers["content-type"] and "/api/autonomo/status" in js.text
    assert "admin" in api.get("/admin").text and 'href="/autonomo"' in api.get("/admin").text


def test_a_tela_nao_injeta_dados_do_servidor_como_html():
    """Nomes de arquivo, mensagens e caminhos vêm de fora: a tela só usa textContent (nunca innerHTML/insertAdjacentHTML)."""
    js = open("static/autonomo.js", encoding="utf-8").read()
    assert "innerHTML" not in js and "insertAdjacentHTML" not in js and "outerHTML" not in js and "document.write" not in js
