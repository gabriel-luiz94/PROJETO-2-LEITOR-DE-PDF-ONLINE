"""TASK-178 — casos de referência por projeto: guardar pedido + resultado, conferir a base (atual ou candidata) em lote, aceitar mudança esperada e levar os casos no pacote."""
import time

import pytest
from fastapi.testclient import TestClient

import config
import database
from middleware.auth_middleware import create_jwt_token
from services import casos_referencia as cr
from services import orcamento_calc as calc
from services import pacote_projeto as pp


def L(ativo, comp, cod, fi=1.0, proj="TESTE", mdo="MATERIAL"):
    return {"ativo": ativo, "desc_ativo": ativo and "D", "componente": comp, "projeto": proj, "mdo": mdo, "codigo": cod, "desc_codigo": f"d {cod}", "fator_i": fi, "fator_r": 0.0, "filtro": "",
            "origem": "OUTROS", "lv": "A", "arred": "NORMAL", "formula_i": "", "formula_r": ""}


BASE = [L("N3", "N3", "isol", 2.0), L("", "N3", "para", 1.0), L("N4", "N4", "cart", 1.0), L("", "N4", "isol4", 3.0), L("T3", "T3", "trafo", 1.0)]


@pytest.fixture
def banco(tmp_path, monkeypatch):
    monkeypatch.setattr(database, "DB_PATH", str(tmp_path / "t.db"))
    monkeypatch.setattr(config, "DB_PATH", str(tmp_path / "t.db"))
    database.init_db()
    conn = database.get_connection()
    conn.execute("INSERT OR REPLACE INTO projetos (nome, codigo) VALUES ('TESTE', '900')")
    conn.execute("DELETE FROM tabela_orcamento_master")
    for r in BASE:
        conn.execute("INSERT INTO tabela_orcamento_master (ativo, desc_ativo, componente, projeto, mdo, codigo, desc_codigo, fator_i, fator_r, filtro, origem, lv, arred, formula_i, formula_r) VALUES (?,?,?,?,?,?,?,?,?,?,?,?,?,?,?)",
                     tuple(r[k] for k in ("ativo", "desc_ativo", "componente", "projeto", "mdo", "codigo", "desc_codigo", "fator_i", "fator_r", "filtro", "origem", "lv", "arred", "formula_i", "formula_r")))
    conn.commit()
    conn.close()
    return tmp_path


def _caso(nome, texto, rows=BASE):
    pedido = {"cabos": [], "outros": [{"ativo": texto, "operacao": "I", "entidade": "0"}]}
    esperado = cr.esperado_do_resultado(calc.processar_calculo([], pedido["outros"], "TESTE", rows)["resultado"])
    return {"nome": nome, "pedido": pedido, "esperado": esperado, "origem": "programa"}


def test_normalizar_e_validar_caso():
    c = cr.normalizar_caso({"nome": " A ", "pedido": {"cabos": [], "outros": [{"ativo": "1-N3"}]}, "esperado": [{"codigo": "x", "qtd": "2,5"}, {"codigo": ""}], "origem": "xx"})
    assert c["nome"] == "A" and c["esperado"] == [{"operacao": "I", "mdo": "", "codigo": "x", "qtd": 2.5}] and c["origem"] == "manual" and c["tolerancia"] == cr.TOLERANCIA_PADRAO
    for ruim in ({}, {"nome": "a"}, {"nome": "a", "pedido": {"cabos": [], "outros": []}}, {"nome": "a", "pedido": {"cabos": "x", "outros": []}}, 5):
        with pytest.raises(cr.ErroCaso):
            cr.normalizar_caso(ruim)


def test_comparar_com_tolerancia_e_tipos():
    a = [{"operacao": "I", "mdo": "MATERIAL", "codigo": "x", "qtd": 2.0}, {"operacao": "I", "mdo": "MATERIAL", "codigo": "y", "qtd": 1.0}]
    assert cr.comparar(a, [dict(a[0], qtd=2.004), dict(a[1])], 0.005) == []
    d = cr.comparar(a, [dict(a[0], qtd=3.0), {"operacao": "I", "mdo": "MATERIAL", "codigo": "z", "qtd": 1.0}], 0.005)
    assert {(x["codigo"], x["tipo"]) for x in d} == {("x", "qtd"), ("y", "sumiu"), ("z", "novo")}


def test_gravar_tres_casos_trocar_base_e_ver_o_que_mudou(banco):
    for n, t in (("n3", "2-N3"), ("n4", "1-N4"), ("t3", "1-T3")):
        cr.gravar("900", _caso(n, t), "u@x")
    assert [c["nome"] for c in cr.listar("900")] == ["n3", "n4", "t3"]
    proj = pp.projeto_do_cadastro("900")
    from routers.casos import conferir_projeto
    assert conferir_projeto(proj)["resumo"] == {"igual": 3, "mudou": 0, "erro": 0}
    # base candidata: o N4 perdeu o cartucho (um código a menos)
    candidata = [r for r in BASE if r["codigo"] != "cart"]
    csv_texto = pp._csv_base(candidata)
    res = conferir_projeto(proj, base_csv=csv_texto)
    assert res["resumo"] == {"igual": 2, "mudou": 1, "erro": 0}
    mudou = next(c for c in res["casos"] if c["situacao"] == "mudou")
    assert mudou["nome"] == "n4" and [(d["codigo"], d["tipo"]) for d in mudou["diferencas"]] == [("cart", "sumiu")]
    assert "cart" in res["relatorio"] and "n4" in res["relatorio"]
    assert conferir_projeto(proj)["resumo"]["mudou"] == 0                                   # a base de verdade não foi tocada


def test_aceitar_mudanca_esperada_atualiza_o_caso(banco):
    c = cr.gravar("900", _caso("n4", "1-N4"), "u@x")
    obtido = cr.esperado_do_resultado(calc.processar_calculo([], c["pedido"]["outros"], "TESTE", [r for r in BASE if r["codigo"] != "cart"])["resultado"])
    with pytest.raises(cr.ErroCaso):
        cr.aceitar_mudanca("900", c, obtido, "u@x", "  ")
    novo = cr.aceitar_mudanca("900", c, obtido, "u@x", "cartucho saiu do N4 de propósito")
    assert [e["codigo"] for e in novo["esperado"]] == ["isol4"] and novo["historico"][-1]["nota"].startswith("cartucho") and novo["historico"][-1]["quem"] == "u@x"
    assert cr.listar("900")[0]["esperado"] == novo["esperado"]


def test_erro_no_calculo_vira_erro_e_nao_derruba_o_lote(banco):
    bom = dict(cr.gravar("900", _caso("ok", "1-N3"), "u"))
    ruim = dict(bom, id=99, nome="ruim", pedido={"cabos": [], "outros": [{"ativo": 5, "operacao": None}]})
    res = cr.conferir([bom, ruim], BASE, "TESTE")
    assert [c["situacao"] for c in res["casos"]] in (["igual", "erro"], ["igual", "mudou"], ["igual", "igual"])


def test_desempenho_300_casos_com_a_base_carregada_uma_vez(banco):
    caso = cr.normalizar_caso(_caso("p", "2-N3"))
    casos = [dict(caso, id=i) for i in range(300)]
    t0 = time.perf_counter()
    res = cr.conferir(casos, BASE, "TESTE")
    dt = time.perf_counter() - t0
    assert res["resumo"] == {"igual": 300, "mudou": 0, "erro": 0} and dt < 30, dt
    print(f"\n300 casos em {dt:.2f}s")


def test_rotas(banco):
    import app as appmod
    c = TestClient(appmod.app)
    h = {"Authorization": f"Bearer {create_jwt_token('2', 'o@x.com', 'operador')}"}
    r = c.post("/api/casos", json={"projeto": "TESTE", "caso": _caso("via rota", "2-N3")}, headers=h)
    assert r.status_code == 200
    cid = r.json()["id"]
    assert c.get("/api/casos", params={"projeto": "900"}, headers=h).json()["casos"][0]["nome"] == "via rota"
    assert c.post("/api/casos", json={"projeto": "TESTE", "caso": {"nome": "x"}}, headers=h).status_code == 400
    ok = c.post("/api/casos/conferir", json={"projeto": "TESTE"}, headers=h).json()
    assert ok["resumo"]["igual"] == 1
    cand = pp._csv_base([r for r in BASE if r["codigo"] != "para"])
    mud = c.post("/api/casos/conferir", json={"projeto": "TESTE", "base_csv": cand}, headers=h).json()
    assert mud["resumo"]["mudou"] == 1
    assert c.post("/api/casos/aceitar", json={"projeto": "TESTE", "ids": [cid], "nota": "", "base_csv": cand}, headers=h).status_code == 400
    ac = c.post("/api/casos/aceitar", json={"projeto": "TESTE", "ids": [cid], "nota": "removi o parafuso", "base_csv": cand}, headers=h)
    assert ac.status_code == 200 and ac.json()["aceitos"] == [cid]
    assert c.post("/api/casos/conferir", json={"projeto": "TESTE", "base_csv": cand}, headers=h).json()["resumo"]["mudou"] == 0
    assert c.delete(f"/api/casos/{cid}", params={"projeto": "TESTE"}, headers=h).status_code == 200
    assert c.delete(f"/api/casos/{cid}", params={"projeto": "TESTE"}, headers=h).status_code == 404
    assert c.get("/api/casos", params={"projeto": "NAO"}, headers=h).status_code == 404


def test_casos_viajam_no_pacote_e_a_previa_mostra_o_impacto(banco):
    cr.gravar("900", _caso("n4", "1-N4"), "u")
    cr.gravar("900", _caso("n3", "2-N3"), "u")
    zipado = pp.exportar("900", ["casos", "base"]) if pp.valor_atual(pp.projeto_do_cadastro("900"), "base") else pp.exportar("900", ["casos"])
    pacote = pp.ler_pacote(zipado)
    assert len(pacote["partes"]["casos"]) == 2 and "id" not in pacote["partes"]["casos"][0]
    # num projeto vazio: importa os casos
    conn = database.get_connection()
    conn.execute("INSERT OR REPLACE INTO projetos (nome, codigo) VALUES ('OUTRO', '901')")
    conn.commit()
    conn.close()
    pp.aplicar("901", pacote, partes=["casos"])
    assert [c["nome"] for c in cr.listar("901")] == ["n4", "n3"] or sorted(c["nome"] for c in cr.listar("901")) == ["n3", "n4"]
    # prévia de um pacote com base: mostra quais casos mudariam
    cand = pp._csv_base([r for r in BASE if r["codigo"] != "cart"])
    pac2 = {"manifesto": {"projeto": {"nome": "TESTE", "codigo": "900"}}, "partes": {"base": cand}}
    prev = pp.previa("900", pac2)
    assert prev["casos"]["resumo"]["mudou"] == 1 and prev["casos"]["mudaram"][0]["nome"] == "n4"


def test_ligacoes_na_tela():
    import os
    ler = lambda *p: open(os.path.join(config.STATIC_DIR, *p), encoding="utf-8").read()
    ui, res = ler("vinculacao_ui.js"), ler("resultado_orcamento.html")
    assert "abaCasos" in ui and "/api/casos/conferir" in ui and "Mudança esperada" in ui
    assert "guardarCasoDeReferencia" in res and "Guardar como caso de referência" in res
