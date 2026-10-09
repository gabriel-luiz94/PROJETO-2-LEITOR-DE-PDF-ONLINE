"""TASK-060 — modelos com ATIVO VARIÁVEL: X(A,B,C), X(nome:A,B) / X(nome:), X(@GRUPO); escolha validada no servidor; avisos; compatibilidade."""
import json
import re
import subprocess

import pytest
from fastapi.testclient import TestClient

import app as appmod
import database
import routers.obras as rota
from middleware.auth_middleware import create_jwt_token
from services import modelos_obra as mo
from services.validacao_planilhas import validar_planilhas
from tests import oraculo_tela


def L(ativo, op="I"):
    return {"entidade": "X", "operacao": op, "ativo": ativo}


def chaves(v):
    return [(x["chave"], x["tipo"]) for x in v]


# ── detecção ────────────────────────────────────────────────────────────────
def test_quantidade_e_ativo_variaveis_no_mesmo_item():
    v = mo.detectar_variaveis([], [L("V(qtd)-X(U3,U4,U5)")])
    assert chaves(v) == [("qtd", "quantidade"), ("ativo:#1", "ativo")]
    assert v[1]["definicoes"] == ["U3,U4,U5"] and v[1]["nome"] is None


def test_numeracao_separada_e_campo_nomeado_compartilhado():
    v = mo.detectar_variaveis([], [L("V-X(A,B) V-X(C,D)"), L("V(q)-X(poste:DT11/300,DT9/150)"), L("2-X(poste:)"), L("*V(q)-X(POSTE:)")])
    assert chaves(v) == [("#1", "quantidade"), ("ativo:#1", "ativo"), ("#2", "quantidade"), ("ativo:#2", "ativo"),
                        ("q", "quantidade"), ("ativo:poste", "ativo")]
    poste = v[-1]
    assert len(poste["ocorrencias"]) == 3 and poste["definicoes"] == ["DT11/300,DT9/150"]            # caixa não distingue; referências sem lista


def test_so_o_formato_exato_e_variavel():
    for texto in ("1-X", "1-X9", "2-AX(A,B)", "1-X(A,B", "X(A,B)", "CAX(A,B)-U4", "3-U4 1-CFU", "2-XX(A,B)"):
        assert not mo.tem_variavel(texto, "outros"), texto
        assert mo.detectar_variaveis([], [L(texto)]) == [], texto
    for texto in ("2-X(U3,U4)", "V-X(A,B)", "*V(q)-X(A,B)", "1-U4 2-X(A,B)", "V-U4"):
        assert mo.tem_variavel(texto, "outros"), texto


def test_cabos_nao_aceitam_ativo_variavel():
    assert mo.detectar_variaveis([L("X(A,B) ABC 35 m")], []) == []


def test_modelo_antigo_so_com_v_continua_igual():
    v = mo.detectar_variaveis([L("CAA 2 ABC V(vão) m")], [L("V-U4"), L("*V(postes)-CFU")])
    assert chaves(v) == [("vão", "quantidade"), ("#1", "quantidade"), ("postes", "quantidade")]
    g = mo.gerar([L("CAA 2 ABC V(vão) m")], [L("V-U4"), L("*V(postes)-CFU")], [], {"vão": 85, "#1": 2, "postes": 3})
    assert g["cabos"][0]["ativo"] == "CAA 2 ABC 85 m" and [x["ativo"] for x in g["outros"]] == ["2-U4", "*3-CFU"]


# ── opções, erros e avisos ──────────────────────────────────────────────────
def analise(outros, cfg=None, grupos=None, universo=None):
    return mo.analisar([], [L(t) if isinstance(t, str) else t for t in outros], cfg, grupos, universo)


def test_opcoes_em_maiusculas_sem_repeticao_e_padrao_conferido():
    a = analise(["V-X(u3,U4,u3)"], cfg=[{"chave": "ativo:#1", "rotulo": "Qual?", "padrao": "u4"}])
    p = next(x for x in a["parametros"] if x["tipo"] == "ativo")
    assert a["erros"] == [] and p["opcoes"] == ["U3", "U4"] and p["padrao"] == "U4" and p["rotulo"] == "Qual?"
    ruim = analise(["V-X(U3,U4)"], cfg=[{"chave": "ativo:#1", "padrao": "U9"}])
    assert any("não está entre as opções" in e for e in ruim["erros"])


@pytest.mark.parametrize("texto,trecho", [
    ("V-X(U3)", "pelo menos 2"), ("V-X(U3,u3)", "pelo menos 2"), ("V-X()", "nenhuma lista"), ("V-X(U3,,U4)", "opção vazia"),
    ("V-X(U3,U-4)", "inválida"), ("V-X(U3,U4*)", "inválida"), ("V-X(U3,U4:)", "inválida"),
    ("V-X(" + ",".join(f"A{i}" for i in range(31)) + ")", "Opções demais"),
    ("V-X(poste:)", "nenhuma lista"), ("V-X(p:A,B) V-X(p:C,D)", "listas diferentes"),
])
def test_erros_de_opcoes(texto, trecho):
    a = analise([texto])
    assert any(trecho in e for e in a["erros"]), a["erros"]


def test_lista_igual_repetida_nao_e_conflito():
    assert analise(["V-X(p:A,B) V-X(p:b,a)"])["erros"] == []


def test_grupos_expandem_na_hora_inclusive_aninhados_e_com_curinga():
    grupos = {"POSTES": ["DT11/300", "DT9/150"], "TODOS": ["@POSTES", "U4"], "DTS": ["DT*"]}
    a = analise(["V-X(@POSTES)", "V-X(@TODOS,CFU)"], grupos=grupos, universo={"DT1", "DT2", "U4", "CFU"})
    ops = [p["opcoes"] for p in a["parametros"] if p["tipo"] == "ativo"]
    assert ops == [["DT11/300", "DT9/150"], ["DT11/300", "DT9/150", "U4", "CFU"]] and a["erros"] == []
    b = analise(["V-X(@DTS)"], grupos=grupos, universo={"DT1", "DT2", "U4"})
    assert [p["opcoes"] for p in b["parametros"] if p["tipo"] == "ativo"] == [["DT1", "DT2"]]
    c = analise(["V-X(@DTS,U4)"], grupos=grupos, universo=None)          # sem a base não dá para expandir o curinga
    assert any("base técnica" in av for av in c["avisos"]) and c["erros"] == []
    d = analise(["V-X(@NAO,U4)"], grupos=grupos)
    assert any("não existe" in e for e in d["erros"])
    e = analise(["V-X(@PEQ)"], grupos={"PEQ": ["A"]})
    assert any("pelo menos 2" in x for x in e["erros"])


def test_opcao_fora_da_base_so_avisa():
    a = analise(["V-X(U3,ZZZ9)"], universo={"U3", "U4"})
    assert a["erros"] == [] and any("ZZZ9" in av and "base técnica" in av for av in a["avisos"])
    assert analise(["V-X(U3,U4)"], universo={"U3", "U4"})["avisos"] == []
    assert analise(["V-X(U3,ZZZ9)"], universo=set())["avisos"] == []              # sem base não há o que conferir


def test_configuracao_aceita_padrao_texto_so_em_campo_de_ativo():
    ok = mo.validar_configuracao([{"chave": "ativo:#1", "rotulo": "A", "padrao": "u4"}, {"chave": "#1", "rotulo": "Q", "padrao": "3,5"}])
    assert ok[0]["padrao"] == "U4" and ok[1]["padrao"] == 3.5
    with pytest.raises(mo.ErroModelo):
        mo.validar_configuracao([{"chave": "ativo:#1", "padrao": "U 4"}])
    with pytest.raises(mo.ErroModelo):
        mo.validar_configuracao([{"chave": "#1", "padrao": "U4"}])               # quantidade continua só número


# ── geração ─────────────────────────────────────────────────────────────────
def test_gera_linha_com_quantidade_e_ativo_escolhido():
    g = mo.gerar([], [L("2-U4 V(q)-X(U3,U4,U5)"), L("*V-X(CFU,CFUA)"), L("2-X(U3,U4)")], [], {"q": "3", "#1": 4, "ativo:#1": "u4", "ativo:#2": "CFUA", "ativo:#3": "U3"})
    assert [x["ativo"] for x in g["outros"]] == ["2-U4 3-U4", "*4-CFUA", "2-U3"]
    assert g["outros"][0]["operacao"] == "I" and not any(mo.tem_variavel(x["ativo"], "outros") for x in g["outros"])


def test_campo_nomeado_vale_para_todas_as_ocorrencias():
    g = mo.gerar([], [L("V(q)-X(luz:LED50,LED100)"), L("1-X(luz:)")], [], {"q": 3, "ativo:luz": "LED100"})
    assert [x["ativo"] for x in g["outros"]] == ["3-LED100", "1-LED100"]


def test_escolha_fora_da_lista_vazia_ou_desconhecida_e_recusada():
    for valor in ("U9", "U3,U4", "'U3'", True):
        with pytest.raises(mo.ErroModelo):
            mo.gerar([], [L("V-X(U3,U4)")], [], {"#1": 1, "ativo:#1": valor})
    with pytest.raises(mo.ErroModelo) as e:
        mo.gerar([], [L("V-X(U3,U4)")], [], {"#1": 1})
    assert "Escolha o ativo" in e.value.mensagens[0]
    with pytest.raises(mo.ErroModelo):
        mo.gerar([], [L("V-X(U3,U4)")], [], {"#1": 1, "ativo:#1": "U3", "ativo:#9": "U3"})


def test_padrao_do_ativo_vale_quando_nada_e_escolhido():
    cfg = [{"chave": "ativo:#1", "rotulo": "Qual", "padrao": "U4"}]
    assert mo.gerar([], [L("V-X(U3,U4)")], cfg, {"#1": 2})["outros"][0]["ativo"] == "2-U4"
    assert mo.gerar([], [L("V-X(U3,U4)")], cfg, {"#1": 2, "ativo:#1": "U3"})["outros"][0]["ativo"] == "2-U3"       # a escolha vence o padrão


def test_modelo_com_erro_nao_gera():
    with pytest.raises(mo.ErroModelo):
        mo.gerar([], [L("V-X(U3)")], [], {"#1": 1})
    with pytest.raises(mo.ErroModelo):
        mo.gerar([], [L("V-X(@NAO,U3)")], [], {"#1": 1, "ativo:#1": "U3"}, grupos={})


def test_geracao_usa_o_grupo_de_hoje():
    modelo = [L("V-X(@POSTES)")]
    assert mo.gerar([], modelo, [], {"#1": 1, "ativo:#1": "B"}, grupos={"POSTES": ["A", "B"]})["outros"][0]["ativo"] == "1-B"
    with pytest.raises(mo.ErroModelo):                     # o grupo mudou: B saiu
        mo.gerar([], modelo, [], {"#1": 1, "ativo:#1": "B"}, grupos={"POSTES": ["A", "C"]})


def test_camada_1_acusa_x_sobrando():
    ach = validar_planilhas([], [L("V-X(A,B)"), L("2-X(A,B)"), L("1-X9"), L("2-U4")])
    assert [(a["linha_id"], a["regra_id"]) for a in ach] == [("OUTROS-0", "C1-VAR"), ("OUTROS-1", "C1-VAR")]


# ── regex do frontend espelha o backend (Regra 4) ───────────────────────────
@pytest.mark.skipif(not oraculo_tela.NODE, reason="node não instalado")
def test_regex_do_frontend_igual_ao_backend():
    fonte = open("static/resumo.js", encoding="utf-8").read()
    m = re.search(r"const RE_V_OUTROS = (/.*/);", fonte)
    assert m
    amostras = ["V-U4", "*V(q)-CFU", "2-X(U3,U4)", "V-X(a,b)", "*2-X(p:A,B)", "1-X(p:)", "1-U4 2-X(A,B)", "1-X", "1-X9", "CAV-U4", "V1-X", "2-AX(A,B)",
                "X(A,B)", "3-U4 1-CFU", "V(q)-X(", "1-X(A,B", "V-X()", ""]
    saida = subprocess.run([oraculo_tela.NODE, "-e", f"const re = {m.group(1)}; const a = JSON.parse(require('fs').readFileSync(0,'utf8')); "
                            f"console.log(JSON.stringify(a.map(t => re.test(t))));"], input=json.dumps(amostras), capture_output=True, text=True, timeout=60)
    assert json.loads(saida.stdout) == [mo.tem_variavel(t, "outros") for t in amostras]


# ── API ─────────────────────────────────────────────────────────────────────
@pytest.fixture
def client(tmp_path, monkeypatch):
    monkeypatch.setattr(database, "DB_PATH", str(tmp_path / "t.db"))
    database.init_db()
    return TestClient(appmod.app)


def cab(uid="a", role="operador"):
    return {"Authorization": f"Bearer {create_jwt_token(uid, f'{uid}@x.com', role)}"}


def dados(outros=("2-U4", "V(qtd)-X(poste:DT11/300,DT9/150)", "1-X(poste:)")):
    return json.dumps({"cabos": {"bodyId": "body-cabos", "data": []}, "outros": {"bodyId": "body-outros", "data": [L(a) for a in outros]}})


def salvar(client, uid="a", oid="M1", outros=None, params=None, projeto="229"):
    corpo = {"id": oid, "nome": "Pontos", "data": "d", "dados_json": dados(*([outros] if outros else [])), "projeto": projeto, "tipo": "modelo"}
    if params is not None:
        corpo["parametros"] = params
    return client.post("/api/obras", json=corpo, headers=cab(uid))


@pytest.fixture
def contexto(monkeypatch):
    estado = {"grupos": {"POSTES": ["DT11/300", "DT9/150"]}, "universo": {"U4", "DT11/300", "DT9/150"}}
    monkeypatch.setattr(rota, "_contexto_modelo", lambda user, projeto: (estado["grupos"], estado["universo"]))
    return estado


def test_salvar_e_usar_modelo_com_escolha_de_ativo(client, contexto):
    r = salvar(client, params=[{"chave": "ativo:poste", "rotulo": "Poste", "padrao": "dt9/150"}, {"chave": "qtd", "rotulo": "Quantos", "padrao": "2"}])
    assert r.status_code == 200 and "avisos" not in r.json()
    snap = json.loads(client.get("/api/obras/M1?projeto=229", headers=cab("b")).json()["dados_json"])
    assert {p["chave"]: p["padrao"] for p in snap["modelo"]["parametros"]} == {"qtd": 2.0, "ativo:poste": "DT9/150"}
    assert snap["outros"]["data"][1]["ativo"] == "V(qtd)-X(poste:DT11/300,DT9/150)"                # o modelo guarda o texto X(...) intacto
    info = client.get("/api/obras/M1/modelo?projeto=229", headers=cab("b")).json()
    p = next(x for x in info["parametros"] if x["tipo"] == "ativo")
    assert p["opcoes"] == ["DT11/300", "DT9/150"] and p["padrao"] == "DT9/150" and p["rotulo"] == "Poste" and info["erros"] == []
    g = client.post("/api/obras/M1/gerar?projeto=229", json={"valores": {"qtd": "3", "ativo:poste": "DT11/300"}}, headers=cab("b")).json()
    assert [x["ativo"] for x in g["outros"]] == ["2-U4", "3-DT11/300", "1-DT11/300"]
    # sem escolher: usa o padrão
    g2 = client.post("/api/obras/M1/gerar?projeto=229", json={"valores": {}}, headers=cab("b")).json()
    assert [x["ativo"] for x in g2["outros"]] == ["2-U4", "2-DT9/150", "1-DT9/150"]
    # o servidor nunca aceita valor fora da lista
    r = client.post("/api/obras/M1/gerar?projeto=229", json={"valores": {"qtd": "3", "ativo:poste": "OUTRO"}}, headers=cab("b"))
    assert r.status_code == 400 and "Escolha inválida" in r.json()["detail"]["erros"][0]
    assert json.loads(client.get("/api/obras/M1?projeto=229", headers=cab("a")).json()["dados_json"])["outros"]["data"][1]["ativo"].startswith("V(qtd)-X(")


def test_salvar_recusa_erros_e_so_avisa_de_opcao_fora_da_base(client, contexto):
    r = salvar(client, outros=("V-X(U4)",))
    assert r.status_code == 400 and any("pelo menos 2" in e for e in r.json()["detail"]["erros"])
    r = salvar(client, params=[{"chave": "ativo:poste", "padrao": "ZZ"}])
    assert r.status_code == 400 and any("não está entre as opções" in e for e in r.json()["detail"]["erros"])
    r = salvar(client, outros=("V-X(U4,ZZZ9)",))
    assert r.status_code == 200 and any("ZZZ9" in a for a in r.json()["avisos"])                  # decisão 4: avisa e salva
    assert client.get("/api/obras/M1?projeto=229", headers=cab("a")).status_code == 200


def test_detectar_devolve_opcoes_erros_e_avisos(client, contexto):
    body = {"cabos": [], "outros": [L("V(q)-X(@POSTES)"), L("1-X(U4,ZZZ9)")], "projeto": "229"}
    r = client.post("/api/obras/modelo/detectar", json=body, headers=cab()).json()
    assert [(p["chave"], p["tipo"]) for p in r["parametros"]] == [("q", "quantidade"), ("ativo:#1", "ativo"), ("ativo:#2", "ativo")]
    assert r["parametros"][1]["opcoes"] == ["DT11/300", "DT9/150"] and r["erros"] == [] and any("ZZZ9" in a for a in r["avisos"])
    r2 = client.post("/api/obras/modelo/detectar", json={"cabos": [], "outros": [L("1-X(U4)")]}, headers=cab()).json()
    assert r2["erros"]
    assert client.post("/api/obras/modelo/detectar", json={"cabos": [], "outros": [], "parametros": [{"chave": "#1", "padrao": "x"}]}, headers=cab()).status_code == 400


def test_grupo_alterado_depois_muda_as_opcoes_na_hora_do_uso(client, contexto):
    salvar(client, outros=("V-X(@POSTES)",))
    contexto["grupos"]["POSTES"] = ["DT11/300", "DT9/150", "DT12/400"]
    info = client.get("/api/obras/M1/modelo?projeto=229", headers=cab("b")).json()
    assert info["parametros"][1]["opcoes"] == ["DT11/300", "DT9/150", "DT12/400"]
    contexto["grupos"].pop("POSTES")                                   # grupo apagado: o modelo acusa em vez de gerar errado
    info = client.get("/api/obras/M1/modelo?projeto=229", headers=cab("b")).json()
    assert info["erros"] and client.post("/api/obras/M1/gerar?projeto=229", json={"valores": {"#1": 1, "ativo:#1": "X"}}, headers=cab("b")).status_code == 400


def test_exportar_e_importar_modelo_com_ativo_variavel(client, contexto):
    salvar(client, params=[{"chave": "ativo:poste", "rotulo": "Poste", "padrao": "DT9/150"}])
    arq = client.get("/api/obras/M1/exportar?projeto=229", headers=cab("b")).json()
    assert arq["tipo"] == "modelo" and {p["chave"]: p["padrao"] for p in arq["parametros"]}["ativo:poste"] == "DT9/150"
    r = client.post("/api/obras/importar?projeto=229", content=json.dumps(arq), headers={**cab("b"), "Content-Type": "application/json"})
    assert r.status_code == 200 and r.json()["tipo"] == "modelo"
    g = client.post(f"/api/obras/{r.json()['id']}/gerar?projeto=229", json={"valores": {"qtd": 1}}, headers=cab("b")).json()
    assert g["outros"][1]["ativo"] == "1-DT9/150"
    arq["parametros"] = [{"chave": "ativo:poste", "rotulo": "P", "padrao": "NAO_EXISTE"}]
    r = client.post("/api/obras/importar?projeto=229", content=json.dumps(arq), headers={**cab("b"), "Content-Type": "application/json"})
    assert r.status_code == 400


def test_contexto_real_le_grupos_e_base_sem_levantar(client):
    grupos, universo = rota._contexto_modelo({"user_id": "a"}, "229")
    assert isinstance(grupos, dict) and isinstance(universo, set)
    assert rota._contexto_modelo({"user_id": "a"}, "nao_existe")[0] is not None


def test_gerar_pela_api_expande_o_grupo_do_projeto(client, contexto):
    salvar(client, outros=("V-X(@POSTES)",))
    r = client.post("/api/obras/M1/gerar?projeto=229", json={"valores": {"#1": 2, "ativo:#1": "DT9/150"}}, headers=cab("b"))
    assert r.status_code == 200 and r.json()["outros"][0]["ativo"] == "2-DT9/150"
    assert client.post("/api/obras/M1/gerar?projeto=229", json={"valores": {"#1": 2, "ativo:#1": "U4"}}, headers=cab("b")).status_code == 400    # fora do grupo
