"""TASK-103 — CONTRATO do plugin nanoCAD (plugin_nanocad/) com a aplicação: o `.obra.json` que o plugin grava ao lado do desenho
é aceito por `POST /api/obras/importar` sem perder nada que o cálculo usa (vínculo cabo<->estrutura incluído), e a análise de vínculo
da PRÓPRIA aplicação sobre esse arquivo dá os mesmos parâmetros que o plugin mostrou ("Ver ativos com parâmetros").

O exemplo é gerado pelo núcleo C# do plugin, sem CAD (plugin_nanocad/tests/Nucleo.Testes, teste Exemplo_versionado_esta_em_dia).
"""
import json
from pathlib import Path

import pytest
from fastapi.testclient import TestClient

import app as appmod
import database
from middleware.auth_middleware import create_jwt_token
from services import obras_arquivo as oa
from services.autonomo.montagem import prefixo_cabo
from services.autonomo.vinculacao import codigo_estrutura, codigos_estrutura
from services.vinculacao_regras import analisar, parametros_por_estrutura

RAIZ = Path(__file__).resolve().parent.parent
EXEMPLO = RAIZ / "plugin_nanocad" / "exemplos" / "Projeto_Rua_das_Flores.obra.json"
CFG_PE = json.loads((RAIZ / "data" / "vinculacao_pernambuco_seed.json").read_text(encoding="utf-8"))


@pytest.fixture
def envelope():
    return json.loads(EXEMPLO.read_text(encoding="utf-8"))


@pytest.fixture
def client(tmp_path, monkeypatch):
    monkeypatch.setattr(database, "DB_PATH", str(tmp_path / "t.db"))
    database.init_db()
    conn = database.get_connection()
    conn.execute("INSERT OR IGNORE INTO projetos (nome, codigo) VALUES ('PERNAMBUCO', '081')")
    conn.commit()
    conn.close()
    return TestClient(appmod.app)


def cab(uid):
    return {"Authorization": f"Bearer {create_jwt_token(uid, f'{uid}@x.com', 'operador')}", "Content-Type": "application/json"}


def test_arquivo_do_plugin_passa_na_validacao_sem_perder_campos(envelope):
    pronta = oa.validar_importacao(envelope, "081")
    assert pronta["nome"] == "Projeto_Rua_das_Flores" and pronta["projeto"] == "081" and pronta["tipo"] == "obra"
    for chave in ("cabos", "outros"):
        assert pronta["dados"][chave] == envelope["dados"][chave], f"{chave}: a importação descartou campos do plugin"


def test_importar_pela_api_e_reabrir_com_vinculos(client, envelope):
    r = client.post("/api/obras/importar?projeto=081", content=json.dumps(envelope), headers=cab("u1"))
    assert r.status_code == 200, r.text
    assert r.json()["cabos"] == 7 and r.json()["outros"] == 8
    obra = client.get(f"/api/obras/{r.json()['id']}?projeto=081", headers=cab("u1")).json()
    dados = json.loads(obra["dados_json"])
    assert dados["cabos"]["bodyId"] == "body-cabos"                              # mesma forma do "Salvar obra" da tela
    v1 = dados["cabos"]["data"][0]
    assert v1["ativo"] == "3xA35P(7,94)-13,8kV 45 m" and v1["vinculoEstruturas"] == [0, 1] and v1["_geoTipo"] == "MT"
    assert dados["outros"]["data"][1]["ativo"] == "9400D, N1, N4" and len(dados["outros"]["data"][1]["_poste"]) == 2
    # exportar de volta devolve as mesmas linhas
    env2 = client.get(f"/api/obras/{r.json()['id']}/exportar?projeto=081", headers=cab("u1")).json()
    assert env2["dados"]["cabos"] == envelope["dados"]["cabos"] and env2["dados"]["outros"] == envelope["dados"]["outros"]


def test_recusa_se_o_projeto_selecionado_for_outro(client, envelope):
    r = client.post("/api/obras/importar?projeto=229", content=json.dumps(envelope), headers=cab("u1"))
    assert r.status_code == 400 and "081" in r.json()["detail"]["erros"][0]


def _normalizar(params):
    """{NOME: valor | [valores]} → {NOME: [valores não vazios]} (forma comum do plugin e da aplicação)."""
    saida = {}
    for nome, v in (params or {}).items():
        lista = [x for x in (v if isinstance(v, list) else [v]) if x]
        if lista:
            saida[nome] = lista
    return saida


def test_parametros_do_vinculo_iguais_aos_da_aplicacao(envelope):
    """A aplicação, analisando o vínculo do arquivo, chega aos MESMOS parâmetros que o plugin calculou (seção cad, só informativa)."""
    d = envelope["dados"]
    app_params = parametros_por_estrutura(d["cabos"], d["outros"], CFG_PE, prefixo_cabo, codigo_estrutura, codigos_estrutura)
    plugin = {}
    linha = 0
    for p in envelope["cad"]["postes"]:
        if not p.get("ativoLido"):
            continue
        for e in p["estruturas"]:
            par = _normalizar(e["parametros"])
            if par:
                plugin.setdefault(linha, {}).setdefault(e["nome"], par)
        linha += 1
    app_norm = {ln: {cod: _normalizar(ps) for cod, ps in porcod.items()} for ln, porcod in app_params.items()}
    assert app_norm == plugin
    assert app_norm[2]["N3"] == {"TIPOCABO": ["P"], "CABO": ["3XA35P_7.94_-13.8KV"]}


def test_estado_do_vinculo_na_aplicacao_bate_com_a_conferencia_do_plugin(envelope):
    d = envelope["dados"]
    a = analisar(d["cabos"], d["outros"], CFG_PE, prefixo_cabo, codigo_estrutura, codigos_estrutura)
    vaos = [v for v in envelope["cad"]["vaos"] if v.get("ativoLido")]
    for ic, est in a["cabos"].items():
        v = vaos[ic]
        completo = est["estado"] == "gr"
        assert completo == (v["estado"] == "OK"), f"{v['numero']}: aplicação={est['estado']} plugin={v['estado']}"
