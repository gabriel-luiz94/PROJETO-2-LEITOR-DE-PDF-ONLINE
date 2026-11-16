"""TASK-177 — pacote de projeto: exportar/importar num arquivo, pontos de partida pelos pacotes (sem `if "081"`), prévia sem gravar e isolamento entre projetos."""
import io
import json
import os
import re
import zipfile

import pytest
from fastapi.testclient import TestClient

import config
import database
from middleware.auth_middleware import create_jwt_token
from services import orcamento_calc as calc
from services import pacote_projeto as pp
from services import obra_params


@pytest.fixture
def banco(tmp_path, monkeypatch):
    monkeypatch.setattr(database, "DB_PATH", str(tmp_path / "t.db"))
    monkeypatch.setattr(config, "DB_PATH", str(tmp_path / "t.db"))
    database.init_db()
    conn = database.get_connection()
    conn.execute("INSERT OR REPLACE INTO projetos (nome, codigo) VALUES ('PERNAMBUCO', '081')")
    conn.execute("INSERT OR REPLACE INTO projetos (nome, codigo) VALUES ('TESTE', '900')")
    conn.execute("INSERT OR REPLACE INTO projetos (nome, codigo) VALUES ('OUTRO', '901')")
    conn.commit()
    conn.close()
    obra_params.limpar_cache()
    return tmp_path / "t.db"


def _base_pe(n=6):
    return [{"ativo": "N3", "desc_ativo": "ESTRUTURA N3", "componente": "N3", "projeto": "PERNAMBUCO", "mdo": "MATERIAL", "codigo": "1000", "desc_codigo": "ISOLADOR", "fator_i": 2.0, "fator_r": 0.0,
             "filtro": "", "origem": "OUTROS", "lv": "A", "arred": "NORMAL", "formula_i": "", "formula_r": ""},
            {"ativo": "", "desc_ativo": "", "componente": "N3", "projeto": "PERNAMBUCO", "mdo": "MATERIAL", "codigo": "1001", "desc_codigo": "PARAFUSO", "fator_i": 1.0, "fator_r": 0.0,
             "filtro": "", "origem": "OUTROS", "lv": "A", "arred": "NORMAL", "formula_i": "", "formula_r": ""}][:n]


def _gravar_base(linhas):
    conn = database.get_connection()
    conn.execute("DELETE FROM tabela_orcamento_master")
    for r in linhas:
        conn.execute("INSERT INTO tabela_orcamento_master (ativo, desc_ativo, componente, projeto, mdo, codigo, desc_codigo, fator_i, fator_r, filtro, origem, lv, arred, formula_i, formula_r) VALUES (?,?,?,?,?,?,?,?,?,?,?,?,?,?,?)",
                     (r["ativo"], r["desc_ativo"], r["componente"], r["projeto"], r["mdo"], r["codigo"], r["desc_codigo"], r["fator_i"], r["fator_r"], r["filtro"], r["origem"], r["lv"], r["arred"], r["formula_i"], r["formula_r"]))
    conn.commit()
    conn.close()


def _resultado(projeto_nome):
    conn = database.get_row_connection()
    rows = [dict(r) for r in conn.execute("SELECT * FROM tabela_orcamento_master WHERE UPPER(TRIM(projeto)) = ?", (projeto_nome,)).fetchall()]
    conn.close()
    r = calc.processar_calculo([], [{"ativo": "2-N3", "operacao": "I", "entidade": "0"}], projeto_nome, rows)
    return sorted((x["codigo"], x["total"]) for x in r["resultado"])


def test_pontos_de_partida_vem_dos_pacotes_embutidos(banco):
    ids = pp.pacotes_embutidos()
    assert {"081", "rondonia"} <= set(ids)
    assert pp.pacote_inicial_de("081")["id"] == "081" and pp.pacote_inicial_de("PERNAMBUCO")["id"] == "081" and pp.pacote_inicial_de("229")["id"] == "rondonia"
    assert pp.pacote_inicial_de("900") is None and pp.config_inicial("900") is None
    cfg = pp.config_inicial("081")
    assert len(cfg["cabos"]) >= 5 and cfg["dicionario_cabos"] and cfg["postes"] and cfg["exportacao"]
    from routers.vinculacao_config import obter_config_com_origem
    c, origem = obter_config_com_origem("081")
    assert origem == "pernambuco" and c["dicionario_cabos"]
    c, origem = obter_config_com_origem("900")
    assert origem in ("vazia", "migrada") and not c["cabos"]
    assert obra_params.tabela_regional("PERNAMBUCO")["regioes"] and obra_params.tabela_regional("pernambuco") == obra_params.tabela_regional("081") and obra_params.tabela_regional("TESTE") == {}


def test_pacote_inicial_do_cadastro_vale_mais_que_o_manifesto(banco):
    conn = database.get_connection()
    conn.execute("UPDATE projetos SET pacote_inicial = '081' WHERE codigo = '900'")
    conn.commit()
    conn.close()
    assert pp.pacote_inicial_de("900")["id"] == "081" and pp.config_inicial("TESTE")["cabos"]


def test_sem_casos_especiais_no_codigo():
    achados = []
    for pasta in ("services", "routers"):
        for raiz, _, nomes in os.walk(os.path.join(config.BASE_DIR, pasta)):
            for n in nomes:
                if not n.endswith(".py") or "autonomo" in raiz:
                    continue
                for i, linha in enumerate(open(os.path.join(raiz, n), encoding="utf-8"), 1):
                    cod = linha.split("#")[0]
                    if re.search(r'==\s*["\'](081|PERNAMBUCO)["\']|["\'](081|PERNAMBUCO)["\']\s*(?:==|:)|PROJETO_PERNAMBUCO|_seed_pernambuco\(|_atualizar_seed_pernambuco|_TABELAS\b', cod):
                        achados.append(f"{pasta}/{n}:{i}: {linha.strip()[:100]}")
    assert not achados, achados


def test_exportar_081_e_importar_num_projeto_vazio_da_o_mesmo_orcamento(banco):
    _gravar_base(_base_pe())
    zipado = pp.exportar("081")
    pacote = pp.ler_pacote(zipado)
    assert pacote["manifesto"]["formato"] == "pacote-projeto" and pacote["manifesto"]["versao_formato"] == 1 and set(pacote["partes"]) >= {"base", "vinculacao", "dicionario_cabos", "postes", "exportacao", "regional"}
    antes = _resultado("PERNAMBUCO")
    assert antes == [("1000", 4.0), ("1001", 2.0)]
    prev = pp.previa("900", pacote)
    assert {p["parte"]: p["situacao"] for p in prev["partes"]}["base"] == "novo" and all(p["situacao"] in ("novo", "substitui") for p in prev["partes"])
    assert _resultado("TESTE") == []                                                                  # a prévia não gravou nada
    r = pp.aplicar("900", pacote, user_id="u")
    assert "base" in r["gravadas"]
    assert _resultado("TESTE") == antes
    from routers.vinculacao_config import obter_config
    cfg = obter_config("900")
    assert cfg["dicionario_cabos"] and cfg["postes"] and cfg["cabos"] == obter_config("081")["cabos"]
    assert obra_params.tabela_regional("TESTE")["regioes"]
    assert {p["parte"]: p["situacao"] for p in pp.previa("900", pacote)["partes"]}["base"] == "igual"


def test_importar_num_projeto_nao_altera_outro(banco):
    _gravar_base(_base_pe())
    pacote = pp.ler_pacote(pp.exportar("081", ["base", "vinculacao"]))
    outro_antes = pp.valor_atual(pp.projeto_do_cadastro("901"), "vinculacao")
    pp.aplicar("900", pacote)
    assert pp.valor_atual(pp.projeto_do_cadastro("901"), "base") is None
    assert pp.valor_atual(pp.projeto_do_cadastro("901"), "vinculacao") == outro_antes
    assert len(pp.valor_atual(pp.projeto_do_cadastro("081"), "base").splitlines()) == 3                 # o projeto de origem continua com as linhas dele
    assert _resultado("PERNAMBUCO") == [("1000", 4.0), ("1001", 2.0)]


def test_parte_ausente_nao_mexe_e_escolha_de_partes(banco):
    _gravar_base(_base_pe())
    pacote = pp.ler_pacote(pp.exportar("081", ["base"]))
    assert list(pacote["partes"]) == ["base"]
    pp.aplicar("900", pacote)
    from routers.vinculacao_config import _config_salva
    assert _config_salva("900") is None                                                                 # só a base foi gravada
    pacote = pp.ler_pacote(pp.exportar("081", ["vinculacao", "postes"]))
    r = pp.aplicar("901", pacote, partes=["postes"])
    assert r["gravadas"] == ["postes"]


def test_cancelar_nao_grava_e_erro_na_validacao_nao_grava_nada(banco):
    _gravar_base(_base_pe())
    zipado = pp.exportar("081")
    pp.previa("900", pp.ler_pacote(zipado))                                                             # só olhar
    assert pp.valor_atual(pp.projeto_do_cadastro("900"), "base") is None
    # pacote com uma parte inválida: nada é gravado, nem as partes boas
    pac = pp.ler_pacote(zipado)
    pac["partes"]["postes"] = {"tipos": []}
    with pytest.raises(pp.ErroPacote):
        pp.aplicar("900", pac)
    assert pp.valor_atual(pp.projeto_do_cadastro("900"), "base") is None


def test_formato_versao_hash_e_arquivo_invalido(banco):
    _gravar_base(_base_pe())
    zipado = pp.exportar("081", ["postes"])

    def reescrever(alterar):
        entrada = zipfile.ZipFile(io.BytesIO(zipado))
        saida = io.BytesIO()
        with zipfile.ZipFile(saida, "w") as z:
            for n in entrada.namelist():
                z.writestr(n, alterar(n, entrada.read(n)))
        return saida.getvalue()

    novo = reescrever(lambda n, b: json.dumps({**json.loads(b), "versao_formato": 99}).encode() if n == "manifest.json" else b)
    with pytest.raises(pp.ErroPacote, match="mais novo"):
        pp.ler_pacote(novo)
    with pytest.raises(pp.ErroPacote, match="integridade"):
        pp.ler_pacote(reescrever(lambda n, b: b + b" " if n == "postes.json" else b))
    with pytest.raises(pp.ErroPacote, match="formato"):
        pp.ler_pacote(reescrever(lambda n, b: json.dumps({**json.loads(b), "formato": "x"}).encode() if n == "manifest.json" else b))
    with pytest.raises(pp.ErroPacote, match="zip"):
        pp.ler_pacote(b"nao e zip")
    with pytest.raises(pp.ErroPacote):
        pp.ler_pacote(reescrever(lambda n, b: b if n != "manifest.json" else b"{}"))
    with pytest.raises(pp.ErroPacote, match="nenhuma"):
        pp.exportar("901", ["base"])


def test_rotas(banco):
    import app as appmod
    c = TestClient(appmod.app)
    admin = {"Authorization": f"Bearer {create_jwt_token('1', 'a@x.com', 'admin')}"}
    operador = {"Authorization": f"Bearer {create_jwt_token('2', 'o@x.com', 'operador')}"}
    _gravar_base(_base_pe())
    assert c.get("/api/pacote/partes", headers=operador).status_code == 200
    r = c.get("/api/pacote/exportar", params={"projeto": "081", "partes": "base,postes"}, headers=operador)
    assert r.status_code == 200 and r.headers["content-type"] == "application/zip" and zipfile.ZipFile(io.BytesIO(r.content)).namelist()[0] == "manifest.json"
    arq = {"file": ("pacote.zip", r.content, "application/zip")}
    p = c.post("/api/pacote/previa", data={"projeto": "TESTE"}, files=arq, headers=operador)
    assert p.status_code == 200 and {x["parte"] for x in p.json()["partes"]} == {"base", "postes"}
    assert c.post("/api/pacote/importar", data={"projeto": "TESTE", "confirmar": "true"}, files=arq, headers=operador).status_code == 403       # só admin importa
    assert c.post("/api/pacote/importar", data={"projeto": "TESTE"}, files=arq, headers=admin).status_code == 400                                 # sem confirmar não grava
    assert pp.valor_atual(pp.projeto_do_cadastro("TESTE"), "base") is None
    ok = c.post("/api/pacote/importar", data={"projeto": "TESTE", "confirmar": "true"}, files=arq, headers=admin)
    assert ok.status_code == 200 and "base" in ok.json()["gravadas"]
    assert c.post("/api/pacote/previa", data={"projeto": "NAOEXISTE"}, files=arq, headers=operador).status_code == 400
    assert c.get("/api/pacote/exportar", params={"projeto": "NAOEXISTE"}, headers=operador).status_code == 400


def test_aba_pacote_na_tela():
    ui = open(os.path.join(config.STATIC_DIR, "vinculacao_ui.js"), encoding="utf-8").read()
    assert "abaPacote" in ui and "/api/pacote/previa" in ui and "Confirmar e importar" in ui
