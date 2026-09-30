"""tests/test_regras_v2.py — motor de regras v2 (TASK-016): linguagem nova e equivalência com o motor v1 congelado."""
import copy
import json
import random

import pytest
from fastapi.testclient import TestClient

import app as appmod
import database
from config import REGRAS_DOMINIO_SEED_PATH
from middleware.auth_middleware import create_jwt_token
from services.regras_dominio import (avaliar, converter_v1, descrever, explicar, validar_conjunto, validar_grupos,
                                     validar_regras)
from tests import referencia_regras_v1 as v1

with open(REGRAS_DOMINIO_SEED_PATH, encoding="utf-8") as f:
    SEED = json.load(f)
with open("tests/fixtures/regras_dominio_seed_v1.json", encoding="utf-8") as f:
    SEED_V1 = json.load(f)


def L(ativo, op="I"):
    return {"ativo": ativo, "operacao": op}


def regra(**kw):
    base = {"id": "T", "versao": 2, "escopo": "linha", "operacoes": ["I", "*I"], "ativa": True, "severidade": "aviso", "mensagem": "m"}
    return {**base, **kw}


def alertas(r, outros, cabos=(), grupos=None):
    outros = [L(o) if isinstance(o, str) else L(*o) for o in outros]
    return avaliar([r], list(cabos), outros, grupos or {})


# ═══════════════════════════════════════════════════════════════════════════
# EQUIVALÊNCIA v1 ↔ v2 (o motor v1 é a cópia congelada em tests/referencia_regras_v1.py)
# ═══════════════════════════════════════════════════════════════════════════
POSTES = ["DT11/300", "DT10/300", "CV10/300", "1-4DT10/300", "CV11/600", "1-DT11/300", "POSTE11", "DT11"]
ITENS = ["CFU", "CFUR", "CFA", "CL", "RCFU", "EF3H", "EF10K", "TR105", "TR125", "TR215", "TR330", "TR330A", "U1", "U3", "U3C",
         "N3", "N3IV", "N1", "T3", "T1", "TE", "R3", "R2", "SUPL", "PR15", "PR220", "RPR", "SI1", "SI2", "SI3", "SI4", "S2",
         "S44", "RA2", "P50", "IP", "U23"]


def cenario(rng):
    def linha():
        partes = [rng.choice(POSTES)] if rng.random() < 0.6 else []
        for _ in range(rng.randint(0, 6)):
            partes.append(f"{rng.choice([1, 1, 1, 2, 3, 0, '2,5'])}-{rng.choice(ITENS)}")
        return " ".join(partes) or "1-IP"
    outros = [{"ativo": linha(), "operacao": rng.choice(["I", "I", "*I", "R", "M"])} for _ in range(rng.randint(1, 6))]
    cabos = [{"ativo": rng.choice(["P50 ABC 2 m", "P50 A 3 m", "CAA 2 ABC 35 m", "P 50 ABC 6 m", "P50"]),
              "operacao": rng.choice(["I", "R", "*I"])} for _ in range(rng.randint(0, 3))]
    return cabos, outros


def chaves(achados):
    return {(a["linha_id"], a["regra_id"], a["severidade"], a["tabela"]) for a in achados}


def ligar(regras):
    regras = copy.deepcopy(regras)
    for r in regras:
        r["ativa"] = True
    return regras


def test_semente_v2_equivale_a_semente_v1_em_3000_cenarios():
    novas, antigas, rng = ligar(SEED["regras"]), ligar(SEED_V1), random.Random(2026)
    total = 0
    for _ in range(3000):
        cabos, outros = cenario(rng)
        a = chaves(v1.avaliar(antigas, cabos, outros))
        assert a == chaves(avaliar(novas, cabos, outros, SEED["grupos"])), (cabos, outros)
        total += len(a)
    assert total > 5000   # o teste só vale se houve muito alerta para comparar


def test_conversor_v1_para_v2_equivale_ao_motor_v1_para_qualquer_regra_v1_valida():
    """Regras v1 antigas (e variações editadas pelo admin) continuam com o mesmo resultado depois de convertidas."""
    variantes = ligar(SEED_V1)
    extra = copy.deepcopy(SEED_V1[3])                              # PR220 mono com dobra
    extra.update(id="X1", ativa=True)
    extra["parametros"].update(qtd_min=1, multiplicador=3, dobra_se=[{"regex": "^SI3$", "qtd_min": 1}])
    extra2 = copy.deepcopy(SEED_V1[7])                             # nao_isolado com min_outros=2 e acompanhantes diferentes
    extra2.update(id="X2", ativa=True)
    extra2["parametros"].update(min_outros=2, acompanhado_por_regex="^(U1|U3|N3|TR\\d.*)$")
    extra3 = copy.deepcopy(SEED_V1[6])                             # proibe sem exceção
    extra3.update(id="X3", ativa=True)
    variantes += [extra, extra2, extra3]
    assert validar_regras(variantes) == []
    convertidas = [converter_v1(r) for r in variantes]
    assert validar_conjunto(convertidas, {}) == []
    rng = random.Random(11)
    for _ in range(2000):
        cabos, outros = cenario(rng)
        assert chaves(v1.avaliar(variantes, cabos, outros)) == chaves(avaliar(convertidas, cabos, outros, {})), (cabos, outros)


def test_motor_aceita_regra_v1_sem_converter_antes():
    cabos, outros = [], [L("DT11/300 1-CFU")]
    assert chaves(avaliar(ligar(SEED_V1), cabos, outros)) == chaves(v1.avaliar(ligar(SEED_V1), cabos, outros))


# ═══════════════════════════════════════════════════════════════════════════
# QUANTIDADE — o pedido que originou a task: "só quando for 1-CFU"
# ═══════════════════════════════════════════════════════════════════════════
CFU_QTD_1 = regra(id="CFU1", quando={"tem": "CFU", "qtd": {"=": 1}}, entao={"soma": "SUPL", "qtd": {">=": 1}})


@pytest.mark.parametrize("linha,dispara", [
    ("DT11/300 1-CFU", True),
    ("1-CFU", True),
    ("DT11/300 1-cfu", True),             # sem distinguir maiúsculas
    ("DT11/300 3-CFU", False),
    ("DT11/300 2-CFU 1-EF3H", False),
    ("DT11/300 11-CFU", False),
    ("DT11/300 *1-CFU", False),           # "*" = negativo
    ("DT11/300 1-CFUR", False),           # outro código
    ("DT11/300 1-CFU 1-SUPL", False),     # exigência atendida
    ("DT11/300 1-SUPL 1-CFU", False),
])
def test_so_dispara_para_1_cfu(linha, dispara):
    assert bool(alertas(CFU_QTD_1, [linha])) is dispara


@pytest.mark.parametrize("cmp_,valores_que_casam", [
    ({"=": 2}, {2}), ({"!=": 2}, {0, 1, 3}), ({">=": 2}, {2, 3}), ({"<=": 2}, {0, 1, 2}),
    ({">": 2}, {3}), ({"<": 2}, {0, 1}), ({"entre": [1, 2]}, {1, 2}), ({">=": 1, "<=": 2}, {1, 2}),
])
def test_todos_os_comparadores_de_quantidade(cmp_, valores_que_casam):
    r = regra(quando={"tem": "X", "qtd": cmp_})
    for q in (0, 1, 2, 3):
        assert bool(alertas(r, [f"{q}-X"])) is (q in valores_que_casam), (cmp_, q)


def test_soma_conta_todos_os_itens_e_tem_qtd_olha_item_a_item():
    linha = "1-P50 1-P50"                                     # duas ocorrências de 1
    assert alertas(regra(quando={"soma": "P50", "qtd": {"=": 2}}), [linha])
    assert not alertas(regra(quando={"tem": "P50", "qtd": {"=": 2}}), [linha])


def test_valor_dinamico_base_vezes_mais():
    r = regra(quando={"tem": "TR105"}, entao={"soma": "PR220", "qtd": {">=": {"base": 2, "vezes": 3, "mais": 1, "se": {"tem": "SI4"}}}})
    assert not alertas(r, ["1-TR105 2-PR220"])                  # sem SI4: exige 2
    assert alertas(r, ["1-TR105 6-PR220 1-SI4"])                # com SI4: exige 2*3+1 = 7
    assert not alertas(r, ["1-TR105 7-PR220 1-SI4"])


# ═══════════════════════════════════════════════════════════════════════════
# COMBINADORES, SELETORES E GRUPOS
# ═══════════════════════════════════════════════════════════════════════════
def test_combinadores_todos_algum_nenhum_nao_aninhados():
    r = regra(quando={"todos": [{"tem": "A"}, {"algum": [{"tem": "B"}, {"tem": "C"}]}, {"nenhum": [{"tem": "D"}]}]})
    assert alertas(r, ["1-A 1-B"]) and alertas(r, ["1-A 1-C"]) and alertas(r, ["1-A 1-B 1-C"])
    assert not alertas(r, ["1-A"]) and not alertas(r, ["1-A 1-B 1-D"]) and not alertas(r, ["1-B"])
    assert alertas(regra(quando={"nao": {"tem": "A"}}), ["1-Z"]) and not alertas(regra(quando={"nao": {"tem": "A"}}), ["1-A"])


def test_excecoes_e_condicao_texto():
    r = regra(quando={"tem": "TR105"}, excecoes={"tem": "CFU"}, entao={"nao": {"tem": "EF3H"}})
    assert alertas(r, ["1-TR105 1-EF3H"]) and not alertas(r, ["1-TR105 1-EF3H 1-CFU"]) and not alertas(r, ["1-TR105"])
    assert alertas(regra(quando={"texto": r"^POSTE\d"}), ["POSTE11 1-X"]) and not alertas(regra(quando={"texto": r"^POSTE\d"}), ["DT11/300"])


def test_seletores_codigo_curinga_regex_lista_e_intersecao():
    def casa(sel, nome):
        return bool(alertas(regra(quando={"tem": sel}), [f"1-{nome}"]))
    assert casa("CFU", "CFU") and casa("cfu", "CFU") and not casa("CFU", "CFUR")
    assert casa("TR[0-9]*", "TR125") and casa("TR[0-9]*", "TR330A") and not casa("TR[0-9]*", "TRAFO")
    assert casa("EF?H", "EF3H") and not casa("EF?H", "EF10K")
    assert casa("U[1-4]", "U3") and not casa("U[1-4]", "U9")        # curinga só com colchetes
    assert casa({"regex": "^U[1-4]$"}, "U3") and not casa({"regex": "^U[1-4]$"}, "U9")
    assert casa(["CFU", "CFA"], "CFA") and not casa(["CFU", "CFA"], "CL")
    assert casa({"e": ["TR*", {"regex": "A$"}]}, "TR330A") and not casa({"e": ["TR*", {"regex": "A$"}]}, "TR330")


def test_grupos_aninhados_e_uma_mudanca_no_grupo_altera_todas_as_regras_que_o_usam():
    novas, grupos = ligar(SEED["regras"]), copy.deepcopy(SEED["grupos"])
    linha = [L("DT10/300 1-U5")]                                # U5 não é estrutura MT no grupo original
    assert not [a for a in avaliar(novas, [], linha, grupos) if a["regra_id"] == "C2-POSTE10-MT"]
    grupos["ESTRUTURA_MT"].append("U5")                         # uma única edição no grupo...
    assert [a["regra_id"] for a in avaliar(novas, [], linha, grupos) if a["regra_id"] == "C2-POSTE10-MT"]
    isolada = [L("DT11/300 1-U3")]                              # ...e a regra de estrutura isolada enxerga o mesmo grupo
    assert [a for a in avaliar(novas, [], isolada, grupos) if a["regra_id"] == "C2-ESTR-ISOL"]
    grupos["ISOLADA"].append("U5")
    assert [a for a in avaliar(novas, [], [L("DT11/300 1-U5")], grupos) if a["regra_id"] == "C2-ESTR-ISOL"]
    assert not [a for a in avaliar(novas, [], [L("DT11/300 1-U5 1-U1")], grupos) if a["regra_id"] == "C2-ESTR-ISOL"]


@pytest.mark.parametrize("grupos,trecho", [
    ({"minusculo": ["X"]}, "MAIÚSCULAS"),
    ({"A": ["@B"], "B": ["@A"]}, "circular"),
    ({"A": ["@A"]}, "circular"),
    ({"A": ["@NADA"]}, "não existe"),
    ({"A": []}, "vazia"),
    ({"A": [{"regex": "(["}]}, "regex inválida"),
    ("nao-e-objeto", "objeto"),
])
def test_grupos_invalidos(grupos, trecho):
    assert any(trecho in e for e in validar_grupos(grupos)), validar_grupos(grupos)


# ═══════════════════════════════════════════════════════════════════════════
# ESCOPO PLANILHA — soma/contagem sobre a planilha inteira (decisão do usuário: além do P50)
# ═══════════════════════════════════════════════════════════════════════════
def total_postes(limite):
    return regra(id="POSTES", escopo="planilha", mensagem="Mais postes que o limite.",
                 deve_ser={"esq": {"soma_qtd": {"regex": r"^\d*(DT|CV)\d"}}, "cmp": "<=", "dir": limite,
                           "rotulos": {"esq": "Postes instalando", "dir": "Limite"}})


def test_total_de_postes_na_planilha_soma_quantidades_e_respeita_as_operacoes():
    outros = ["DT11/300 1-CFU", "3-DT9/150", ("DT11/300", "R")]      # 1 + 3 instalando; 1 retirado não conta
    assert not alertas(total_postes(4), outros)
    achado = alertas(total_postes(3), outros)
    assert len(achado) == 1 and achado[0]["linha_id"] == "GERAL"
    assert "Postes instalando = 4" in achado[0]["explicacao"] and "Limite = 3" in achado[0]["explicacao"]


def test_planilha_mistura_cabos_e_outros_e_rotula_a_tabela():
    r = regra(escopo="planilha", deve_ser={"esq": {"soma_metros": "P50"}, "cmp": ">=", "dir": {"vezes": [2, {"soma_qtd": "PR15"}]}})
    assert alertas(r, ["1-PR15 1-PR15"], [L("P50 ABC 3 m")])[0]["tabela"] == "cabos+outros"
    assert not alertas(r, ["1-PR15 1-PR15"], [L("P50 ABC 4 m")])
    assert alertas(total_postes(0), ["1-DT11/300"])[0]["tabela"] == "outros"


# ═══════════════════════════════════════════════════════════════════════════
# EXPLICAÇÃO E DESCRIÇÃO
# ═══════════════════════════════════════════════════════════════════════════
def test_achado_traz_explicacao_em_portugues():
    a = alertas(CFU_QTD_1, ["DT11/300 1-CFU"])[0]
    assert a["explicacao"] == "tem 1-CFU, mas a soma de SUPL é 0 (esperado ≥ 1)"
    assert "mensagem" in a and a["mensagem"] == "m"


def test_explicacao_da_regra_proibitiva_e_da_excecao():
    r = regra(quando={"tem": "TR105"}, excecoes={"tem": "CFU"}, entao={"nao": {"tem": "EF3H"}})
    assert alertas(r, ["1-TR105 1-EF3H"])[0]["explicacao"] == "tem 1-TR105, mas tem 1-EF3H (sem exceção: não tem CFU)"


def test_explicar_devolve_alerta_ok_e_nao_aplica_com_o_porque():
    linhas = [L("DT11/300 1-CFU"), L("DT11/300 1-CFU 1-SUPL"), L("DT11/300 3-CFU"), L("DT11/300 1-SUPL")]
    situacoes = [x["situacao"] for x in explicar([CFU_QTD_1], [], linhas)]
    assert situacoes == ["alerta", "ok", "nao_aplica", "nao_aplica"]
    porque = explicar([CFU_QTD_1], [], linhas)[2]["explicacao"]
    assert "3-CFU" in porque and "quantidade esperada é = 1" in porque
    assert explicar([regra(ativa=False, quando={"tem": "X"})], [], linhas) == []      # desligada não entra


def test_descrever_gera_frase_para_toda_regra_da_semente_sem_mostrar_json():
    for r in SEED["regras"]:
        frase = descrever(r, SEED["grupos"])
        assert frase.startswith(("Em cada linha de Outros", "Na planilha inteira")) and "Gravidade:" in frase
        assert '{"' not in frase and '"quando"' not in frase and "'tem'" not in frase


def test_descrever_frases_chave():
    assert descrever(CFU_QTD_1) == ("Em cada linha de Outros (operação I, *I): se a linha tem CFU com quantidade = 1, "
                                    "então a soma de SUPL deve ser ≥ 1. Gravidade: aviso.")
    proibida = regra(quando={"tem": "X"})
    assert "alerta quando a linha tem X" in descrever(proibida)
    assert "Na planilha inteira" in descrever(total_postes(30)) and "≤ 30" in descrever(total_postes(30))
    assert descrever(SEED_V1[0]).startswith("Em cada linha de Outros")     # regra v1 também tem frase


# ═══════════════════════════════════════════════════════════════════════════
# API
# ═══════════════════════════════════════════════════════════════════════════
@pytest.fixture
def client(tmp_path, monkeypatch):
    monkeypatch.setattr(database, "DB_PATH", str(tmp_path / "teste.db"))
    database.init_db()
    return TestClient(appmod.app)


def cab(role="admin"):
    return {"Authorization": f"Bearer {create_jwt_token('u1', f'{role}@x.com', role)}"}


def test_get_devolve_regras_v2_com_frase_e_grupos(client):
    r = client.get("/api/validacao/regras", headers=cab("operador")).json()
    assert len(r["regras"]) == 11 and "ESTRUTURA_MT" in r["grupos"]
    assert all(x["versao"] == 2 and x["frase"] for x in r["regras"])


def test_descrever_e_testar_com_explicacao_pela_api(client):
    d = client.post("/api/validacao/regras/descrever", json={"regra": CFU_QTD_1}, headers=cab()).json()
    assert d["erros"] == [] and "quantidade = 1" in d["frase"]
    ruim = client.post("/api/validacao/regras/descrever", json={"regra": {**CFU_QTD_1, "quando": {"tem": "@NADA"}}}, headers=cab()).json()
    assert ruim["frase"] == "" and any("não existe" in e for e in ruim["erros"])
    t = client.post("/api/validacao/regras/testar", headers=cab(), json={
        "regras": [CFU_QTD_1], "grupos": {}, "outros": [L("DT11/300 1-CFU"), L("DT11/300 3-CFU")], "explicar": True}).json()
    assert [a["linha_id"] for a in t["achados"]] == ["OUTROS-0"]
    assert [x["situacao"] for x in t["explicacoes"]] == ["alerta", "nao_aplica"]


def test_salvar_regra_com_condicao_de_quantidade_e_grupo_novo_pela_api(client):
    grupos = {**SEED["grupos"], "MEUS_POSTES": ["DT11/300", "DT11/600"]}
    nova = regra(id="MINHA", quando={"tem": "@MEUS_POSTES"}, entao={"soma": "SUPL", "qtd": {">=": 1}})
    regras = client.get("/api/validacao/regras", headers=cab()).json()["regras"] + [nova]
    assert client.post("/api/validacao/regras", json={"projeto_codigo": "DEFAULT", "regras": regras, "grupos": grupos}, headers=cab()).status_code == 200
    r = client.get("/api/validacao/regras", headers=cab()).json()
    assert "MEUS_POSTES" in r["grupos"] and any(x["id"] == "MINHA" for x in r["regras"])
    ruim = client.post("/api/validacao/regras", json={"projeto_codigo": "DEFAULT", "regras": regras, "grupos": {"a": ["X"]}}, headers=cab())
    assert ruim.status_code == 400 and any("MAIÚSCULAS" in e for e in ruim.json()["detail"]["erros"])


def test_regra_v1_salva_antes_continua_funcionando_pela_rota_de_planilhas(client):
    """Instalação antiga: o banco tem uma LISTA de regras v1. Nada deve quebrar nem mudar."""
    conn = database.get_connection()
    conn.execute("INSERT OR REPLACE INTO regras_dominio (projeto_codigo, regras_json) VALUES ('229', ?)",
                 (json.dumps(ligar(SEED_V1)),))
    conn.commit(); conn.close()
    r = client.post("/api/validacao/planilhas", headers=cab("operador"), json={
        "outros": [L("DT11/300 1-CFU")], "projeto_codigo": "229"}).json()
    assert "C2-CFU-SUPL" in [a["regra_id"] for a in r["achados"]]
    assert client.get("/api/validacao/regras?projeto_codigo=229", headers=cab()).json()["regras"][0]["versao"] == 2
