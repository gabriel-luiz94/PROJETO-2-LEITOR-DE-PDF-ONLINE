"""tests/test_validacao_planilhas.py — Camada 1 da validação (TASK-011) e não-regressão do parser."""
import pytest

from services.orcamento_calc import extrair_ativos_outros, processar_calculo, tokenizar_outros
from services.validacao_planilhas import validar_base, validar_planilhas


def regras(achados):
    return sorted(a["regra_id"] for a in achados)


def cabo(ativo, op="I"):
    return validar_planilhas([{"ativo": ativo, "operacao": op}], [])


def outro(ativo, op="I"):
    return validar_planilhas([], [{"ativo": ativo, "operacao": op}])


# ── Cabos ────────────────────────────────────────────────────────────────────
@pytest.mark.parametrize("ativo", [
    "CAA 2 ABC 35 m",     # nome com espaço (formato da tela)
    "CAA2 ABC 35",        # sem o "m"
    "M3X1X70+70 3 20 m",
    "P50",                # standalone: herda a fase da linha seguinte (RN-04)
    "CAA 2 35 m",         # comprimento sem fase é aceito pelo frontend
    "CAA 2 ABC 3,5 m",    # vírgula decimal
])
def test_cabos_validos(ativo):
    assert cabo(ativo) == []


def test_cabo_no_formato_de_outros():
    assert regras(cabo("3-CFU")) == ["C1-CABO-PARTE"]


def test_cabo_comprimento_nao_numerico_ou_ausente():
    assert regras(cabo("CAA 2 ABC xx m")) == ["C1-CABO-NUM"]
    assert regras(cabo("CAA 2 ABC")) == ["C1-CABO-NUM"]


def test_cabo_comprimento_zero_e_aviso():
    achados = cabo("CAA 2 ABC 0 m")
    assert regras(achados) == ["C1-QTD-ZERO"]
    assert achados[0]["severidade"] == "aviso"


def test_linha_sem_ativo_e_ignorada_como_no_calculo():
    assert cabo("", op="") == []
    assert outro("   ", op="") == []


# ── Outros ───────────────────────────────────────────────────────────────────
@pytest.mark.parametrize("ativo", [
    "DT11/300 1-CFU 1-EF3H",  # poste sem quantidade abre a linha
    "3-IP 2-RECAL",
    "*1-TERRA3",              # "*" = negativo, aceito pelo parser
    "1,5-P50",
])
def test_outros_validos(ativo):
    assert outro(ativo) == []


def test_outros_token_orfao_seria_descartado():
    achados = outro("3-IP RECAL")
    assert regras(achados) == ["C1-OUT-ORFAO"]
    assert "RECAL" in achados[0]["mensagem"]


def test_outros_quantidade_nao_numerica():
    assert regras(outro("1-A x-CFU")) == ["C1-OUT-NUM"]            # "x-CFU" sozinho é um nome (compacto); com um "qtd-ativo" na linha vale o formato antigo


def test_outros_no_formato_de_cabo():
    assert regras(outro("CAA 2 35 m")) == ["C1-OUT-CABO"]


def test_outros_quantidade_zero_e_duplicidade_sao_aviso():
    assert regras(outro("0-P50")) == ["C1-QTD-ZERO"]
    achados = outro("1-P50 1-P50")
    assert regras(achados) == ["C1-DUP"]
    assert achados[0]["severidade"] == "aviso"


# ── Operação e ids ───────────────────────────────────────────────────────────
@pytest.mark.parametrize("op", ["I", "*I", "R", "*R", "M", "*M", "i"])
def test_operacoes_validas(op):
    assert cabo("P50", op) == []


@pytest.mark.parametrize("op", ["", "X", "0"])
def test_operacao_invalida(op):
    assert regras(cabo("P50", op)) == ["C1-OP"]
    assert regras(outro("1-P50", op)) == ["C1-OP"]


def test_ids_padrao_seguem_o_contexto_do_chat_e_id_explicito_prevalece():
    achados = validar_planilhas(
        [{"ativo": "P50", "operacao": "I"}, {"id": "CABOS-9", "ativo": "3-CFU", "operacao": "I"}],
        [{"ativo": "1-P50", "operacao": "I"}, {"ativo": "1-A x-CFU", "operacao": "I"}],
    )
    assert sorted(a["linha_id"] for a in achados) == ["CABOS-9", "OUTROS-1"]


# ── Paridade com o parser do cálculo ─────────────────────────────────────────
@pytest.mark.parametrize("ativo", [
    "DT11/300 1-CFU 1-EF3H", "3-IP RECAL", "*1-TERRA3", "cv11/300 2-P50", "1,5-P50", "3-IP 2-RECAL",
])
def test_tokenizacao_e_a_mesma_do_calculo(ativo):
    tokens = tokenizar_outros(ativo)
    itens = extrair_ativos_outros({"ativo": ativo, "operacao": "I"})
    assert [i["ativo"] for i in itens] == [t.upper() for t in tokens[1::2]]


# ── C1-BASE ──────────────────────────────────────────────────────────────────
BASE = [
    {"ativo": "P50", "componente": "C1", "codigo": "100", "mdo": "MATERIAL", "fator_i": 1, "fator_r": 1,
     "desc_codigo": "CABO", "projeto": "", "origem": "CABOS"},
    {"ativo": "CFU", "componente": "C2", "codigo": "200", "mdo": "MATERIAL", "fator_i": 1, "fator_r": 1,
     "desc_codigo": "CHAVE", "projeto": "", "origem": "OUTROS"},
]


def test_base_encontra_e_aponta_ausentes():
    achados = validar_base(
        [{"id": "T-C1", "ativo": "P50 1 10", "operacao": "I"}, {"id": "T-C2", "ativo": "XYZ 1 5", "operacao": "I"}],
        [{"id": "T-O1", "ativo": "1-CFU", "operacao": "I"}, {"id": "T-O2", "ativo": "2-ZZZ", "operacao": "I"}],
        "RONDONIA", BASE,
    )
    assert [(a["linha_id"], a["regra_id"]) for a in achados] == [("T-C2", "C1-BASE"), ("T-O2", "C1-BASE")]


def test_base_respeita_o_filtro_de_origem():
    # P50 só existe para CABOS: em OUTROS não é encontrado
    achados = validar_base([], [{"id": "T-O1", "ativo": "1-P50", "operacao": "I"}], None, BASE)
    assert [a["linha_id"] for a in achados] == ["T-O1"]


# ── Não-regressão do cálculo (extração do parser não mudou o resultado) ──────
def test_processar_calculo_inalterado():
    res = processar_calculo([{"ativo": "P50 1 10", "operacao": "I"}], [{"ativo": "1-CFU", "operacao": "R"}], None, BASE)
    assert res["nao_encontrados"] == []
    assert [(r["operacao"], r["codigo"], r["total"]) for r in res["resultado"]] == [("I", "100", 10.0), ("R", "200", 1.0)]


# ── TASK-122: C1-BASE avalia SÓ o nome; parâmetros têm regra própria (C1-PARAM); a descrição não resolve nome ──
def _lp(ativo, comp="", cod="", desc="", origem="OUTROS"):
    return {"ativo": ativo, "desc_ativo": desc, "componente": comp, "codigo": cod, "mdo": "MATERIAL", "fator_i": 1, "fator_r": 1, "desc_codigo": "d", "projeto": "PERNAMBUCO",
            "origem": origem, "lv": "A"}


BASE_PARAM = [_lp("IT_2=IT_2(poste=D,cabobt=35), IT-2"), _lp("IT_2", "IT_2", "3417025", "IT-2"), _lp("IT_2(poste=%D)", "IT_2.D", "3480305", "FLABIT"),
              _lp("IT_2(poste=%R)", "IT_2.R", "3480270", "FLABIT [ R ] POSTE CIRCULAR"), _lp("CFU", "CFU", "200", "CHAVE")]


def _val(texto, base=BASE_PARAM, projeto="PERNAMBUCO"):
    return validar_base([], [{"id": "L1", "ativo": texto, "operacao": "I"}], projeto, [dict(r) for r in base])


def test_c1_base_avalia_so_o_nome_e_a_descricao_nao_resolve():
    a = _val("1-FLABIT(D,35)")                                  # FLABIT só existe como DESCRIÇÃO: o nome não existe (e a mensagem não leva os parâmetros)
    assert [(x["regra_id"], x["mensagem"]) for x in a] == [("C1-BASE", "Ativo 'FLABIT' não encontrado na base técnica (projeto PERNAMBUCO).")]
    assert [x["regra_id"] for x in _val("1-FLABIT")] == ["C1-BASE"]                # sem parâmetros também (antes achava 1 linha pela descrição)
    assert _val("1-IT_2(D,35)") == [] and _val("1-IT-2(poste=D)".replace("1-IT-2", "IT-2")) == []      # ativo e apelido com parâmetros certos: nada a avisar


def test_parametros_com_problema_vao_para_c1_param_e_nao_para_c1_base():
    a = _val("1-IT_2(XX,35)")                                    # o nome existe; o valor do parâmetro não está cadastrado
    assert [x["regra_id"] for x in a] == ["C1-PARAM"] and "POSTE=XX não cadastrado" in a[0]["mensagem"] and "IT_2(XX,35)" in a[0]["mensagem"]
    a = _val("1-IT_2(D,35,7)")                                   # valores demais
    assert [x["regra_id"] for x in a] == ["C1-PARAM"]
    a = _val("2-CFU(A,B)")                                       # CFU existe mas não declara parâmetros
    assert [x["regra_id"] for x in a] == ["C1-PARAM"] and "sem parâmetros declarados" in a[0]["mensagem"]


def test_c1_base_origem_e_cabos_continuam_valendo():
    base = BASE_PARAM + [_lp("P50", "C1", "100", origem="CABOS")]
    assert [x["regra_id"] for x in validar_base([], [{"id": "O", "ativo": "1-P50", "operacao": "I"}], "PERNAMBUCO", [dict(r) for r in base])] == ["C1-BASE"]   # P50 só para CABOS
    a = validar_base([{"id": "C", "ativo": "XYZ 1 5", "operacao": "I"}], [], "PERNAMBUCO", [dict(r) for r in base])
    assert [(x["linha_id"], x["regra_id"], x["tabela"]) for x in a] == [("C", "C1-BASE", "cabos")]
