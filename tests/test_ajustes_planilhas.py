"""tests/test_ajustes_planilhas.py — motor de ajustes das planilhas (TASK-022, ADR-006).

Decisões do usuário: ajuste pode mudar a operação só se for determinado explicitamente; a lista de 7 ações está completa;
nada se aplica sem pré-visualização (o motor só devolve o diff)."""
import copy

import pytest
from fastapi.testclient import TestClient

import app as appmod
import database
from middleware.auth_middleware import create_jwt_token
from services.ajustes_planilhas import ajustar, descrever_acao, validar_acoes

GRUPOS = {"SUPLS": ["SUPL", "SUP-L"], "CHAVES": ["CFU", "CFUR"]}


def o(ativo, op="I", **extra):
    return {"operacao": op, "ativo": ativo, **extra}


def textos(res, tabela="outros"):
    return [r["ativo"] for r in res[tabela]]


def acao(nome, tabela="outros", **kw):
    return {"acao": nome, "tabela": tabela, **kw}


def rodar(acoes, outros=(), cabos=(), grupos=None):
    return ajustar(acoes, list(cabos), list(outros), grupos)


def idempotente(acoes, outros=(), cabos=(), grupos=None):
    r1 = rodar(acoes, outros, cabos, grupos)
    r2 = ajustar(acoes, r1["cabos"], r1["outros"], grupos)
    assert r2["operacoes"] == [], r2["operacoes"]
    return r1


# ── substituir ───────────────────────────────────────────────────────────────
def test_substituir_sup_l_por_supl_por_palavra_inteira():
    r = idempotente([acao("substituir", de="SUP-L", para="SUPL")], [o("DT11/300 1-CFU 1-SUP-L"), o("1-SUP-LX"), o("1-sup-l")])
    assert textos(r) == ["DT11/300 1-CFU 1-SUPL", "1-SUP-LX", "1-SUPL"]   # sem distinguir maiúsculas; SUP-LX intacto


def test_substituir_sem_palavra_inteira_pega_pedaco():
    r = rodar([acao("substituir", de="SUP-L", para="SUPL", palavra_inteira=False)], [o("1-SUP-LX")])
    assert textos(r) == ["1-SUPLX"]


def test_substituir_regex_com_referencia():
    r = idempotente([acao("substituir", de={"regex": r"\bSUP[-_ ]?L\b"}, para="SUPL")], [o("1-SUP_L 1-SUP-L 1-SUPL")])
    assert textos(r) == ["1-SUPL 1-SUPL 1-SUPL"]
    r = rodar([acao("substituir", de={"regex": r"^(\d)-DT"}, para=r"\1-DT")], [o("1-DT11/300")])
    assert r["operacoes"] == []


def test_substituir_modo_item_usa_seletor_grupo_e_preserva_quantidade():
    r = rodar([acao("substituir", modo="item", de="@CHAVES", para="CFA")], [o("DT11/300 2-CFU 1-CFUR 3-SUPL")], grupos=GRUPOS)
    assert textos(r) == ["DT11/300 2-CFA 1-CFA 3-SUPL"]
    r = rodar([acao("substituir", modo="item", de="TR*", para="TR3")], [o("DT11/300 1-TR100 *1-TR200")])
    assert textos(r) == ["DT11/300 1-TR3 *1-TR3"]


def test_substituir_em_cabos_e_ambos():
    r = rodar([acao("substituir", "ambos", de="ABC", para="XYZ")], [o("1-ABC")], [o("CAA 2 ABC 35 m")])
    assert textos(r, "cabos") == ["CAA 2 XYZ 35 m"] and textos(r) == ["1-XYZ"]


# ── operação: só muda se determinado explicitamente ──────────────────────────
def test_operacao_e_mantida_por_padrao_e_muda_so_com_operacao_nova():
    sem = rodar([acao("substituir", de="SUP-L", para="SUPL")], [o("1-SUP-L", "I")])
    assert sem["outros"][0]["operacao"] == "I"
    com = rodar([acao("substituir", de="SUP-L", para="SUPL", operacao_nova="*i")], [o("1-SUP-L", "I"), o("1-OUTRO", "I")])
    assert [r["operacao"] for r in com["outros"]] == ["*I", "I"]          # só na linha alterada
    assert com["operacoes"][0]["antes"]["operacao"] == "I" and com["operacoes"][0]["depois"]["operacao"] == "*I"


def test_operacao_nova_nao_se_aplica_se_o_texto_nao_mudou():
    r = rodar([acao("substituir", de="NADA", para="X", operacao_nova="R")], [o("1-CFU")])
    assert r["operacoes"] == [] and r["outros"][0]["operacao"] == "I"


# ── filtros: operacoes e quando ──────────────────────────────────────────────
def test_filtro_de_operacoes_e_condicao_quando():
    linhas = [o("1-AA", "I"), o("1-AA", "R"), o("1-AA 1-CFU", "I")]
    r = rodar([acao("substituir", de="AA", para="BB", operacoes=["I"], quando={"tem": "CFU"})], linhas)
    assert textos(r) == ["1-AA", "1-AA", "1-BB 1-CFU"]
    r = rodar([acao("substituir", de="AA", para="BB", operacoes=["i", "*I"])], linhas)
    assert textos(r) == ["1-BB", "1-AA", "1-BB 1-CFU"]


# ── normalizar ───────────────────────────────────────────────────────────────
def test_normalizar_espacos_maiusculas_e_poste():
    r = idempotente([acao("normalizar", regras=["espacos", "maiusculas", "poste"])],
                    [o("  1-dt11/300   1-cfu "), o("1-AB"), o("DT10/300")])
    assert textos(r) == ["DT11/300 1-CFU", "1-AB", "DT10/300"]


def test_normalizar_padrao_so_espacos():
    assert textos(rodar([acao("normalizar")], [o("1-a   1-b")])) == ["1-a 1-b"]


def test_normalizar_poste_so_no_inicio_da_linha_e_mesclar_nao_reformata_quem_nao_repete():
    assert textos(rodar([acao("normalizar", regras=["poste"])], [o("DT11/300 1-DT10"), o("  1-CV10 1-X")])) == ["DT11/300 1-DT10", "  CV10 1-X"]
    assert textos(rodar([acao("mesclar_duplicadas")], [o("2,5-X 01-Y")])) == ["2,5-X 01-Y"]      # nada repetido: texto intacto


# ── ordenar ──────────────────────────────────────────────────────────────────
def test_ordenar_estavel_natural_e_por_operacao_personalizada():
    linhas = [o("1-B", "R"), o("1-A10", "I"), o("1-A2", "I"), o("1-C", "*I")]
    r = idempotente([acao("ordenar", por=[{"coluna": "operacao", "valores": ["I", "*I", "R"]}, {"coluna": "ativo"}])], linhas)
    assert [(x["operacao"], x["ativo"]) for x in r["outros"]] == [("I", "1-A2"), ("I", "1-A10"), ("*I", "1-C"), ("R", "1-B")]
    mov = [x for x in r["operacoes"] if x["op"] == "mover"]
    assert len(mov) == 1 and mov[0]["ordem_antes"] == ["OUTROS-0", "OUTROS-1", "OUTROS-2", "OUTROS-3"]
    assert mov[0]["ordem_depois"] == ["OUTROS-2", "OUTROS-1", "OUTROS-3", "OUTROS-0"] and mov[0]["acoes"] == [0]
    assert r["resumo"]["editar"] == 0                                    # só reordenou: nenhuma linha editada


def test_ordenar_desc_e_ja_ordenado_nao_gera_operacao():
    r = rodar([acao("ordenar", por=[{"coluna": "ativo", "ordem": "desc"}])], [o("1-A"), o("1-C"), o("1-B")])
    assert textos(r) == ["1-C", "1-B", "1-A"]
    assert rodar([acao("ordenar", por=[{"coluna": "ativo"}])], [o("1-A"), o("1-B")])["operacoes"] == []


def test_ordenar_cabos_avisa_sobre_a_fase_herdada():
    r = rodar([acao("ordenar", "cabos", por=[{"coluna": "ativo"}])], cabos=[o("CAA 2 B 3 m"), o("CAA 2 A 3 m")])
    assert textos(r, "cabos") == ["CAA 2 A 3 m", "CAA 2 B 3 m"] and any("Cabos" in a for a in r["avisos"])


# ── excluir_linhas ───────────────────────────────────────────────────────────
def test_excluir_vazias_duplicadas_condicao_e_texto():
    linhas = [o(""), o("   "), o("1-A"), o("1-a", "I"), o("1-A", "R"), o("1-CFU 1-X")]
    assert textos(idempotente([acao("excluir_linhas", onde={"vazias": True})], linhas)) == ["1-A", "1-a", "1-A", "1-CFU 1-X"]
    r = idempotente([acao("excluir_linhas", onde={"duplicadas": True})], linhas)
    assert textos(r) == ["", "   ", "1-A", "1-A", "1-CFU 1-X"]            # mantém a primeira; R é outra linha
    assert textos(rodar([acao("excluir_linhas", onde={"condicao": {"tem": "CFU"}})], linhas)) == ["", "   ", "1-A", "1-a", "1-A"]
    r = rodar([acao("excluir_linhas", onde={"texto": "^1-A$"}, operacoes=["R"])], linhas)
    assert len(r["outros"]) == 5 and [x["op"] for x in r["operacoes"]] == ["excluir"]
    assert r["operacoes"][0]["antes"] == {"operacao": "R", "ativo": "1-A"}


# ── adicionar_linha ──────────────────────────────────────────────────────────
def test_adicionar_linha_fim_inicio_e_so_se_nao_existir():
    a_fim = acao("adicionar_linha", valores={"operacao": "I", "ativo": "1-RA2"})
    r = idempotente([a_fim], [o("1-A")])
    assert textos(r) == ["1-A", "1-RA2"] and r["outros"][1]["id"].startswith("NOVA-")
    assert r["operacoes"][0] == {"op": "inserir", "tabela": "outros", "linha_id": r["outros"][1]["id"],
                                 "depois": {"operacao": "I", "ativo": "1-RA2", "entidade": "0"}, "depois_de": "OUTROS-0", "acoes": [0]}
    assert textos(rodar([acao("adicionar_linha", valores={"operacao": "I", "ativo": "1-RA2"}, posicao="inicio")], [o("1-A")])) == ["1-RA2", "1-A"]
    dup = rodar([acao("adicionar_linha", valores={"operacao": "I", "ativo": "1-RA2"}, apenas_se_nao_existir=False)], [o("1-RA2")])
    assert textos(dup) == ["1-RA2", "1-RA2"]


def test_adicionar_linha_depois_e_antes_de_cada_linha_que_satisfaz():
    linhas = [o("DT11/300 1-SI3"), o("DT11/300 1-CFU"), o("DT11/300 1-SI4")]
    val = {"operacao": "I", "ativo": "1-RA2"}
    r = idempotente([acao("adicionar_linha", valores=val, posicao={"depois_de": {"tem": "SI*"}})], linhas)
    assert textos(r) == ["DT11/300 1-SI3", "1-RA2", "DT11/300 1-CFU", "DT11/300 1-SI4", "1-RA2"]
    r = idempotente([acao("adicionar_linha", valores=val, posicao={"antes_de": {"tem": "CFU"}})], linhas)
    assert textos(r) == ["DT11/300 1-SI3", "1-RA2", "DT11/300 1-CFU", "DT11/300 1-SI4"]


# ── adicionar_ativo / remover_ativo / mesclar ────────────────────────────────
def test_adicionar_ativo_se_ja_existe():
    linhas = [o("DT11/300 1-CFU"), o("DT11/300 1-CFU 1-SUPL"), o("DT11/300 1-CFU 2-SUPL")]
    cond = {"tem": "CFU"}
    r = idempotente([acao("adicionar_ativo", ativo="SUPL", quando=cond)], linhas)
    assert textos(r) == ["DT11/300 1-CFU 1-SUPL", "DT11/300 1-CFU 1-SUPL", "DT11/300 1-CFU 2-SUPL"]
    assert textos(rodar([acao("adicionar_ativo", ativo="SUPL", qtd=2, se_ja_existe="somar", quando=cond)], linhas)) == \
        ["DT11/300 1-CFU 2-SUPL", "DT11/300 1-CFU 3-SUPL", "DT11/300 1-CFU 4-SUPL"]
    assert textos(rodar([acao("adicionar_ativo", ativo="SUPL", qtd=5, se_ja_existe="substituir")], linhas))[2] == "DT11/300 1-CFU 5-SUPL"


def test_adicionar_ativo_qtd_negativa_passa_na_validacao_mas_zero_continua_rejeitado():
    assert erros_de(acao("adicionar_ativo", ativo="PR", qtd=-1), GRUPOS) == []
    assert any("'qtd'" in e for e in erros_de(acao("adicionar_ativo", ativo="PR", qtd=0), GRUPOS))


def test_adicionar_ativo_com_quantidade_negativa_gera_prefixo_asterisco():
    """qtd negativo ("retirar"/linha viva) tem que sair como "*<n>-ATIVO", nunca "-<n>-ATIVO" —
    o tokenizador de Outros (orcamento_calc.tokenizar_outros) só lê sinal negativo via "*"."""
    r = rodar([acao("adicionar_ativo", ativo="PR", qtd=-1)], [o("1-TR110")])
    assert textos(r) == ["1-TR110 *1-PR"]


def test_adicionar_ativo_negativo_somando_com_token_existente_tambem_negativo():
    linhas = [o("1-TR110 *1-PR")]
    r = rodar([acao("adicionar_ativo", ativo="PR", qtd=-2, se_ja_existe="somar")], linhas)
    assert textos(r) == ["1-TR110 *3-PR"]


def test_adicionar_ativo_negativo_somando_com_token_existente_positivo_pode_virar_positivo():
    linhas = [o("1-TR110 5-PR")]
    r = rodar([acao("adicionar_ativo", ativo="PR", qtd=-2, se_ja_existe="somar")], linhas)
    assert textos(r) == ["1-TR110 3-PR"]   # 5 + (-2) = 3: continua positivo, sem "*"


def test_adicionar_ativo_negativo_substituindo_token_existente():
    r = rodar([acao("adicionar_ativo", ativo="PR", qtd=-4, se_ja_existe="substituir")], [o("1-TR110 1-PR")])
    assert textos(r) == ["1-TR110 *4-PR"]


def test_adicionar_ativo_em_linha_com_condicao_de_quantidade_exata():
    # o pedido original do usuário: só quando for 1-CFU
    linhas = [o("DT11/300 1-CFU"), o("DT11/300 2-CFU")]
    r = rodar([acao("adicionar_ativo", ativo="SUPL", quando={"tem": "CFU", "qtd": {"=": 1}})], linhas)
    assert textos(r) == ["DT11/300 1-CFU 1-SUPL", "DT11/300 2-CFU"]


# ── adicionar_ativo com qtd dinâmica (TASK-045) ──────────────────────────────
def test_adicionar_ativo_qtd_dinamica_mesma_quantidade_negativa():
    """Pedido real do usuário: "se tem U3, adicionar 90277 na mesma quantidade de U3, negativo"."""
    linhas = [o("1-U3"), o("3-U3 1-CFU"), o("DT11/300 1-CFU")]
    acoes = [acao("adicionar_ativo", ativo="90277", qtd={"soma": "U3", "fator": -1}, quando={"tem": "U3"})]
    r = idempotente(acoes, linhas)
    assert textos(r) == ["1-U3 *1-90277", "3-U3 1-CFU *3-90277", "DT11/300 1-CFU"]


def test_adicionar_ativo_qtd_dinamica_sem_fator_e_positiva():
    r = rodar([acao("adicionar_ativo", ativo="X", qtd={"soma": "U3"})], [o("2-U3")])
    assert textos(r) == ["2-U3 2-X"]


def test_adicionar_ativo_qtd_dinamica_com_fator_diferente_de_1():
    r = rodar([acao("adicionar_ativo", ativo="X", qtd={"soma": "U3", "fator": 2})], [o("2-U3")])
    assert textos(r) == ["2-U3 4-X"]


def test_adicionar_ativo_qtd_dinamica_soma_zero_nao_adiciona_nada():
    # sem "quando", a ação roda em toda linha; se o seletor não casar nessa linha, soma = 0 = no-op
    r = rodar([acao("adicionar_ativo", ativo="90277", qtd={"soma": "U3", "fator": -1})], [o("DT11/300 1-CFU")])
    assert textos(r) == ["DT11/300 1-CFU"]
    assert r["operacoes"] == []


def test_adicionar_ativo_qtd_dinamica_soma_varios_itens_do_mesmo_seletor():
    # usa @CHAVES (CFU/CFUR, sem hífen no código) para não disparar o guard pré-existente de
    # contrato com códigos hifenados (ex.: SUP-L) combinados a um token negativo "*" — ver
    # test_adicionar_ativo_qtd_dinamica_soma_zero_nao_adiciona_nada para o comportamento isolado.
    r = rodar([acao("adicionar_ativo", ativo="X", qtd={"soma": "@CHAVES", "fator": -1})], [o("1-CFU 2-CFUR")], grupos=GRUPOS)
    assert textos(r) == ["1-CFU 2-CFUR *3-X"]


def test_adicionar_ativo_qtd_dinamica_respeita_se_ja_existe():
    r = rodar([acao("adicionar_ativo", ativo="90277", qtd={"soma": "U3", "fator": -1}, se_ja_existe="somar")],
              [o("2-U3 *1-90277")])
    assert textos(r) == ["2-U3 *3-90277"]


def test_adicionar_ativo_qtd_dinamica_valida_schema():
    assert erros_de(acao("adicionar_ativo", ativo="X", qtd={"soma": "U3", "fator": -1}), GRUPOS) == []
    assert erros_de(acao("adicionar_ativo", ativo="X", qtd={"soma": "U3"}), GRUPOS) == []
    assert any("qtd" in e for e in erros_de(acao("adicionar_ativo", ativo="X", qtd={}), GRUPOS))
    assert any("qtd" in e for e in erros_de(acao("adicionar_ativo", ativo="X", qtd={"soma": "U3", "fator": "a"}), GRUPOS))
    assert any("qtd" in e for e in erros_de(acao("adicionar_ativo", ativo="X", qtd={"soma": "U3", "extra": 1}), GRUPOS))
    assert any("@NAOEXISTE" in e for e in erros_de(acao("adicionar_ativo", ativo="X", qtd={"soma": "@NAOEXISTE"}), GRUPOS))


def test_adicionar_ativo_descreve_qtd_dinamica_em_portugues():
    assert descrever_acao(acao("adicionar_ativo", ativo="90277", qtd={"soma": "U3", "fator": -1})) == \
        "Em Outros: adicionar '90277' na mesma quantidade de U3 (negativa) à linha."
    assert descrever_acao(acao("adicionar_ativo", ativo="X", qtd={"soma": "U3"})) == \
        "Em Outros: adicionar 'X' na mesma quantidade de U3 à linha."


# ── adicionar_ativo com ativo dinâmico (TASK-046) ────────────────────────────
def test_adicionar_ativo_dinamico_mesma_quantidade_negativa_pedido_real_do_usuario():
    """Pedido real: "se tem TR1*, adicione o próprio texto + VP no fim, na mesma quantidade negativa"."""
    linhas = [o("1-TR110"), o("2-TR127 1-CFU"), o("DT11/300 1-CFU")]
    acoes = [acao("adicionar_ativo", ativo={"igual_a": "TR1*", "sufixo": "VP"}, qtd=-1, quando={"tem": "TR1*"})]
    r = idempotente(acoes, linhas)
    assert textos(r) == ["1-TR110 *1-TR110VP", "2-TR127 1-CFU *2-TR127VP", "DT11/300 1-CFU"]


def test_adicionar_ativo_dinamico_sem_fator_e_positivo():
    r = rodar([acao("adicionar_ativo", ativo={"igual_a": "TR1*", "sufixo": "VP"})], [o("3-TR115")])
    assert textos(r) == ["3-TR115 3-TR115VP"]


def test_adicionar_ativo_dinamico_com_prefixo_e_sufixo():
    r = rodar([acao("adicionar_ativo", ativo={"igual_a": "TR1*", "prefixo": "X", "sufixo": "Y"}, qtd=1)], [o("1-TR110")])
    assert textos(r) == ["1-TR110 1-XTR110Y"]


def test_adicionar_ativo_dinamico_sem_match_nao_adiciona_nada():
    r = rodar([acao("adicionar_ativo", ativo={"igual_a": "TR1*", "sufixo": "VP"}, qtd=-1)], [o("DT11/300 1-CFU")])
    assert textos(r) == ["DT11/300 1-CFU"]
    assert r["operacoes"] == []


def test_adicionar_ativo_dinamico_varios_matches_na_mesma_linha_geram_varios_tokens():
    r = rodar([acao("adicionar_ativo", ativo={"igual_a": "TR1*", "sufixo": "VP"}, qtd=-1)], [o("1-TR110 2-TR115")])
    assert textos(r) == ["1-TR110 2-TR115 *1-TR110VP *2-TR115VP"]


def test_adicionar_ativo_dinamico_respeita_se_ja_existe():
    r = rodar([acao("adicionar_ativo", ativo={"igual_a": "TR1*", "sufixo": "VP"}, qtd=-1, se_ja_existe="somar")],
              [o("1-TR110 *2-TR110VP")])
    assert textos(r) == ["1-TR110 *3-TR110VP"]


def test_adicionar_ativo_dinamico_valida_schema():
    assert erros_de(acao("adicionar_ativo", ativo={"igual_a": "TR1*", "sufixo": "VP"}, qtd=-1), GRUPOS) == []
    assert erros_de(acao("adicionar_ativo", ativo={"igual_a": "TR1*"}), GRUPOS) == []
    assert any("ativo" in e for e in erros_de(acao("adicionar_ativo", ativo={}), GRUPOS))
    assert any("ativo" in e for e in erros_de(acao("adicionar_ativo", ativo={"igual_a": "TR1*", "extra": 1}), GRUPOS))
    assert any("ativo" in e for e in erros_de(acao("adicionar_ativo", ativo={"igual_a": "TR1*", "sufixo": "a b"}), GRUPOS))
    assert any("qtd" in e for e in erros_de(acao("adicionar_ativo", ativo={"igual_a": "TR1*"}, qtd={"soma": "TR1*"}), GRUPOS))
    assert any("qtd" in e for e in erros_de(acao("adicionar_ativo", ativo={"igual_a": "TR1*"}, qtd=0), GRUPOS))
    assert any("@NAOEXISTE" in e for e in erros_de(acao("adicionar_ativo", ativo={"igual_a": "@NAOEXISTE"}), GRUPOS))


def test_adicionar_ativo_descreve_ativo_dinamico_em_portugues():
    assert descrever_acao(acao("adicionar_ativo", ativo={"igual_a": "TR1*", "sufixo": "VP"}, qtd=-1)) == \
        "Em Outros: para cada TR1* encontrado na linha, adicionar código encontrado + 'VP', com a mesma quantidade negativa à linha."
    assert descrever_acao(acao("adicionar_ativo", ativo={"igual_a": "TR1*"})) == \
        "Em Outros: para cada TR1* encontrado na linha, adicionar o próprio código encontrado, com a mesma quantidade à linha."


def test_remover_ativo_por_seletor_e_nao_deixa_espacos_sobrando():
    r = idempotente([acao("remover_ativo", ativo="@SUPLS")], [o("DT11/300 1-SUPL 1-CFU"), o("1-SUPL")], grupos=GRUPOS)
    assert textos(r) == ["DT11/300 1-CFU", ""]


def test_mesclar_duplicadas_soma_quantidades_e_ignora_nao_numericas():
    r = idempotente([acao("mesclar_duplicadas")], [o("DT11/300 1-CFU 2-SUPL 1-supl 3-CFU"), o("1-A 1-B"), o("1-X 1-X")])
    assert textos(r) == ["DT11/300 4-CFU 3-SUPL", "1-A 1-B", "2-X"]


# ── camada 1 como guarda ─────────────────────────────────────────────────────
def test_ajuste_que_gera_erro_de_contrato_e_descartado():
    # trocar o ativo por texto com espaço deixa um token sem par (C1-OUT-ORFAO): a linha fica como estava
    r = rodar([acao("substituir", de="CFU", para="CFU EXTRA", palavra_inteira=True)], [o("DT11/300 1-CFU")])
    assert textos(r) == ["DT11/300 1-CFU"] and r["operacoes"] == []
    assert r["descartadas"][0]["acao"] == 0 and r["descartadas"][0]["linha_id"] == "OUTROS-0" and "C1-OUT-ORFAO" in r["descartadas"][0]["motivo"]


def test_operacao_nova_invalida_pela_camada_1_e_barrada_na_validacao_e_linha_ja_invalida_pode_ser_consertada():
    assert any("operacao_nova" in e for e in validar_acoes([acao("substituir", de="A", para="B", operacao_nova="X")]))
    # linha com erro prévio (operação XX) pode ser editada: o erro não é novo
    r = rodar([acao("substituir", de="A", para="B")], [o("1-A", "XX")])
    assert textos(r) == ["1-B"]


def test_linha_nova_invalida_e_descartada():
    r = rodar([acao("adicionar_linha", valores={"operacao": "I", "ativo": "ABC"})], [o("1-A")])     # 'ABC' = token sem par
    assert textos(r) == ["1-A"] and "C1-OUT-ORFAO" in r["descartadas"][0]["motivo"]


# ── ações em sequência, imutabilidade e ids ──────────────────────────────────
def test_acoes_em_sequencia_cada_uma_ve_o_resultado_da_anterior():
    r = idempotente([acao("substituir", de="SUP-L", para="SUPL"), acao("mesclar_duplicadas"),
                     acao("ordenar", por=[{"coluna": "ativo"}]), acao("excluir_linhas", onde={"vazias": True})],
                    [o("1-SUPL 1-SUP-L"), o(""), o("1-A")])
    assert textos(r) == ["1-A", "2-SUPL"]
    assert [x["op"] for x in r["operacoes"]].count("excluir") == 1


def test_nao_altera_a_entrada_e_preserva_campos_extras():
    entrada = [o("1-SUP-L", id="meu-id", qtdAtivos=7, entidade="POSTE")]
    copia = copy.deepcopy(entrada)
    r = rodar([acao("substituir", de="SUP-L", para="SUPL")], entrada)
    assert entrada == copia
    assert r["outros"][0]["id"] == "meu-id" and r["outros"][0]["qtdAtivos"] == 7 and r["outros"][0]["entidade"] == "POSTE"
    assert r["operacoes"][0]["linha_id"] == "meu-id"


def test_sem_acoes_ou_sem_mudanca_nao_ha_operacoes():
    assert rodar([], [o("1-A")])["resumo"] == {"editar": 0, "inserir": 0, "excluir": 0, "mover": 0}


# ── validação do schema ──────────────────────────────────────────────────────
def erros_de(a, grupos=None):
    return validar_acoes([a], grupos)


@pytest.mark.parametrize("a,trecho", [
    ({"acao": "fazer", "tabela": "outros"}, "'acao' precisa ser"),
    ({"acao": "normalizar"}, "'tabela' precisa ser"),
    ({"acao": "normalizar", "tabela": "x"}, "'tabela' precisa ser"),
    ({"acao": "adicionar_ativo", "tabela": "cabos", "ativo": "A"}, "'tabela' precisa ser outros"),
    ({"acao": "substituir", "tabela": "cabos", "modo": "item", "de": "A", "para": "B"}, "'tabela' precisa ser outros"),
    ({"acao": "adicionar_linha", "tabela": "ambos", "valores": {"operacao": "I", "ativo": "1-A"}}, "'tabela' precisa ser"),
    ({"acao": "normalizar", "tabela": "outros", "foo": 1}, "campo desconhecido"),
    ({"acao": "ordenar", "tabela": "outros", "por": [{"coluna": "ativo"}], "operacoes": ["I"]}, "não permitido"),
    ({"acao": "normalizar", "tabela": "outros", "operacoes": ["Z"]}, "'operacoes'"),
    ({"acao": "normalizar", "tabela": "outros", "operacoes": []}, "'operacoes'"),
    ({"acao": "normalizar", "tabela": "outros", "regras": ["x"]}, "'regras'"),
    ({"acao": "substituir", "tabela": "outros", "de": "", "para": "B"}, "texto obrigatório"),
    ({"acao": "substituir", "tabela": "outros", "de": "A"}, "'para' precisa ser um texto"),
    ({"acao": "substituir", "tabela": "outros", "de": {"regex": "(["}, "para": "B"}, "regex inválida"),
    ({"acao": "substituir", "tabela": "outros", "de": {"regex": "A"}, "para": r"\1"}, "referência inválida"),
    ({"acao": "substituir", "tabela": "outros", "de": {"x": 1}, "para": "B"}, "regex"),
    ({"acao": "substituir", "tabela": "outros", "modo": "item", "de": "@NADA", "para": "B"}, "não existe"),
    ({"acao": "substituir", "tabela": "outros", "modo": "item", "de": "A", "para": "B C"}, "sem espaços"),
    ({"acao": "substituir", "tabela": "outros", "modo": "xx", "de": "A", "para": "B"}, "'modo'"),
    ({"acao": "substituir", "tabela": "outros", "de": "A", "para": "B", "palavra_inteira": "sim"}, "palavra_inteira"),
    ({"acao": "substituir", "tabela": "cabos", "de": "A", "para": "B", "quando": {"tem": "A"}}, "só vale a condição 'texto'"),
    ({"acao": "substituir", "tabela": "outros", "de": "A", "para": "B", "quando": {"foo": 1}}, "condição inválida"),
    ({"acao": "ordenar", "tabela": "outros", "por": []}, "'por'"),
    ({"acao": "ordenar", "tabela": "outros", "por": [{"coluna": "cor"}]}, "por[0]"),
    ({"acao": "ordenar", "tabela": "outros", "por": [{"coluna": "ativo", "ordem": "x"}]}, "ordem"),
    ({"acao": "ordenar", "tabela": "outros", "por": [{"coluna": "ativo", "valores": []}]}, "valores"),
    ({"acao": "excluir_linhas", "tabela": "outros"}, "exatamente um"),
    ({"acao": "excluir_linhas", "tabela": "outros", "onde": {"vazias": True, "texto": "x"}}, "exatamente um"),
    ({"acao": "excluir_linhas", "tabela": "outros", "onde": {"vazias": False}}, "use true"),
    ({"acao": "excluir_linhas", "tabela": "outros", "onde": {"texto": "(["}}, "regex inválida"),
    ({"acao": "adicionar_linha", "tabela": "outros", "valores": {"operacao": "Z", "ativo": "1-A"}}, "valores.operacao"),
    ({"acao": "adicionar_linha", "tabela": "outros", "valores": {"operacao": "I", "ativo": " "}}, "valores.ativo"),
    ({"acao": "adicionar_linha", "tabela": "outros", "valores": {"operacao": "I", "ativo": "1-A"}, "posicao": "meio"}, "posicao"),
    ({"acao": "adicionar_linha", "tabela": "outros", "valores": {"operacao": "I", "ativo": "1-A"}, "posicao": {"depois_de": {"foo": 1}}}, "condição inválida"),
    ({"acao": "adicionar_linha", "tabela": "outros", "valores": {"operacao": "I", "ativo": "1-A"}, "apenas_se_nao_existir": 1}, "apenas_se_nao_existir"),
    ({"acao": "adicionar_ativo", "tabela": "outros", "ativo": "A B"}, "sem espaços"),
    ({"acao": "adicionar_ativo", "tabela": "outros", "ativo": "A", "qtd": 0}, "'qtd'"),
    ({"acao": "adicionar_ativo", "tabela": "outros", "ativo": "A", "se_ja_existe": "x"}, "se_ja_existe"),
    ({"acao": "remover_ativo", "tabela": "outros", "ativo": []}, "lista de seletores vazia"),
])
def test_schema_rejeita(a, trecho):
    assert any(trecho in e for e in erros_de(a, GRUPOS)), erros_de(a, GRUPOS)


def test_schema_aceita_o_exemplo_completo_e_rejeita_o_que_nao_e_lista():
    ok = [acao("substituir", de="SUP-L", para="SUPL", operacoes=["I"], quando={"tem": "CFU"}),
          acao("normalizar", "ambos", regras=["espacos"]), acao("ordenar", por=[{"coluna": "ativo", "ordem": "desc"}]),
          acao("excluir_linhas", "cabos", onde={"texto": "^$"}), acao("adicionar_linha", "cabos", valores={"operacao": "I", "ativo": "CAA 2 A 1 m"}),
          acao("adicionar_ativo", ativo="SUPL"), acao("remover_ativo", ativo=["A", {"regex": "^B"}]), acao("mesclar_duplicadas")]
    assert validar_acoes(ok, GRUPOS) == []
    assert validar_acoes({"a": 1}) == ["O payload de ações precisa ser uma lista."]
    assert validar_acoes(["x"]) == ["Ação #1: precisa ser um objeto."]
    assert any("Ações demais" in e for e in validar_acoes([acao("normalizar")] * 51))
    with pytest.raises(ValueError):
        ajustar([{"acao": "nada"}], [], [])
    with pytest.raises(ValueError):
        ajustar([], [], [o("1-A")] * 5001)


# ── descrição em português ───────────────────────────────────────────────────
def test_descrever_acao():
    assert descrever_acao(acao("substituir", de="SUP-L", para="SUPL")) == "Em Outros: substituir 'SUP-L' por 'SUPL'."
    d = descrever_acao(acao("adicionar_ativo", ativo="SUPL", quando={"tem": "CFU", "qtd": {"=": 1}}, operacoes=["I"], operacao_nova="*I"))
    assert d == "Em Outros: adicionar 1-SUPL à linha (só nas linhas com operação I e em que a linha tem CFU com quantidade = 1) e mudar a operação para *I."
    assert descrever_acao(acao("ordenar", "ambos", por=[{"coluna": "operacao"}, {"coluna": "ativo", "ordem": "desc"}])) == \
        "Em Cabos e Outros: ordenar por operacao crescente, ativo decrescente."
    assert descrever_acao(acao("excluir_linhas", onde={"duplicadas": True})) == "Em Outros: excluir linhas duplicadas (mantém a primeira)."
    assert "no fim, só se ainda não existir" in descrever_acao(acao("adicionar_linha", valores={"operacao": "I", "ativo": "1-RA2"}))
    assert "depois de cada linha em que a linha tem CFU" in descrever_acao(
        acao("adicionar_linha", valores={"operacao": "I", "ativo": "1-A"}, posicao={"depois_de": {"tem": "CFU"}}))
    assert descrever_acao(acao("remover_ativo", ativo="@SUPLS")).startswith("Em Outros: remover um item do grupo SUPLS da linha")
    assert "somar as quantidades" in descrever_acao(acao("mesclar_duplicadas"))
    assert "normalizar (espaços, maiúsculas, poste sem '1-' no início)" in descrever_acao(acao("normalizar", regras=["espacos", "maiusculas", "poste"]))
    assert descrever_acao(acao("excluir_linhas", onde={"texto": "^X"})) == "Em Outros: excluir linhas cujo texto casa com /^X/."


# ── rotas ────────────────────────────────────────────────────────────────────
@pytest.fixture
def client(tmp_path, monkeypatch):
    monkeypatch.setattr(database, "DB_PATH", str(tmp_path / "teste.db"))
    database.init_db()
    return TestClient(appmod.app)


def cab(role="operador"):
    return {"Authorization": f"Bearer {create_jwt_token('u1', f'{role}@x.com', role)}"}


def test_rota_preview_devolve_o_diff_e_exige_login(client):
    corpo = {"acoes": [acao("substituir", de="SUP-L", para="SUPL")], "outros": [o("DT11/300 1-SUP-L")]}
    assert client.post("/api/validacao/ajustes/preview", json=corpo).status_code == 401
    r = client.post("/api/validacao/ajustes/preview", json=corpo, headers=cab()).json()
    assert r["resumo"]["editar"] == 1 and r["outros"][0]["ativo"] == "DT11/300 1-SUPL"


def test_rota_preview_usa_os_grupos_do_projeto_e_recusa_schema_invalido(client):
    corpo = {"acoes": [acao("remover_ativo", ativo="@CHAVE_MT")], "outros": [o("DT11/300 1-CFU")], "projeto_codigo": "229"}
    assert client.post("/api/validacao/ajustes/preview", json=corpo, headers=cab()).status_code == 200      # grupo da semente
    ruim = {**corpo, "acoes": [acao("remover_ativo", ativo="@NAO_EXISTE")]}
    r = client.post("/api/validacao/ajustes/preview", json=ruim, headers=cab())
    assert r.status_code == 400 and any("não existe" in e for e in r.json()["detail"]["erros"])
    grande = {"acoes": [], "outros": [o("1-A")] * 5001}
    assert client.post("/api/validacao/ajustes/preview", json=grande, headers=cab()).status_code == 400


def test_rota_descrever(client):
    r = client.post("/api/validacao/ajustes/descrever", json={"acoes": [acao("mesclar_duplicadas")]}, headers=cab()).json()
    assert r["erros"] == [] and r["frases"][0].startswith("Em Outros: somar as quantidades")
    r = client.post("/api/validacao/ajustes/descrever", json={"acoes": [{"acao": "x"}]}, headers=cab()).json()
    assert r["frases"] == [] and r["erros"]
