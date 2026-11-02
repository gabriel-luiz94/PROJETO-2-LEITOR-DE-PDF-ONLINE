"""TASK-105 — formato antigo de Outros: o que o cálculo descarta/assume ("1-U3-TM") aparece em nao_encontrados em vez de sumir em silêncio."""
from services import anotacao_outros as ao
from services import orcamento_calc as calc

BASE = [{"ativo": "U3", "desc_ativo": "ESTRUTURA U3", "componente": "CU3", "projeto": "", "mdo": "MATERIAL", "codigo": "c1", "desc_codigo": "d",
         "fator_i": 1, "fator_r": 0, "filtro": "", "origem": "OUTROS", "lv": "A"}]


def calcula(texto):
    return calc.processar_calculo([], [{"ativo": texto, "operacao": "I"}], "", BASE)


def test_hifen_dentro_do_ativo_avisa_e_nao_some():
    r = calcula("1-U3-TM")
    assert len(r["nao_encontrados"]) == 1 and "U3-TM" in r["nao_encontrados"][0] and "TM" in r["nao_encontrados"][0]
    assert len(r["resultado"]) == 1                                     # o que se entende (U3) continua calculado


def test_token_sem_par_e_quantidade_nao_numerica_avisam():
    assert "sem par" in calcula("1-U3 TM")["nao_encontrados"][0]
    assert any("não numérica" in m for m in calcula("1-U3 X-U3")["nao_encontrados"])


def test_linhas_validas_nao_geram_aviso():
    for t in ["1-U3", "2-U3 1-U3", "DT11/300 2-U3", "*2-U3", "2,5-U3", "2-U3(XLPE,DT-A)"]:
        assert not [m for m in calcula(t)["nao_encontrados"] if "descart" in m or "sem par" in m or "não suportado" in m or "não numérica" in m], t


def test_compacto_com_hifen_no_nome_segue_valendo():
    r = ao.interpretar_linha("U3-TM, 2 N1", "I", calc._classico_outros, None)
    assert r["erros"] == [] and [(i["ativo"], i["qtd"]) for i in r["itens"]] == [("U3-TM", 1.0), ("N1", 2.0)]
    assert calcula("U3-TM, 2 U3")["nao_encontrados"] == ["U3-TM"]       # compacto: nome inteiro, avisa que não existe


def test_erro_de_formato_aparece_no_calculo():
    r = calcula("U3,N1")                                                  # vírgula sem espaço
    assert r["resultado"] == [] and "vírgula" in r["nao_encontrados"][0]
