"""Ativo cadastrado com a linha de declaração SEM componente e os itens (só remoção) abaixo trazendo o componente: o cálculo não pode ficar vazio."""
from services.orcamento_calc import processar_calculo


def L(ativo, cod="", desc="", fi=0, fr=0, comp="", mdo="MATERIAL", lv="A"):
    return {"ativo": ativo, "desc_ativo": "d", "componente": comp, "projeto": "PE", "mdo": mdo, "origem": "", "codigo": cod, "desc_codigo": desc,
            "fator_i": fi, "fator_r": fr, "filtro": "F", "lv": lv, "arred": "NORMAL", "formula_i": "", "formula_r": ""}


BASE = [L("N3_2, N32, N23"),                                                    # declaração sem COMPONENTE
        L("N3_2", "SDEVU", "RETIRAR LV", 0, 1, "N3_2", "MAO-DE-OBRA", "LV"),
        L("N3_2", "7931008", "SUCATA", 0, 2, "N3_2"),
        L("N3_2", "SDEMU", "RETIRAR", 0, 1, "N3_2", "MAO-DE-OBRA", "LM")]


def calc(nome, op):
    r = processar_calculo([], [{"entidade": "0", "operacao": op, "ativo": nome, "qtdAtivos": 1, "texto": ""}], "PE", BASE)
    return {(x["operacao"], x["codigo"]): x["total"] for x in r["resultado"]}, r["nao_encontrados"]


def test_remocao_calcula_mesmo_com_declaracao_sem_componente():
    res, nao = calc("N3_2", "R")
    assert res == {("R", "SDEMU"): 1.0, ("R", "7931008"): 2.0} and nao == []
    assert calc("N32", "R")[0] == res                                           # pelo apelido também


def test_linha_viva_pega_o_item_lv():
    assert calc("N3_2", "*R")[0] == {("R", "SDEVU"): 1.0, ("R", "7931008"): 2.0}


def test_instalacao_de_ativo_so_de_remocao_continua_vazia():
    assert calc("N3_2", "I")[0] == {}
