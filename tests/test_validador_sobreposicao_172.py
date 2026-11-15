"""TASK-172 — validador da base: menos alarme falso em "grupos que podem valer juntos" (padrões %X% e aviso resumo por ativo)."""
import pytest

from services import ativo_parametros as ap


def L(ativo, comp="", cod="", proj="PERNAMBUCO"):
    return {"ativo": ativo, "componente": comp, "codigo": cod, "projeto": proj, "desc_ativo": "d", "mdo": "MATERIAL", "desc_codigo": "d", "fator_i": 1, "fator_r": 0,
            "filtro": "", "origem": "OUTROS", "lv": "A", "arred": "NORMAL", "formula_i": "", "formula_r": ""}


@pytest.mark.parametrize("a,b,esperado", [
    ("%30KVA%", "%10KVA%", False), ("%A35%", "%S1/0%", False), ("%S2/0%", "%W2/0%", False),
    ("%A35%", "%A35P%", True), ("%A%", "%A185P%", True), ("%a35%", "%A35P%", True),          # sem diferenciar maiúsculas
    ("C%", "%S04%", True), ("%A", "%C", False), ("D%", "%T", True), ("D%", "C%", False),      # os demais formatos: regra atual
])
def test_padroes_contem(a, b, esperado):
    assert ap._alts_coincidem(a, b) is esperado


def avisos_de(grupos):
    base = [L("RL=RL(cabo=A35P)")] + [L(f"RL(cabo={g})", f"RL.{i}", "2212015") for i, g in enumerate(grupos)]
    return ap.validar_base(ap.preencher_componentes(base)[0])["avisos"]


def test_trafos_diferentes_nao_avisam_mais():
    assert avisos_de(["%30KVA%", "%10KVA%"]) == []
    assert len(avisos_de(["%A35%", "%A35P%"])) == 1                                              # a duplicidade possível continua


def test_duplicidade_real_continua_avisando():
    linhas = [L("N1_TT=N1_TT(tipocabo=A)"), L("N1_TT(tipocabo=%C)", "N1_TT.C", "3428086"), L("N1_TT(tipocabo=%)", "N1_TT.Q", "3428086")]
    avisos = ap.validar_base(ap.preencher_componentes(linhas)[0])["avisos"]
    assert len(avisos) == 1 and "3428086" in avisos[0]


def _aviso(cod, g1, g2, n1=1, n2=2):
    return f"Código {cod} aparece nos grupos {g1} (linha {n1}) e {g2} (linha {n2}), que podem valer juntos: a quantidade soma duas vezes."


def test_agrupar_um_resumo_por_ativo_com_mais_de_5_e_a_lista_recolhida():
    avisos = [_aviso(f"22{i}", "RL(CABO=%A35P%)", "RL(CABO2=%A35P%)") for i in range(61)] + [_aviso("77", "TE(CABO=%A%)", "TE(CABO=%A1%)")] + ["Outro aviso qualquer."]
    mostrar, detalhe = ap.agrupar_avisos(avisos)
    assert len(mostrar) == 3 and mostrar[2] == "Outro aviso qualquer." and mostrar[1] == avisos[61]            # TE (1 aviso) e o resto passam como estão
    resumo = mostrar[0]
    assert resumo.startswith("RL: 61 códigos aparecem em grupos que podem valer juntos (ex.: 220 em RL(CABO=%A35P%) e RL(CABO2=%A35P%))") and resumo.endswith("ignore.")
    assert detalhe == {resumo: avisos[:61]} and detalhe[resumo][0].startswith("Código 220 aparece")          # o texto de cada aviso individual não muda


def test_ate_5_avisos_do_ativo_nao_agrupa():
    avisos = [_aviso(f"1{i}", "TE(CABO=%A%)", "TE(CABO=%A1%)") for i in range(5)]
    assert ap.agrupar_avisos(avisos) == (avisos, {})
    assert ap.agrupar_avisos([]) == ([], {})


def test_validar_base_continua_devolvendo_a_lista_completa():
    base = [L("RL=RL(cabo=A35P)")] + [L(f"RL(cabo=%A35{i}%)", f"RL.{i}", "2212015") for i in range(7)]
    # %A350%, %A351%… não coincidem entre si (um não contém o outro): sem avisos
    assert ap.validar_base(ap.preencher_componentes(base)[0])["avisos"] == []
    prefixos = [L("RL=RL(cabo=A35P)")] + [L(f"RL(cabo=%A35{'P' * i}%)", f"RL.{i}", "2212015") for i in range(7)]       # cada um contém o anterior: avisos de verdade
    avisos = ap.validar_base(ap.preencher_componentes(prefixos)[0])["avisos"]
    assert len(avisos) == 21                                                                         # C(7,2): a lista completa não é agrupada
    mostrar, detalhe = ap.agrupar_avisos(avisos)
    assert len(mostrar) == 1 and mostrar[0].startswith("RL: 21 códigos") and len(list(detalhe.values())[0]) == 21


def test_respostas_da_base_usam_o_agrupamento(monkeypatch):
    import inspect
    from routers import admin
    from services import rascunho_ativo
    assert "agrupar_avisos" in inspect.getsource(admin) and "agrupar_avisos" in inspect.getsource(rascunho_ativo)
