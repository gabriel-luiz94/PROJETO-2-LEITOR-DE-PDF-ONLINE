"""TASK-139 — ajuste `somar_linhas`: junta, entre linhas de Outros, as quantidades dos ativos escolhidos na 1ª linha em que aparecem."""
import pytest

from services.ajustes_planilhas import ajustar, descrever_acao, validar_acoes


def L(op, ativo, ent="0"):
    return {"operacao": op, "ativo": ativo, "entidade": ent}


OUTROS = [L("I", "1-CFU 2-SUPL"), L("I", "2-CFU"), L("I", "1-CFU 1-N1"), L("*I", "*1-CFU"), L("I", "4-CFU", "5"), L("I", "1-PR"), L("I", "2-PR")]


def visivel(r):
    return [(x["operacao"], x["ativo"], x["entidade"]) for x in r["outros"]]


def test_soma_na_primeira_linha_e_remove_das_outras():
    r = ajustar([{"acao": "somar_linhas", "tabela": "outros", "ativo": "CFU"}], [], OUTROS)
    assert visivel(r) == [("I", "4-CFU 2-SUPL", "0"), ("I", "1-N1", "0"), ("*I", "*1-CFU", "0"), ("I", "4-CFU", "5"), ("I", "1-PR", "0"), ("I", "2-PR", "0")]
    assert r["resumo"]["excluir"] == 1 and r["resumo"]["editar"] == 2 and not r["descartadas"]


def test_ignorar_entidade_e_varios_ativos():
    r = ajustar([{"acao": "somar_linhas", "tabela": "outros", "ativo": ["CFU", "PR"], "por_entidade": False}], [], OUTROS)
    assert [x for x in visivel(r) if x[1] in ("8-CFU 2-SUPL", "3-PR")] == [("I", "8-CFU 2-SUPL", "0"), ("I", "3-PR", "0")]
    assert ("I", "4-CFU", "5") not in visivel(r) and ("*I", "*1-CFU", "0") in visivel(r)      # *I nunca se mistura com I


def test_so_os_selecionados_e_idempotente():
    a = [{"acao": "somar_linhas", "tabela": "outros", "ativo": "PR"}]
    r1 = ajustar(a, [], OUTROS)
    assert ("I", "3-PR", "0") in visivel(r1) and ("I", "1-N1", "0") not in visivel(r1) and ("I", "1-CFU 1-N1", "0") in visivel(r1)
    assert ajustar(a, [], r1["outros"])["operacoes"] == []


def test_soma_zero_remove_e_linha_compacta_nao_e_editada():
    r = ajustar([{"acao": "somar_linhas", "tabela": "outros", "ativo": "PR"}], [], [L("I", "2-PR"), L("I", "9400D, 1-PR"), L("I", "1-PR")])
    assert visivel(r) == [("I", "3-PR", "0"), ("I", "9400D, 1-PR", "0")]


def test_validacao_e_frase():
    assert validar_acoes([{"acao": "somar_linhas", "tabela": "outros", "ativo": "CFU", "por_entidade": False}]) == []
    assert validar_acoes([{"acao": "somar_linhas", "tabela": "cabos", "ativo": "CFU"}])
    assert validar_acoes([{"acao": "somar_linhas", "tabela": "outros", "ativo": "CFU", "por_entidade": "sim"}])
    assert "juntar em uma só linha" in descrever_acao({"acao": "somar_linhas", "tabela": "outros", "ativo": "CFU"})
