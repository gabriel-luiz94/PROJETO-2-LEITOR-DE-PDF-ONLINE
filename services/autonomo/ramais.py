"""
services/autonomo/ramais.py — Ativos gerados pelos RAMAIS (TASK-031, fase B).

PORTE FIEL de `recalcularAtivosRamais` + `adicionarAtivosRamais` (static/resumo.js, modal RAMAIS): lê os textos "TROCAR n RS M AC" etc. dos itens
RAMAIS/IP/APOIO, soma os ativos a instalar e a remover e os acrescenta na tabela Outros (entidade '0', texto 'RAMAIS (GERADO)'). Decisão do usuário
(2026-10-02): no modo autônomo os ramais ENTRAM no orçamento. Paridade provada em tests/test_autonomo_ramais.py contra o JS real.
"""
import re

from services.autonomo.montagem import _S

_RE_TROCAR_N = re.compile(r"TROCAR" + _S + r"+([0-9]+)" + _S + r"+RS", re.I)
_RE_TROCAR = re.compile(r"TROCAR" + _S + r"+RS", re.I)

# (trecho do texto em maiúsculas, ativo instalado, fator, ativo removido, fator) — na MESMA ordem do `if / else if` do JS
_CASOS = [
    (("RS M AC", "RS MAC"), "MAC", 20, "MAC", 15),
    (("RS M AM", "RS MAM"), "MAM", 20, "MAM", 15),
    (("RS T AM", "RS TAM"), "TAM", 20, "TAM", 15),
    (("RS M AA", "RS MAA"), "MAM", 20, "MAA", 1),
    (("RS T AA", "RS TAA"), "TAM", 20, "MAA", 2),
]


def calcular_ramais(textos: list) -> tuple:
    """(instalando, removendo): dicionários {ativo: quantidade}."""
    inst, rem = {}, {}
    for texto in textos:
        texto = (texto or "").strip(" \t\n\v\f\r ﻿")
        m = _RE_TROCAR_N.search(texto)
        qtd = int(m.group(1)) if m else (1 if _RE_TROCAR.search(texto) else 0)
        if qtd <= 0:
            continue
        up = texto.upper()
        for gatilhos, a_inst, f_inst, a_rem, f_rem in _CASOS:
            if any(g in up for g in gatilhos):
                inst[a_inst] = inst.get(a_inst, 0) + qtd * f_inst
                rem[a_rem] = rem.get(a_rem, 0) + qtd * f_rem
                break
        if "CP-REDE" in up or "CP REDE" in up or "CPREDE" in up:
            inst["CPREDE"] = inst.get("CPREDE", 0) + qtd
    return inst, rem


def linhas_ramais(textos: list) -> list:
    """Linhas a acrescentar em Outros: primeiro as de instalação (I), depois as de remoção (R), cada grupo na ORDEM EM QUE O ATIVO APARECEU (o JS usa `Object.keys`, sem ordenar; só o texto de resumo do modal é ordenado)."""
    inst, rem = calcular_ramais(textos)
    linhas = [{"entidade": "0", "operacao": "I", "ativo": f"{inst[k]}-{k}", "texto": "RAMAIS (GERADO)"} for k in inst]
    linhas += [{"entidade": "0", "operacao": "R", "ativo": f"{rem[k]}-{k}", "texto": "RAMAIS (GERADO)"} for k in rem]
    return linhas
