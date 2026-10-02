"""
services/autonomo/aplicar.py — Aplica no backend as operações do diff de ajustes (TASK-031, fase B).

PORTE FIEL de `aplicarOperacoesAjuste` (static/resumo.js): editar → excluir → inserir → reordenar, por tabela; mudança cuja linha não
confere mais com o "antes" é ignorada; `autoClassifyEntidade` classifica linhas novas/editadas; Cabos recalcula `qtdAtivos`.
A paridade é provada em tests/test_autonomo_aplicar.py executando o JS real. O backend recebe sempre as tabelas ORIGINAIS e o diff:
reaplicar do zero (em vez de acumular) evita que os ids posicionais (`OUTROS-3`) se desencontrem depois de inserções.
"""
import copy

from services.autonomo.leitor_js import obter_leitor
from services.autonomo.montagem import recalcular_qtd_ativos

PREFIXO_TABELA = {"cabos": "CABOS", "outros": "OUTROS"}


def _trim(s):
    return (s or "").strip()


def aplicar_operacoes(cabos: list, outros: list, operacoes: list, escolhidas=None, regras_cls=None, leitor=None) -> dict:
    """`escolhidas` = índices de `operacoes` a aplicar (None = todas). Devolve {cabos, outros, aplicadas, ignoradas, ignoradas_detalhe}."""
    leitor = leitor or obter_leitor()
    tabelas = {"cabos": copy.deepcopy(cabos or []), "outros": copy.deepcopy(outros or [])}
    ops = [o for i, o in enumerate(operacoes) if escolhidas is None or i in escolhidas]
    aplicadas = ignoradas = 0
    detalhe = []
    tocadas = set()

    def classifica(ativo):
        return leitor.classificar_entidade(ativo, regras_cls or [])

    for tabela in ("cabos", "outros"):
        do_tabela = [o for o in ops if o["tabela"] == tabela]
        if not do_tabela:
            continue
        lista = [{"id": f"{PREFIXO_TABELA[tabela]}-{i}", "row": row} for i, row in enumerate(tabelas[tabela]) if row]

        def achar(linha_id):
            return next((x for x in lista if x["id"] == linha_id), None)

        def confere(row, antes):
            return _trim(row.get("ativo")) == _trim(antes.get("ativo")) and (row.get("operacao") or "") == antes.get("operacao")

        for o in (o for o in do_tabela if o["op"] == "editar"):
            x = achar(o["linha_id"])
            if not x or not confere(x["row"], o["antes"]):
                ignoradas += 1
                detalhe.append(o["linha_id"])
                continue
            x["row"]["ativo"] = o["depois"]["ativo"]
            x["row"]["operacao"] = o["depois"]["operacao"]
            if x["row"].get("entidade") == "0":
                x["row"]["entidade"] = classifica(o["depois"]["ativo"])
            aplicadas += 1
            tocadas.add(tabela)
        for o in (o for o in do_tabela if o["op"] == "excluir"):
            x = achar(o["linha_id"])
            if not x or not confere(x["row"], o["antes"]):
                ignoradas += 1
                detalhe.append(o["linha_id"])
                continue
            lista = [y for y in lista if y is not x]
            aplicadas += 1
            tocadas.add(tabela)
        for o in (o for o in do_tabela if o["op"] == "inserir"):
            ent = classifica(o["depois"]["ativo"])
            entidade = ent if ent != "0" else ("CABO" if tabela == "cabos" else (o["depois"].get("entidade") or "0"))
            nova = {"entidade": entidade, "operacao": o["depois"].get("operacao") or "I", "ativo": o["depois"]["ativo"]}
            ref = next((i for i, y in enumerate(lista) if y["id"] == o.get("depois_de")), -1) if o.get("depois_de") else -1
            if o.get("depois_de") and ref >= 0:
                pos = ref + 1
            else:
                pos = len(lista) if o.get("depois_de") else 0
            lista.insert(pos, {"id": o["linha_id"], "row": nova})
            aplicadas += 1
            tocadas.add(tabela)
        mover = next((o for o in do_tabela if o["op"] == "mover"), None)
        if mover:
            pos = {id_: i for i, id_ in enumerate(mover["ordem_depois"])}
            ultima = -1
            chaves = []
            for x in lista:
                if x["id"] in pos:
                    ultima = pos[x["id"]]
                chaves.append(ultima + (0 if x["id"] in pos else 0.5))
            ordem = sorted(range(len(lista)), key=lambda i: (chaves[i], i))
            lista = [lista[i] for i in ordem]
            aplicadas += 1
            tocadas.add(tabela)
        tabelas[tabela] = [x["row"] for x in lista]

    if "cabos" in tocadas:
        recalcular_qtd_ativos(tabelas["cabos"])
    return {"cabos": tabelas["cabos"], "outros": tabelas["outros"], "aplicadas": aplicadas, "ignoradas": ignoradas, "ignoradas_detalhe": detalhe}
