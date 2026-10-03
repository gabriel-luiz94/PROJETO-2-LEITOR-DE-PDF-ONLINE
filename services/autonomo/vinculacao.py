"""
services/autonomo/vinculacao.py — Vínculo cabo<->estrutura/poste no pipeline autônomo (TASK-036).

PORTE FIEL do bloco "VINCULAÇÃO CABO<->ESTRUTURA/POSTE" de `static/resumo.js` (TASK-032/033/034) —
só a parte AUTOMÁTICA por coordenada: não existe vinculação manual fora da tela (não há usuário no
momento do processamento autônomo). Decisão do usuário (TASK-036): quando o auto-link não encontra
combinação válida para um cabo, o pipeline não bloqueia nem vira pendência de confirmação — vira só
um achado de aviso no relatório (`avaliar_vinculacao`, igual a um ativo não encontrado na base), e o
cabo segue sem `vinculoEstruturas`.

Fidelidade ao JavaScript: `\\d` só ASCII (`tipo_estrutura`), mesmo motivo já documentado em
`montagem.py`.
"""
import re

from services.autonomo.montagem import prefixo_cabo

_RE_PRIMEIRO_DIGITO = re.compile(r"[0-9]")
_RE_RETCA = re.compile(r"\bRETCA\b", re.I)
# O ativo de uma linha Outros é "<qtd>-<código>" (ex.: "1-U4"), nunca o código puro — isola o código
# (depois do prefixo de quantidade) antes de procurar o dígito do tipo, senão pega o dígito da
# QUANTIDADE em vez do tipo da estrutura. Mesmo bug corrigido em static/resumo.js (TASK-036).
_RE_QTD_PREFIXO = re.compile(r"^[*\-]?[0-9]+(?:[.,][0-9]+)?[Xx\-](.+)\Z")


def estruturas_disponiveis(outros):
    """[(i, r), ...] — mesmo filtro de _estruturasDisponiveisVinculo (resumo.js)."""
    return [(i, r) for i, r in enumerate(outros) if r and r.get("entidade") in ("ESTRUTURA", "POSTE")]


def qtd_estruturas_para_operacao(operacao):
    op = (operacao or "").upper()
    return 1 if op in ("M", "*M") else 2


def codigo_estrutura(ativo_estrutura):
    """Código da estrutura sem o prefixo de quantidade de Outros (ex.: "U4" de "1-U4")."""
    texto = (ativo_estrutura or "").strip()
    sem_qtd = _RE_QTD_PREFIXO.match(texto)
    return sem_qtd.group(1) if sem_qtd else texto


def tipo_estrutura(ativo_estrutura):
    m = _RE_PRIMEIRO_DIGITO.search(codigo_estrutura(ativo_estrutura))
    return m.group(0) if m else None


def regra_para_tipo(regras, tipo):
    return next((r for r in (regras or []) if r.get("tipo_estrutura") == tipo), None)


def _ocupacao_atual(cabos, estruturas):
    ocupacao = {i: [] for i, _ in estruturas}
    for idx_cabo, cabo in enumerate(cabos):
        if not cabo:
            continue
        for idx_est in (cabo.get("vinculoEstruturas") or []):
            if idx_est in ocupacao:
                ocupacao[idx_est].append(idx_cabo)
    return ocupacao


def tentar_vincular_automaticamente(cabos, outros, regras):
    """Vincula cada cabo sem vínculo às estruturas mais próximas por coordenada (`_x`/`_y`),
    respeitando quantidade/compatibilidade da regra do tipo. Nunca sobrescreve vínculo já existente;
    nunca "adivinha" quando não há combinação válida — o cabo fica sem vínculo. Altera `cabos` NO
    LUGAR (campo `vinculoEstruturas`: lista de índices em `outros`). Devolve um resumo das contagens
    (mesmo texto que `window.tentarVincularAutomaticamente` mostra no modal)."""
    estruturas = estruturas_disponiveis(outros)
    resumo = {"vinculados": 0, "ja_tinham": 0, "sem_candidato": 0, "sem_coordenada": 0}
    if not estruturas:
        return resumo

    ocupacao = _ocupacao_atual(cabos, estruturas)

    for idx_cabo, cabo in enumerate(cabos):
        if not cabo:
            continue
        if cabo.get("vinculoEstruturas"):
            resumo["ja_tinham"] += 1
            continue
        if cabo.get("_x") is None or cabo.get("_y") is None:
            resumo["sem_coordenada"] += 1
            continue

        qtd_necessaria = qtd_estruturas_para_operacao(cabo.get("operacao"))
        prefixo = prefixo_cabo(cabo.get("ativo"))

        candidatas = sorted(
            ({"i": i, "r": r, "dist": ((r["_x"] - cabo["_x"]) ** 2 + (r["_y"] - cabo["_y"]) ** 2) ** 0.5}
             for i, r in estruturas if r.get("_x") is not None and r.get("_y") is not None),
            key=lambda c: c["dist"],
        )

        escolhidas = []
        for cand in candidatas:
            tipo = tipo_estrutura(cand["r"].get("ativo"))
            regra = regra_para_tipo(regras, tipo)
            capacidade = regra["qtd_cabos"] if regra else qtd_necessaria
            ocupados = ocupacao.get(cand["i"], [])
            if len(ocupados) + 1 > capacidade:
                continue
            if regra and regra.get("compatibilidade") == "MESMO_TIPO_FASE_OPERACAO" and ocupados:
                outro_cabo = cabos[ocupados[0]]
                prefixo_outro = prefixo_cabo(outro_cabo.get("ativo"))
                if prefixo_outro != prefixo or (outro_cabo.get("operacao") or "").upper() != (cabo.get("operacao") or "").upper():
                    continue
            escolhidas.append(cand)
            if len(escolhidas) == qtd_necessaria:
                break

        if len(escolhidas) == qtd_necessaria:
            cabo["vinculoEstruturas"] = [e["i"] for e in escolhidas]
            for e in escolhidas:
                ocupacao[e["i"]].append(idx_cabo)
            resumo["vinculados"] += 1
        else:
            resumo["sem_candidato"] += 1

    return resumo


def avaliar_vinculacao(cabos, outros, regras):
    """Achados (severidade "aviso", nunca bloqueia) de estrutura sem o vínculo certo pro seu tipo —
    mesmo formato de `services/validacao_planilhas.py`, porte fiel de `avaliarVinculacaoLocal`
    (TASK-033), incluindo a exceção de cabo M/`*M`/RETCA (só para zero vínculos)."""
    achados = []
    estruturas = estruturas_disponiveis(outros)
    if not estruturas:
        return achados

    ocupacao = _ocupacao_atual(cabos, estruturas)

    existe_cabo_mantendo = any((c.get("operacao") or "").upper() in ("M", "*M") for c in cabos if c)
    existe_retca = any(_RE_RETCA.search(r.get("ativo") or "") for r in outros if r)
    tem_excecao_zero = existe_cabo_mantendo or existe_retca

    for idx_est, estrutura in estruturas:
        ocupados = ocupacao.get(idx_est, [])
        tipo = tipo_estrutura(estrutura.get("ativo"))
        regra = regra_para_tipo(regras, tipo)
        if not regra:
            continue  # sem regra cadastrada pro tipo: não dá pra avaliar

        if not ocupados:
            if tem_excecao_zero:
                continue
            achados.append({
                "linha_id": f"OUTROS-{idx_est}", "tabela": "outros", "regra_id": "VINCULO-ESTRUTURA", "severidade": "aviso",
                "mensagem": f"Estrutura \"{estrutura.get('ativo')}\" sem nenhum cabo vinculado.",
                "explicacao": (f"Tipo {tipo} exige {regra['qtd_cabos']} cabo(s) vinculado(s); nenhum encontrado, "
                               "e não há cabo M/*M nem RETCA no projeto para justificar."),
            })
            continue

        if len(ocupados) != regra["qtd_cabos"]:
            achados.append({
                "linha_id": f"OUTROS-{idx_est}", "tabela": "outros", "regra_id": "VINCULO-ESTRUTURA", "severidade": "aviso",
                "mensagem": f"Estrutura \"{estrutura.get('ativo')}\" com {len(ocupados)} cabo(s) vinculado(s), esperado {regra['qtd_cabos']}.",
                "explicacao": f"Tipo {tipo} exige exatamente {regra['qtd_cabos']} cabo(s) vinculado(s).",
            })
            continue

        if regra.get("compatibilidade") == "MESMO_TIPO_FASE_OPERACAO" and len(ocupados) > 1:
            primeiro = cabos[ocupados[0]]
            prefixo_ref = prefixo_cabo(primeiro.get("ativo"))
            op_ref = (primeiro.get("operacao") or "").upper()
            incompat = any(
                prefixo_cabo(cabos[idx].get("ativo")) != prefixo_ref or (cabos[idx].get("operacao") or "").upper() != op_ref
                for idx in ocupados[1:]
            )
            if incompat:
                achados.append({
                    "linha_id": f"OUTROS-{idx_est}", "tabela": "outros", "regra_id": "VINCULO-ESTRUTURA", "severidade": "aviso",
                    "mensagem": f"Estrutura \"{estrutura.get('ativo')}\" com cabos vinculados incompatíveis entre si.",
                    "explicacao": f"Tipo {tipo} exige que os cabos vinculados sejam do mesmo tipo, fase e operação.",
                })

    return achados


def itens_vinculo_para_totalizadora(cabos, outros):
    """Pseudo-itens origem "VINCULO" (um por aresta cabo<->estrutura), para as Regras de Conversão
    casarem contra o ativo composto (ex.: `CAA2_N4`) — porte fiel do bloco em `syncTotalizadora`
    (TASK-034). Só existe aqui dentro: nunca grava nas tabelas Cabos/Outros."""
    itens = []
    for idx_cabo, cabo in enumerate(cabos):
        if not cabo or not cabo.get("vinculoEstruturas"):
            continue
        prefixo = prefixo_cabo(cabo.get("ativo"))
        if not prefixo:
            continue
        for idx_estrutura in cabo["vinculoEstruturas"]:
            if not (0 <= idx_estrutura < len(outros)):
                continue
            estrutura = outros[idx_estrutura]
            if not estrutura or not estrutura.get("ativo"):
                continue
            itens.append({
                "baseId": f"TOT-VINC-{idx_cabo}-{idx_estrutura}", "obs": "VINCULO",
                "operacao": cabo.get("operacao") or "I",
                "ativo": f"{prefixo}_{codigo_estrutura(estrutura['ativo']).upper()}",
                "qtd": 1.0, "desc": "", "naoEncontrado": False, "origem": "VINCULO",
            })
    return itens
