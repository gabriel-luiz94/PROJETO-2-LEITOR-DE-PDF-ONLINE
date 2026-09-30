"""tests/referencia_regras_v1.py — CÓPIA CONGELADA do motor de regras v1 (TASK-013), usada só como ORÁCULO nos testes.

Não importar em código de produção e não "consertar" aqui: serve para provar que o motor v2 (TASK-016) produz
exatamente os mesmos achados que o v1 para as regras da semente e para regras v1 já salvas.
"""
import re
from functools import lru_cache

from services.orcamento_calc import extrair_ativos_outros
from services.validacao_planilhas import OPERACOES_VALIDAS, _para_float, _tokens_cabo_tabela

TIPOS = {"requer", "proibe", "nao_isolado", "texto", "minimo_total"}
SEVERIDADES = {"erro", "aviso", "info"}
ESCOPOS = {"outros", "cabos+outros"}
OPERACOES_PADRAO = ["I", "*I"]


@lru_cache(maxsize=512)
def _re(padrao: str):
    return re.compile(padrao, re.IGNORECASE)


# ═══════════════════════════════════════════════════════════════════════════
# VALIDAÇÃO DO SCHEMA — nunca deixa salvar uma regra que o motor não saberia executar
# ═══════════════════════════════════════════════════════════════════════════
def _regex(valor, campo, erros, obrigatorio=True):
    if valor in (None, ""):
        if obrigatorio:
            erros.append(f"{campo}: obrigatório.")
        return
    if not isinstance(valor, str):
        erros.append(f"{campo}: precisa ser texto.")
        return
    try:
        re.compile(valor)
    except re.error as e:
        erros.append(f"{campo}: regex inválida ({e}).")


def _numero(valor, campo, erros, minimo=None):
    if isinstance(valor, bool) or not isinstance(valor, (int, float)):
        erros.append(f"{campo}: precisa ser numérico.")
    elif minimo is not None and valor < minimo:
        erros.append(f"{campo}: precisa ser >= {minimo}.")


def _validar_parametros(tipo, p, prefixo, erros):
    if tipo == "requer":
        _regex(p.get("se_regex"), f"{prefixo}.se_regex", erros)
        _regex(p.get("exige_regex"), f"{prefixo}.exige_regex", erros)
        _regex(p.get("exceto_se_regex"), f"{prefixo}.exceto_se_regex", erros, obrigatorio=False)
        if "qtd_min" in p:
            _numero(p["qtd_min"], f"{prefixo}.qtd_min", erros, minimo=0)
        if "multiplicador" in p:
            _numero(p["multiplicador"], f"{prefixo}.multiplicador", erros, minimo=1)
        dobra = p.get("dobra_se")
        if dobra is not None:
            if not isinstance(dobra, list):
                erros.append(f"{prefixo}.dobra_se: precisa ser uma lista.")
            else:
                for i, cond in enumerate(dobra):
                    if not isinstance(cond, dict):
                        erros.append(f"{prefixo}.dobra_se[{i}]: precisa ser um objeto {{regex, qtd_min}}.")
                        continue
                    _regex(cond.get("regex"), f"{prefixo}.dobra_se[{i}].regex", erros)
                    _numero(cond.get("qtd_min", 1), f"{prefixo}.dobra_se[{i}].qtd_min", erros, minimo=0)
    elif tipo == "proibe":
        _regex(p.get("se_regex"), f"{prefixo}.se_regex", erros)
        _regex(p.get("com_regex"), f"{prefixo}.com_regex", erros)
        _regex(p.get("exceto_se_regex"), f"{prefixo}.exceto_se_regex", erros, obrigatorio=False)
    elif tipo == "nao_isolado":
        _regex(p.get("se_regex"), f"{prefixo}.se_regex", erros)
        _regex(p.get("acompanhado_por_regex"), f"{prefixo}.acompanhado_por_regex", erros)
        if "min_outros" in p:
            _numero(p["min_outros"], f"{prefixo}.min_outros", erros, minimo=1)
    elif tipo == "texto":
        _regex(p.get("texto_regex"), f"{prefixo}.texto_regex", erros)
    elif tipo == "minimo_total":
        _regex(p.get("ativo_regex"), f"{prefixo}.ativo_regex", erros)
        contrib = p.get("contribuicoes")
        if not isinstance(contrib, list) or not contrib:
            erros.append(f"{prefixo}.contribuicoes: precisa ser uma lista com ao menos uma contribuição.")
        else:
            for i, c in enumerate(contrib):
                if not isinstance(c, dict):
                    erros.append(f"{prefixo}.contribuicoes[{i}]: precisa ser um objeto.")
                    continue
                _regex(c.get("se_regex"), f"{prefixo}.contribuicoes[{i}].se_regex", erros)
                _numero(c.get("metros_por_unidade"), f"{prefixo}.contribuicoes[{i}].metros_por_unidade", erros, minimo=0)


def validar_regras(regras) -> list:
    """Lista de erros (vazia = válido)."""
    if not isinstance(regras, list):
        return ["O payload de regras precisa ser uma lista."]
    erros, ids = [], set()
    for idx, r in enumerate(regras):
        prefixo = f"Regra #{idx + 1}"
        if not isinstance(r, dict):
            erros.append(f"{prefixo}: precisa ser um objeto.")
            continue
        rid = r.get("id")
        if not isinstance(rid, str) or not rid.strip():
            erros.append(f"{prefixo}: 'id' é obrigatório.")
        elif rid in ids:
            erros.append(f"{prefixo}: 'id' duplicado ({rid}).")
        else:
            ids.add(rid)
            prefixo = f"Regra {rid}"
        tipo = r.get("tipo")
        if tipo not in TIPOS:
            erros.append(f"{prefixo}: 'tipo' precisa ser um de {sorted(TIPOS)} (recebido: {tipo!r}).")
        if r.get("severidade") not in SEVERIDADES:
            erros.append(f"{prefixo}: 'severidade' precisa ser um de {sorted(SEVERIDADES)}.")
        if r.get("escopo", "outros") not in ESCOPOS:
            erros.append(f"{prefixo}: 'escopo' precisa ser um de {sorted(ESCOPOS)}.")
        if tipo == "minimo_total" and r.get("escopo") != "cabos+outros":
            erros.append(f"{prefixo}: 'minimo_total' exige escopo 'cabos+outros'.")
        if not isinstance(r.get("mensagem"), str) or not r["mensagem"].strip():
            erros.append(f"{prefixo}: 'mensagem' é obrigatória.")
        if not isinstance(r.get("ativa"), bool):
            erros.append(f"{prefixo}: 'ativa' precisa ser true ou false.")
        ops = r.get("operacoes")
        if ops is not None and (not isinstance(ops, list) or not ops
                                or any(str(o).upper() not in OPERACOES_VALIDAS for o in ops)):
            erros.append(f"{prefixo}: 'operacoes' precisa ser uma lista com valores de {sorted(OPERACOES_VALIDAS)}.")
        params = r.get("parametros")
        if not isinstance(params, dict):
            erros.append(f"{prefixo}: 'parametros' precisa ser um objeto.")
        elif tipo in TIPOS:
            _validar_parametros(tipo, params, f"{prefixo}.parametros", erros)
    return erros


# ═══════════════════════════════════════════════════════════════════════════
# AVALIAÇÃO
# ═══════════════════════════════════════════════════════════════════════════
def _achado(linha_id, tabela, regra, detalhe=""):
    mensagem = regra["mensagem"] + (f" ({detalhe})" if detalhe else "")
    return {"linha_id": linha_id, "tabela": tabela, "regra_id": regra["id"],
            "severidade": regra["severidade"], "mensagem": mensagem}


def _soma(itens, padrao):
    rx = _re(padrao)
    return sum(i["qtd"] for i in itens if rx.search(i["ativo"]))


def _tem(itens, padrao):
    rx = _re(padrao)
    return any(rx.search(i["ativo"]) for i in itens)


def _ops(regra):
    return {str(o).strip().upper() for o in (regra.get("operacoes") or OPERACOES_PADRAO)}


def _avaliar_linha(regra, itens, texto):
    """Devolve o detalhe do problema (str) ou None se a linha está de acordo com a regra."""
    p, tipo = regra["parametros"], regra["tipo"]
    if tipo == "requer":
        if not _tem(itens, p["se_regex"]):
            return None
        if p.get("exceto_se_regex") and _tem(itens, p["exceto_se_regex"]):
            return None
        minimo = p.get("qtd_min", 1)
        if any(_soma(itens, c["regex"]) >= c.get("qtd_min", 1) for c in p.get("dobra_se") or []):
            minimo *= p.get("multiplicador", 2)
        tem = _soma(itens, p["exige_regex"])
        return f"encontrado {tem:g}, mínimo {minimo:g}" if tem < minimo else None
    if tipo == "proibe":
        if p.get("exceto_se_regex") and _tem(itens, p["exceto_se_regex"]):
            return None
        return "ativos incompatíveis na mesma linha" if _tem(itens, p["se_regex"]) and _tem(itens, p["com_regex"]) else None
    if tipo == "nao_isolado":
        if not _tem(itens, p["se_regex"]):
            return None
        outros = _soma(itens, p["acompanhado_por_regex"])
        if _tem([i for i in itens if _re(p["se_regex"]).search(i["ativo"])], p["acompanhado_por_regex"]):
            outros -= 1  # o próprio gatilho também casa com "acompanhado_por": desconta uma unidade dele
        minimo = p.get("min_outros", 1)
        return f"acompanhantes {outros:g}, mínimo {minimo:g}" if outros < minimo else None
    if tipo == "texto":
        return "texto do ativo fora do padrão" if _re(p["texto_regex"]).search(texto) else None
    return None


def _comprimento_bruto(cabo_ativo):
    tokens = _tokens_cabo_tabela(cabo_ativo)
    if len(tokens) < 2:
        return tokens[0] if tokens else "", 0.0
    valor = _para_float(tokens[-1])
    return tokens[0], (valor if valor is not None else 0.0)


def _avaliar_minimo_total(regra, cabos, outros):
    p, ops = regra["parametros"], _ops(regra)
    itens_outros = []
    for linha in outros or []:
        if (linha.get("operacao") or "").strip().upper() in ops:
            itens_outros += extrair_ativos_outros({"ativo": linha.get("ativo") or "", "operacao": ""})
    exigido = sum(c["metros_por_unidade"] * _soma(itens_outros, c["se_regex"]) for c in p["contribuicoes"])
    if exigido <= 0:
        return None
    rx = _re(p["ativo_regex"])
    presente = 0.0
    for linha in cabos or []:
        if (linha.get("operacao") or "").strip().upper() not in ops or not (linha.get("ativo") or "").strip():
            continue
        nome, metros = _comprimento_bruto(linha["ativo"])
        if rx.search(nome):
            presente += metros
    return f"exigido {exigido:g} m no total, presente {presente:g} m" if presente < exigido else None


def avaliar(regras: list, cabos: list, outros: list) -> list:
    """Achados da camada 2. Regras inativas são ignoradas. `regras` deve ter passado por validar_regras."""
    achados = []
    for regra in regras or []:
        if not regra.get("ativa"):
            continue
        if regra["tipo"] == "minimo_total":
            detalhe = _avaliar_minimo_total(regra, cabos, outros)
            if detalhe:
                achados.append(_achado("GERAL", "cabos+outros", regra, detalhe))
            continue
        ops = _ops(regra)
        for i, linha in enumerate(outros or []):
            if (linha.get("operacao") or "").strip().upper() not in ops:
                continue
            texto = (linha.get("ativo") or "").strip()
            if not texto:
                continue
            itens = extrair_ativos_outros({"ativo": texto, "operacao": ""})
            detalhe = _avaliar_linha(regra, itens, texto)
            if detalhe:
                achados.append(_achado(linha.get("id") or f"OUTROS-{i}", "outros", regra, detalhe))
    return achados
