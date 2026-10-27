"""
services/anotacao_outros.py — Notação compacta na tabela Outros (TASK-081).

Uma linha de Outros pode vir em dois formatos (nunca misturados). O usuário vê o que digitou; a tradução para o que o cálculo entende
é implícita (esta função):

    ANTIGO    ``DT11/600 1-N1``              tokens ``qtd-ativo`` separados por espaço (como sempre foi)
    COMPACTO  ``11600D, 2 N3, T3``           itens ``[quantidade ]ativo`` separados por ``, `` (vírgula + espaço)
              ``2 11600D, 2 N3, T3``

  * vírgula SEM espaço depois = separador decimal (``0,5 N3``); vírgula entre não-dígitos sem espaço (``11600D,N1``) é erro.
  * poste em notação compacta: ``<altura><esforço><letra>`` — D=DT, R=RC, F=CV, FV=FV (11300D = DT11/300, 7150FV = FV7/150).
  * apelidos cadastrados na coluna ATIVO da base (``DT11/300, 11300D``) valem em qualquer item, nos dois formatos.
  * cada estrutura ganha o parâmetro ``poste`` (``N1(POSTE=11600D)``) do poste mais próximo à ESQUERDA, se a assinatura do ativo tiver
    o parâmetro ``poste``; sem poste à esquerda vale o padrão.
  * operação M/*M: o poste é ignorado e os demais itens ficam I/*I.

Este arquivo tem a MESMA lógica de static/anotacao_outros.js (paridade testada em tests/test_anotacao_outros.py).
O "classico" é o leitor do formato antigo de cada chamador (cálculo, Totalizadora…): recebe o texto e devolve [{"ativo", "qtd"}].
"""
import re

from services.ativo_parametros import (
    analisar_celula_base, apelidos_do_projeto, assinaturas_do_projeto, dividir_tokens, nome_base,
)

SUFIXO_PARA_TIPO = {"D": "DT", "R": "RC", "F": "CV", "FV": "FV"}
TIPO_PARA_SUFIXO = {v: k for k, v in SUFIXO_PARA_TIPO.items()}
ALTURA_MIN, ALTURA_MAX = 7, 16

_RE_COMPACTO = re.compile(r"^(\d+)(FV|D|R|F)$", re.I)
_RE_POSTE = re.compile(r"^(DT|RC|CV|FV)(\d+)/(\d+)$")
_RE_QTD = re.compile(r"^\*?\d+(?:[.,]\d+)?$")
_RE_QTD_ANTIGA = re.compile(r"^\*?\d+(?:[.,]\d+)?[-xX][A-Za-z(].*$")
_RE_QTD_CLASSICA = re.compile(r"^[*-]?\d+(?:[.,]\d+)?[Xx-].+$")          # token do formato antigo (qtd-ativo): decide entre o formato antigo e o compacto numa linha sem vírgula+espaço


def eh_poste(nome) -> bool:
    """O nome (sem parâmetros) é um poste: DT|RC|CV|FV + altura/esforço?"""
    return bool(_RE_POSTE.match(str(nome if nome is not None else "").strip().upper()))


def poste_para_compacto(nome):
    """DT11/600 → 11600D (None se não for poste)."""
    m = _RE_POSTE.match(str(nome if nome is not None else "").strip().upper())
    return f"{m.group(2)}{m.group(3)}{TIPO_PARA_SUFIXO[m.group(1)]}" if m else None


def compacto_para_poste(token, postes=None):
    """Notação compacta → poste. Retorna None se o token não é compacto; senão {"ativo", "compacto", "erro"}.

    Para separar altura de esforço prefere o poste que EXISTE na base (``postes`` = conjunto de nomes em maiúsculas); sem base, usa
    alturas 7–16 e esforço de 3–4 dígitos sem zero à esquerda. Ambíguo ou impossível → erro."""
    t = str(token if token is not None else "").strip()
    m = _RE_COMPACTO.match(t)
    if not m:
        return None
    digitos, suf = m.group(1), m.group(2).upper()
    tipo = SUFIXO_PARA_TIPO[suf]
    cands = []
    for hl in (2, 1):
        h, e = digitos[:hl], digitos[hl:]
        if not 3 <= len(e) <= 4 or h.startswith("0") or e.startswith("0"):
            continue
        cands.append((int(h), f"{tipo}{int(h)}/{int(e)}"))
    r = {"ativo": None, "compacto": t.upper(), "erro": None}
    if postes:
        na_base = [n for _, n in cands if n in postes]
        if len(na_base) == 1:
            r["ativo"] = na_base[0]
            return r
    plausiveis = [n for h, n in cands if ALTURA_MIN <= h <= ALTURA_MAX]
    if len(plausiveis) == 1:
        r["ativo"] = plausiveis[0]
    else:
        r["erro"] = f"não foi possível separar altura e esforço de '{t.upper()}'" if not plausiveis else f"'{t.upper()}' é ambíguo ({' ou '.join(plausiveis)})"
    return r


def _virgulas(texto):
    """Vírgulas fora de parênteses: [(posição, 'separador'|'decimal'|'invalida')]."""
    saida, prof = [], 0
    for i, ch in enumerate(texto):
        if ch == "(":
            prof += 1
        elif ch == ")" and prof > 0:
            prof -= 1
        elif ch == "," and prof == 0:
            prox = texto[i + 1] if i + 1 < len(texto) else ""
            ant = texto[i - 1] if i > 0 else ""
            if prox == "" or prox.isspace():
                saida.append((i, "separador"))
            elif ant.isdigit() and prox.isdigit():
                saida.append((i, "decimal"))
            else:
                saida.append((i, "invalida"))
    return saida


def eh_compacta(texto) -> bool:
    """A linha usa ", " (vírgula + espaço) fora de parênteses → formato compacto."""
    return any(t == "separador" for _, t in _virgulas(str(texto if texto is not None else "")))


def _resolver_nome(nome, ctx):
    """Apelido cadastrado → nome oficial; notação compacta de poste → DT11/600; preserva os parâmetros (…(PARAMS))."""
    nome = str(nome).strip().upper()
    i = nome.find("(")
    base, params = (nome[:i], nome[i:]) if i != -1 else (nome, "")
    ctx = ctx or {}
    if base in (ctx.get("nomes") or ()):
        return base + params, None
    if base in (ctx.get("apelidos") or {}):
        return ctx["apelidos"][base] + params, None
    c = compacto_para_poste(base, ctx.get("postes"))
    if c is not None:
        if c["erro"]:
            return base + params, c["erro"]
        return c["ativo"] + params, None
    return base + params, None


def _qtd(txt):
    t = txt.strip()
    neg = t.startswith("*")
    v = float((t[1:] if neg else t).replace(",", "."))
    return -v if neg else v


def _com_param(ativo, nome, valor, assinaturas, k=0):
    """Acrescenta ``NOME=valor`` ao pedido quando a assinatura tem o parâmetro e o pedido não o informou (``N1`` ou ``N1(A)``).
    ``k`` (TASK-096) = a repetição do parâmetro (0 = o 1º ``cabo``, 1 = o 2º ``cabo``…)."""
    i = ativo.find("(")
    base = ativo[:i] if i != -1 else ativo
    sig = assinaturas.get(base)
    alvo = nome if not k else f"{nome}#{k + 1}"
    if not sig or alvo not in sig["nomes"]:
        return ativo
    if i == -1:
        return f"{ativo}({nome}={valor})"
    if not ativo.endswith(")"):
        return ativo
    itens = [x.strip() for x in ativo[i + 1:-1].split(",") if x.strip()]
    if sum(1 for x in itens if x.upper().startswith(nome + "=")) > k:
        return ativo
    if not any("=" in x for x in itens) and len(itens) > sig["nomes"].index(alvo):
        return ativo                                                  # pedido posicional que já traz o parâmetro
    return f"{ativo[:-1]},{nome}={valor})"


def _resolve_para_poste(token, ctx) -> bool:
    """O token sozinho é um poste (compacto ou apelido de um poste da base)?"""
    nome, erro = _resolver_nome(token, ctx)
    return erro is None and eh_poste(nome) and not re.match(r"^\*?\d+(?:[.,]\d+)?[-xX]", token)


def interpretar_linha(texto, operacao, classico, ctx=None) -> dict:
    """Traduz uma linha de Outros. Retorna {"formato": "compacto"|"antigo", "itens": [{"ativo","qtd","operacao"}], "erros": [{"cod","msg"}]}.

    ``classico(texto)`` interpreta o formato ANTIGO como o chamador já fazia (lista de {"ativo","qtd"}); ``ctx`` (opcional) =
    {"apelidos": {...}, "nomes": set, "postes": set, "assinaturas": {...}} vindo da base do projeto (contexto_da_base); opcional
    "parametros_linha": {"CABO": "C"} = parâmetros que o vínculo cabo<->estrutura entrega a ESTA linha (TASK-084)."""
    txt = str(texto if texto is not None else "").strip()
    op = str(operacao if operacao is not None else "").strip().upper()
    r = {"formato": "antigo", "itens": [], "erros": []}
    if not txt:
        return r
    virgulas = _virgulas(txt)
    ruins = [i for i, t in virgulas if t == "invalida"]
    if ruins:
        r["erros"].append({"cod": "VIRGULA", "msg": f"vírgula sem espaço em '{txt}': use vírgula + espaço para separar itens (decimal só entre dígitos, ex.: 0,5)"})
        r["formato"] = "compacto" if any(t == "separador" for _, t in virgulas) else "antigo"
        return r
    brutos = []
    if not virgulas and len(dividir_tokens(txt)) == 1 and _resolve_para_poste(txt, ctx):
        # TASK-082: poste sozinho na linha ("11300D", como o Leitor de Pernambuco entrega) = 1 poste, sem quantidade
        r["formato"] = "compacto"
        brutos = [{"ativo": txt, "qtd": 1.0}]
    elif any(t == "separador" for _, t in virgulas) or not any(_RE_QTD_CLASSICA.match(p) for p in dividir_tokens(txt)):
        # o ÚNICO delimitador é ", " (vírgula + espaço): um ativo com espaço no nome ("CE3 PR") é UM item. Linha sem vírgula e sem nenhum "qtd-ativo" também é compacta (um item só).
        r["formato"] = "compacto"
        cortes = [i for i, t in virgulas if t == "separador"]
        partes, ini = [], 0
        for c in cortes:
            partes.append(txt[ini:c])
            ini = c + 1
        partes.append(txt[ini:])
        for parte in partes:
            palavras = dividir_tokens(parte.strip())
            if not palavras:
                r["erros"].append({"cod": "FORMATO", "msg": f"item vazio em '{txt}'"})
                continue
            if len(palavras) >= 2 and _RE_QTD.match(palavras[0]):
                qtd, nome = _qtd(palavras[0]), " ".join(palavras[1:])      # "2 CE3 PR" = quantidade 2 de "CE3 PR"
            else:
                qtd, nome = 1.0, " ".join(palavras)                          # "CE3 PR" = 1 de "CE3 PR"
            if _RE_QTD_ANTIGA.match(nome):
                r["erros"].append({"cod": "MISTURA", "msg": f"'{nome}' usa a quantidade do formato antigo (qtd-ativo) numa linha em vírgulas; use 'quantidade ativo'"})
                continue
            brutos.append({"ativo": nome, "qtd": qtd})
    else:
        brutos = [{"ativo": i["ativo"], "qtd": i["qtd"]} for i in classico(txt)]
    if r["erros"]:
        return r
    # nomes: apelido / notação compacta → nome oficial
    itens = []
    for b in brutos:
        nome, erro = _resolver_nome(b["ativo"], ctx)
        if erro:
            r["erros"].append({"cod": "FORMATO", "msg": erro})
        itens.append({"ativo": nome, "qtd": b["qtd"], "operacao": op})
    if r["erros"]:
        return r
    # parâmetro poste: o poste mais próximo à esquerda
    assinaturas = (ctx or {}).get("assinaturas") or {}
    parametros_linha = (ctx or {}).get("parametros_linha") or {}
    parametros_estrutura = (ctx or {}).get("parametros_estrutura") or {}        # TASK-096: por estrutura da linha (N1 e N4 não se misturam)
    poste_atual = None
    for it in itens:
        if eh_poste(it["ativo"]):
            poste_atual = poste_para_compacto(it["ativo"])
        else:
            for nome, valor in (parametros_estrutura.get(nome_base(it["ativo"])) or parametros_linha).items():          # TASK-084: parâmetro que a estrutura recebe do vínculo (ex.: CABO=C); TASK-096: lista = um valor por repetição
                for k, v in enumerate(valor if isinstance(valor, list) else [valor]):
                    it["ativo"] = _com_param(it["ativo"], nome, v, assinaturas, k)
            if poste_atual:
                it["ativo"] = _com_param(it["ativo"], "POSTE", poste_atual, assinaturas)
    # operação M/*M: poste ignorado, demais itens I/*I
    if op in ("M", "*M"):
        itens = [dict(i, operacao="I" if op == "M" else "*I") for i in itens if not eh_poste(nome_base(i["ativo"]))]
    r["itens"] = itens
    return r


def texto_equivalente(texto, classico, ctx=None) -> str:
    """Texto no formato antigo (``1-DT11/600 1-N1``), só para CLASSIFICAR a linha (sem regra do M e sem parâmetros)."""
    r = interpretar_linha(texto, "", classico, ctx)
    if r["erros"] or not r["itens"]:
        return str(texto if texto is not None else "")
    return " ".join(f"{_fmt(i['qtd'])}-{re.sub(r'[(].*$', '', i['ativo'])}" for i in r["itens"])


def _fmt(q) -> str:
    return str(int(q)) if float(q).is_integer() else str(q)


def contexto_da_base(linhas, projeto) -> dict:
    """Contexto da base do projeto para interpretar_linha: apelidos, nomes oficiais, postes e assinaturas dos parâmetros."""
    p = str(projeto or "").strip().upper()
    nomes, postes = set(), set()
    for l in linhas or []:
        lp = str(l.get("projeto") or "").strip().upper()
        if lp and p and p not in [x.strip() for x in lp.split("/")]:
            continue
        cel = (l.get("ativo") or "").strip()
        if not cel:
            continue
        a = analisar_celula_base(cel)
        if a["erro"] or not a["base"]:
            continue
        nomes.add(a["base"])
        if eh_poste(a["base"]):
            postes.add(a["base"])
    return {"apelidos": apelidos_do_projeto(linhas, projeto), "nomes": nomes, "postes": postes, "assinaturas": assinaturas_do_projeto(linhas, projeto)}


def codigos_estruturas(texto, classico, ctx=None) -> list:
    """Nomes de TODOS os itens que não são poste, na ordem da linha (TASK-094: `9400D, N1, N4` → ["N1", "N4"]); vazio se só houver poste."""
    r = interpretar_linha(texto, "", classico, ctx)
    return [nome_base(i["ativo"]) for i in r["itens"] if not eh_poste(nome_base(i["ativo"]))]


def codigo_estrutura(texto, classico, ctx=None) -> str:
    """Nome do primeiro item que NÃO é poste (a estrutura da linha, para o vínculo); vazio se não houver."""
    r = interpretar_linha(texto, "", classico, ctx)
    for i in r["itens"]:
        if not eh_poste(nome_base(i["ativo"])):
            return nome_base(i["ativo"])
    return ""
