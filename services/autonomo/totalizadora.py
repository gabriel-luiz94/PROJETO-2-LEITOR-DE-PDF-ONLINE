"""
services/autonomo/totalizadora.py — Tabela Totalizadora e payload do orçamento (TASK-031, fase B).

PORTE FIEL de `syncTotalizadora` e `obterPayloadCalculo` (static/resumo.js): converte Cabos/Outros em linhas de orçamento (uma por
ativo, com quantidade), aplica as REGRAS DE CONVERSÃO do projeto (`/api/regras/conversao`) e gera o payload de `POST /api/orcamento/calcular`.
Paridade provada em tests/test_autonomo_totalizadora.py executando o JS real. Fidelidade ao JavaScript: `parseFloat`, `Math.round`
(meio para cima, não o arredondamento do Python), `\\s`/`trim`, `\\d` ASCII, `$` só no fim, número→texto do JS.
Limite conhecido: as regex digitadas nas regras de conversão são avaliadas pelo `re` do Python (dialeto quase igual ao do JS).
"""
import math
import re
from decimal import Decimal

from services.autonomo.montagem import _S, _WS, _trim, linha_standalone

_PONTO = "[^\n\r  ]"   # o `.` do JS não casa com quebras de linha
_RE_M1 = re.compile(r"([0-9.,]+)" + _S + r"*m" + _S + r"*\|" + _S + "*(" + _PONTO + "*)", re.I)
_RE_M2 = re.compile("(" + _PONTO + r"+?)" + _S + r"+([0-9.,]+)" + _S + r"*m\Z", re.I)
_RE_OUTROS = re.compile(r"([*\-]?[0-9]+(?:\.[0-9]+)?)[Xx\-](" + _PONTO + r"+)\Z", re.I)
_RE_P_ESPACO = re.compile(r"\bP" + _S + "+", re.ASCII)
_RE_ESPACOS = re.compile(_S + "+")
_RE_FLOAT = re.compile(r"[+-]?(?:Infinity|[0-9]+\.?[0-9]*(?:[eE][+-]?[0-9]+)?|\.[0-9]+(?:[eE][+-]?[0-9]+)?)")
NAN = float("nan")

REGRA_PADRAO = {"origem": "CABOS", "op_de": "", "ativo_de": "", "acao": "ADICAO", "op_para": "", "ativo_para": "",
                "fator": 1.05, "arredondamento": "NORMAL", "val_min": "", "val_max": ""}


# ── primitivas do JavaScript ────────────────────────────────────────────────
def js_para_texto(v) -> str:
    if v is None:
        return "null"
    if isinstance(v, bool):
        return "true" if v else "false"
    if isinstance(v, (int, float)):
        return js_numero_para_texto(v)
    return str(v)


def js_parse_float(v) -> float:
    s = js_para_texto(v).lstrip(_WS)
    m = _RE_FLOAT.match(s)
    if not m:
        return NAN
    t = m.group(0)
    return float(t.replace("Infinity", "inf"))


def js_numero_para_texto(x) -> str:
    """Number.prototype.toString do JS."""
    if isinstance(x, bool):
        return "true" if x else "false"
    x = float(x)
    if math.isnan(x):
        return "NaN"
    if math.isinf(x):
        return "Infinity" if x > 0 else "-Infinity"
    if x == 0:
        return "0"
    sinal = "-" if x < 0 else ""
    d = Decimal(repr(abs(x))).normalize()
    digitos = "".join(map(str, d.as_tuple().digits)).rstrip("0") or "0"
    k = len(digitos)
    n = len(d.as_tuple().digits) + d.as_tuple().exponent    # valor = 0.digitos × 10^n
    if k <= n <= 21:
        t = digitos + "0" * (n - k)
    elif 0 < n <= 21:
        t = digitos[:n] + "." + digitos[n:]
    elif -6 < n <= 0:
        t = "0." + "0" * (-n) + digitos
    else:
        e = n - 1
        t = digitos[0] + ("." + digitos[1:] if k > 1 else "") + "e" + ("+" if e >= 0 else "-") + str(abs(e))
    return sinal + t


def js_round(x: float) -> float:
    """Math.round: meio para cima (+∞)."""
    if math.isnan(x) or math.isinf(x):
        return x
    f = math.floor(x)
    return float(f + 1) if x - f >= 0.5 else float(f)


def _normaliza(x):
    """Como o JSON.stringify do JS grava números: 5.0 → 5; NaN → null."""
    if isinstance(x, float):
        if math.isnan(x) or math.isinf(x):
            return None
        if x.is_integer():
            return int(x)
    return x


# ── Totalizadora ────────────────────────────────────────────────────────────
def _buscar_descricao(base, ativo_nome, projeto_nome):
    ativo = _trim(ativo_nome).upper()
    proj = (projeto_nome or "").upper()
    for r in base:
        if ((r.get("ativo") or "").upper() == ativo or (r.get("codigo") or "").upper() == ativo) and \
                ((r.get("projeto") or "").strip().upper() == proj or (r.get("projeto") or "").strip() == ""):
            return (r.get("desc_ativo") or r.get("componente") or r.get("desc_codigo") or ""), False
    return "", True


def _confere(padrao, alvo) -> bool:
    if padrao == "":
        return True
    if padrao == alvo:
        return True
    try:
        if "%" in padrao:
            rx = "^" + re.sub(r"[.+?^${}()|\[\]\\]", lambda m: "\\" + m.group(0), padrao).replace("%", ".*") + "$"
        else:
            rx = padrao
            if not rx.startswith("^"):
                rx = "^" + rx
            if not rx.endswith("$"):
                rx = rx + "$"
        return re.search(rx, alvo, re.I) is not None
    except re.error:
        return False


def montar_totalizadora(cabos: list, outros: list, regras: list, base_orcamento: list, projeto_nome: str, itens_vinculo: list = None) -> list:
    """Linhas {id, obs, operacao, ativo, qtd, desc, naoEncontrado, origem}. `qtd` pode ser float('nan') (como o NaN do JS).
    `itens_vinculo` (TASK-036): pseudo-itens origem "VINCULO" (ver `services/autonomo/vinculacao.py`), que nunca aparecem
    aqui sem regra de conversão correspondente — diferente de CABOS/OUTROS, que sempre mantêm o item original sem match."""
    base = base_orcamento or []
    brutos = []
    contador = 1
    for idx, item in enumerate(cabos):
        if not item:
            continue
        base_id = contador
        contador += 1
        q, at = 1.0, item.get("ativo") or ""
        m1 = _RE_M1.search(at)
        if m1:
            q = js_parse_float(m1.group(1).replace(",", "."))
            at = _trim(m1.group(2))
        else:
            m2 = _RE_M2.search(at)
            if m2:
                q = js_parse_float(m2.group(2).replace(",", "."))
                at = _trim(m2.group(1))
            elif linha_standalone(at) and idx + 1 < len(cabos):
                prox = cabos[idx + 1]
                if prox and prox.get("ativo"):
                    n1, n2 = _RE_M1.search(prox["ativo"]), _RE_M2.search(prox["ativo"])
                    if n1:
                        q = js_parse_float(n1.group(1).replace(",", "."))
                    elif n2:
                        q = js_parse_float(n2.group(2).replace(",", "."))
        f = at.upper().replace("/", "").replace("CAA ", "CAA", 1).replace("CA ", "CA", 1).replace("CU ", "CU", 1).replace("CAZ ", "CAZ", 1)
        f = _RE_P_ESPACO.sub("P", f)
        base_ativo = _RE_ESPACOS.split(_trim(f))[0] or at
        desc, nao = _buscar_descricao(base, base_ativo, projeto_nome)
        iteracoes = 1
        if item.get("qtdAtivos"):
            p = js_parse_float(item["qtdAtivos"])
            if not math.isnan(p) and p > 0:
                iteracoes = math.floor(p)
        for _ in range(iteracoes):
            brutos.append({"baseId": f"TOT-{base_id}", "obs": item.get("entidade"), "operacao": item.get("operacao") or "I", "ativo": base_ativo,
                           "qtd": q, "desc": desc, "naoEncontrado": nao, "origem": "CABOS"})
    for item in outros:
        if not item:
            continue
        base_id = contador
        contador += 1
        for p in _RE_ESPACOS.split(item.get("ativo") or ""):
            if not _trim(p):
                continue
            q, nome = 1.0, _trim(p)
            m = _RE_OUTROS.match(p)
            if m:
                q_str, negativo = m.group(1), False
                if q_str.startswith("*") or q_str.startswith("-"):
                    negativo, q_str = True, q_str[1:]
                q = js_parse_float(q_str)
                if negativo:
                    q = -q
                nome = _trim(m.group(2))
            desc, nao = _buscar_descricao(base, nome, projeto_nome)
            brutos.append({"baseId": f"TOT-{base_id}", "obs": item.get("entidade"), "operacao": item.get("operacao") or "I", "ativo": nome,
                           "qtd": q, "desc": desc, "naoEncontrado": nao, "origem": "OUTROS"})
    brutos.extend(itens_vinculo or [])

    novos = []
    for item in brutos:
        achou = substitui = False
        for regra in regras or []:
            r_origem = (regra.get("origem") or "").strip().upper()
            r_op = (regra.get("op_de") or "").strip().upper()
            r_ativo = (regra.get("ativo_de") or "").strip().upper()
            if not (_confere(r_origem, item["origem"]) and _confere(r_op, item["operacao"]) and _confere(r_ativo, item["ativo"])):
                continue
            achou = True
            if regra.get("acao") in ("SUBST", "SUBSTITUIÇÃO", "SUBSTITUICAO"):
                substitui = True
            nova_op = regra["op_para"].strip().upper() if regra.get("op_para") else item["operacao"]
            novo_ativo = regra["ativo_para"].strip().upper() if regra.get("ativo_para") else item["ativo"]
            q = item["qtd"]
            fator = js_parse_float(regra.get("fator"))
            if not math.isnan(fator):
                q = q * fator
            arr = regra.get("arredondamento")
            if arr == "PARA CIMA":
                q = float(math.ceil(q)) if math.isfinite(q) else q
            elif arr == "PARA BAIXO":
                q = float(math.floor(q)) if math.isfinite(q) else q
            elif arr == "INTEIRO":
                q = js_round(q)
            else:
                q = js_round(q * 100) / 100
            vmin = regra.get("val_min")
            if vmin != "" and vmin is not None:
                mn = js_parse_float(vmin)
                if not math.isnan(mn) and q < mn:
                    q = mn
            vmax = regra.get("val_max")
            if vmax != "" and vmax is not None:
                mx = js_parse_float(vmax)
                if not math.isnan(mx) and q > mx:
                    q = mx
            d, nao = _buscar_descricao(base, novo_ativo, projeto_nome)
            novos.append({"id": item["baseId"], "obs": item["obs"], "operacao": nova_op, "ativo": novo_ativo, "qtd": q, "desc": d,
                          "naoEncontrado": nao, "origem": item["origem"]})
        if not achou and item["origem"] == "VINCULO":
            continue   # TASK-036: composto de vínculo sem regra correspondente não é um ativo real — nunca aparece sem match
        if not achou or not substitui:
            c = dict(item)
            c["id"] = item["baseId"]   # o JS mantém também `baseId` na cópia
            novos.append(c)
    return novos


def normalizar_para_json(totalizadora: list) -> list:
    return [{**r, "qtd": _normaliza(r["qtd"])} for r in totalizadora]


def payload_calculo(totalizadora: list) -> dict:
    """obterPayloadCalculo: {cabos, outros} no formato de POST /api/orcamento/calcular."""
    cabos, outros = [], []
    for item in totalizadora:
        if not item:
            continue
        q = item["qtd"]
        if q == "" or q is None or js_parse_float(q) == 0:
            continue
        operacao = item.get("operacao") or "I"
        if isinstance(q, (int, float)) and not isinstance(q, bool) and q < 0:
            q_str = "*" + js_numero_para_texto(abs(q))
        elif isinstance(q, str) and q.startswith("-"):
            q_str = "*" + q[1:]
        else:
            q_str = js_para_texto(q)
        if item["origem"] == "CABOS":
            ativo_final = f"{item['ativo']} 1 {js_para_texto(item['qtd'])}"
        else:
            ativo_final = f"{q_str}-{item['ativo']}"
        obj = {"entidade": item.get("obs") or "0", "operacao": operacao, "ativo": ativo_final}
        (cabos if item["origem"] == "CABOS" else outros).append(obj)
    return {"cabos": cabos, "outros": outros}
