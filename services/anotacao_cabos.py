"""
services/anotacao_cabos.py — Notação do LABEL na tabela Cabos (TASK-083).

Uma linha de Cabos pode vir no formato ANTIGO (``CAA2 ABC 35 m`` = ATIVO FASE COMPRIMENTO) ou na notação do LABEL do projeto
(``3xA35P(7,94)-13,8kV 45 m``). O usuário vê o que digitou; a tradução é implícita (esta função), para a Totalizadora, os contadores
e o vínculo:

    [Nx]CODIGO[(neutro)][-TENSÃOkV] COMPRIMENTO [m]
    * Nx = fases (1x implícito); a tensão é ignorada (mas guardada: com tensão = rede de média, sem = baixa).
    * neutro entre parênteses: com vírgula (7,94) → CAZ7,94; numérico (16) → mesma família do cabo com a seção do parêntese
      (C16(16) → C16; S50P(35) → S35P); (S04) → o próprio cabo S04. Vira um item à parte (qtd 1, mesmo comprimento).
    * (S04) / [C16] soltos = cabo de neutro próprio; "16" sozinho = C16; RAMA25/PB35 = 1xRAMA25 / 1xPB35;
      cabo que começa com C (C 3x1x95(95)+20kV) = um único código sem espaços, 1 fase.

Mesma lógica de static/anotacao_cabos.js (paridade testada em tests/test_anotacao_cabos.py).
"""
import re

# `\s`/`trim` do JavaScript (inclui BOM e NBSP), `\d`/`\w` só ASCII e `$` só no fim da string: assim a paridade com static/anotacao_cabos.js vale
# até para texto estranho (testada com ruído em tests/test_anotacao_cabos.py).
_WS = "\t\n\v\f\r \u00a0\u1680\u2000\u2001\u2002\u2003\u2004\u2005\u2006\u2007\u2008\u2009\u200a\u2028\u2029\u202f\u205f\u3000\ufeff"
_S = "[" + _WS + "]"
_A = re.ASCII

_RE_SEPARA = re.compile(r"^(.*[^" + _WS + r"])" + _S + r"+([0-9]+(?:[.,][0-9]+)?)" + _S + r"*[mM]?\Z", _A | re.S)
_RE_TENSAO = re.compile(_S + r"*[-+]" + _S + r"*([0-9]+(?:[.,][0-9]+)?)" + _S + r"*KV" + _S + r"*\Z", _A | re.I)
_RE_SOLTO = re.compile(r"^[(\[]" + _S + r"*([A-Z][A-Za-z0-9_]*)" + _S + r"*[)\]]\Z", _A)
_RE_POTENCIA = re.compile(r"^C" + _S + r"*[0-9][xX]", _A)
_RE_CODIGO = re.compile(r"^(?:([0-9])[xX])?([A-Z]{0,4}[0-9][0-9A-Z/,.]*)(?:\(([^)]*)\))?\Z", _A)
_RE_SECAO = re.compile(r"^([A-Z]*)([0-9]+)([A-Z]*)\Z", _A)
_RE_CURTO = re.compile(r"^(?:RAMA|PB)[0-9]+\Z", _A)
_RE_NUMERO_SOZINHO = re.compile(r"^[0-9]{2,3}\Z", _A)
_RE_MARCA = re.compile(r"^[0-9][xX]|^C" + _S + r"*[0-9][xX]|[(\[]|KV", _A | re.I)
_RE_ESPACOS = re.compile(_S + "+")


def _trim(t: str) -> str:
    return t.strip(_WS)


def _separar(texto: str):
    t = _trim(_RE_ESPACOS.sub(" ", str(texto if texto is not None else "")))
    m = _RE_SEPARA.match(t)
    return (_trim(m.group(1)), m.group(2)) if m else (t, None)


def eh_notacao(texto) -> bool:
    """A linha usa a notação do LABEL (e não o formato antigo ATIVO FASE COMPRIMENTO)?"""
    rotulo, _ = _separar(texto)
    r = rotulo.upper()
    if not r:
        return False
    return bool(_RE_MARCA.search(r) or _RE_CURTO.match(r) or _RE_NUMERO_SOZINHO.match(r))


def _num(txt) -> float:
    return float(str(txt).replace(",", "."))


def _neutro_de(codigo: str, valor: str):
    v = _trim(valor).upper()
    if not v:
        return None
    if "," in v:
        return "CAZ" + v
    if re.fullmatch(r"[0-9]+", v, _A):
        m = _RE_SECAO.match(codigo)
        return f"{m.group(1)}{v}{m.group(3)}" if m else codigo
    return v


def interpretar(texto) -> dict:
    """{"formato": "notacao"|"antigo", "itens": [{"ativo","fases","comprimento","neutro"}], "tensao": str|None, "erros": [msg]}."""
    r = {"formato": "antigo", "itens": [], "tensao": None, "erros": []}
    if not eh_notacao(texto):
        return r
    r["formato"] = "notacao"
    rotulo, comp = _separar(texto)
    rotulo = rotulo.upper()
    comprimento = _num(comp) if comp is not None else None
    if comprimento is None:
        r["erros"].append(f"falta o comprimento em '{rotulo}' (use: rótulo comprimento m)")
    m = _RE_TENSAO.search(rotulo)
    if m:
        r["tensao"] = m.group(1)
        rotulo = _trim(rotulo[:m.start()])

    def add(ativo, fases, neutro):
        r["itens"].append({"ativo": ativo, "fases": fases, "comprimento": comprimento, "neutro": neutro})

    solto = _RE_SOLTO.match(rotulo)
    if solto:
        add(solto.group(1), 1, True)
    elif _RE_POTENCIA.match(rotulo):
        add(_RE_ESPACOS.sub("", rotulo), 1, False)
    elif _RE_NUMERO_SOZINHO.match(rotulo):
        add("C" + rotulo, 1, False)
    else:
        c = _RE_CODIGO.match(rotulo.replace(" ", ""))
        if not c:
            r["erros"].append(f"rótulo '{rotulo}' não reconhecido (ex.: 3xA35P(7,94)-13,8kV 45 m)")
            return r
        fases = int(c.group(1)) if c.group(1) else 1
        codigo = c.group(2)
        add(codigo, fases, False)
        if c.group(3) is not None:
            neutro = _neutro_de(codigo, c.group(3))
            if neutro is None:
                r["erros"].append(f"neutro vazio em '{rotulo}'")
            else:
                add(neutro, 1, True)
    return r


def fases_da_linha(texto) -> int:
    """Quantidade de ativos (fases) do cabo principal; 1 se não for notação ou não entender."""
    r = interpretar(texto)
    principais = [i for i in r["itens"] if not i["neutro"]] or r["itens"]
    return principais[0]["fases"] if principais else 1


def prefixo_e_comprimento(texto):
    """Para o vínculo/contadores: (código do cabo principal, comprimento) — None se a linha não for notação ou não tiver comprimento."""
    r = interpretar(texto)
    principais = [i for i in r["itens"] if not i["neutro"]] or r["itens"]
    if not principais or principais[0]["comprimento"] is None:
        return None
    return principais[0]["ativo"], principais[0]["comprimento"]
