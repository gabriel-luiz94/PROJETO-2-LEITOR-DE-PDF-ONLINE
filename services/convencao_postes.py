"""
services/convencao_postes.py — Convenção de postes por projeto (TASK-174, ADR-007).

O que é poste na linha de Outros (tipos, formatos de nome, notação compacta, faixas de altura/esforço) vem de uma tabela do projeto (`postes`, guardada na
configuração da vinculação). Sem tabela, vale a de Pernambuco (PADRAO), idêntica ao que era fixo no código.

  {"tipos": [{"codigo": "DT", "sufixo": "D", "nome": "duplo T"}, …],
   "formatos": ["{tipo}{altura}/{esforco}", "{sufixo}{altura}/{esforco}", "{altura}{esforco}{sufixo}"],   # formas aceitas de escrever o poste (a 1ª é o nome oficial)
   "altura": [7, 16], "esforco_digitos": [3, 4], "compacto": "{altura}{esforco}{sufixo}"}
`{altura}{esforco}` colados (sem separador) são ambíguos: a base decide (poste que existe vence) e, sem base, a faixa de altura e de dígitos do esforço.
Mesma lógica de static/convencao_postes.js (paridade testada em tests/test_convencao_postes_174.py).
"""
import json
import re

PADRAO = {
    "tipos": [{"codigo": "DT", "sufixo": "D", "nome": "duplo T"}, {"codigo": "RC", "sufixo": "R", "nome": "redondo"},
              {"codigo": "CV", "sufixo": "F", "nome": "concreto"}, {"codigo": "FV", "sufixo": "FV", "nome": "fibra de vidro"}],
    "formatos": ["{tipo}{altura}/{esforco}", "{sufixo}{altura}/{esforco}", "{altura}{esforco}{sufixo}"],
    "altura": [7, 16], "esforco_digitos": [3, 4], "compacto": "{altura}{esforco}{sufixo}",
}
_PH = re.compile(r"(\{tipo\}|\{sufixo\}|\{altura\}\{esforco\}|\{altura\}|\{esforco\})")
_cache = {}


def _trim(t) -> str:
    return str(t if t is not None else "").strip()


def normalizar(conv) -> dict:
    """Tabela completa: sem `tipos` válidos vale o padrão de Pernambuco; campos que faltam usam o padrão."""
    if not isinstance(conv, dict) or not isinstance(conv.get("tipos"), list) or not any(isinstance(t, dict) and _trim(t.get("codigo")) for t in conv["tipos"]):
        return json.loads(json.dumps(PADRAO))
    c = {"tipos": [{"codigo": _trim(t.get("codigo")).upper(), "sufixo": (_trim(t.get("sufixo")) or _trim(t.get("codigo"))).upper(), "nome": _trim(t.get("nome"))}
                   for t in conv["tipos"] if isinstance(t, dict) and _trim(t.get("codigo"))]}
    f = [_trim(x) for x in conv["formatos"] if _trim(x)] if isinstance(conv.get("formatos"), list) else []
    c["formatos"] = f or list(PADRAO["formatos"])
    for k, padrao in (("altura", PADRAO["altura"]), ("esforco_digitos", PADRAO["esforco_digitos"])):
        v = conv.get(k)
        c[k] = [int(v[0]), int(v[1])] if isinstance(v, list) and len(v) == 2 and all(isinstance(x, (int, float)) and not isinstance(x, bool) for x in v) else list(padrao)
    c["compacto"] = _trim(conv.get("compacto")) or PADRAO["compacto"]
    return c


def validar(conv) -> list:
    erros = []
    if conv is None:
        return erros
    if not isinstance(conv, dict):
        return ["'postes' precisa ser um objeto."]
    tipos = conv.get("tipos")
    if not isinstance(tipos, list) or not tipos:
        erros.append("Postes: 'tipos' precisa ser uma lista com pelo menos um tipo ({\"codigo\": \"DT\", \"sufixo\": \"D\"}).")
        tipos = []
    for n, t in enumerate(tipos, 1):
        if not isinstance(t, dict) or not _trim(t.get("codigo")):
            erros.append(f"Postes, tipo {n}: falta o 'codigo'.")
    fmts = conv.get("formatos")
    if fmts is not None and (not isinstance(fmts, list) or not all(isinstance(x, str) for x in fmts)):
        erros.append("Postes: 'formatos' precisa ser uma lista de textos (ex.: \"{tipo}{altura}/{esforco}\").")
    for fm in (fmts if isinstance(fmts, list) else []) + ([conv["compacto"]] if isinstance(conv.get("compacto"), str) else []):
        if not isinstance(fm, str):
            continue
        for ph in ("tipo", "sufixo", "altura", "esforco"):
            if fm.count("{" + ph + "}") > 1:
                erros.append(f"Postes: o formato '{fm}' repete {{{ph}}}.")
        if "{altura}" not in fm or "{esforco}" not in fm:
            erros.append(f"Postes: o formato '{fm}' precisa de {{altura}} e {{esforco}}.")
        if "{tipo}" not in fm and "{sufixo}" not in fm:
            erros.append(f"Postes: o formato '{fm}' precisa de {{tipo}} ou {{sufixo}}.")
        if "{esforco}{altura}" in fm:
            erros.append(f"Postes: no formato '{fm}', {{altura}}{{esforco}} colados precisam estar nessa ordem.")
    for k in ("altura", "esforco_digitos"):
        v = conv.get(k)
        if v is not None and not (isinstance(v, list) and len(v) == 2 and all(isinstance(x, (int, float)) and not isinstance(x, bool) for x in v) and v[0] <= v[1]):
            erros.append(f"Postes: '{k}' precisa ser [mínimo, máximo].")
    return erros


def _compilar(conv):
    chave = json.dumps(conv, sort_keys=True)
    if chave in _cache:
        return _cache[chave]
    tipos = sorted({t["codigo"] for t in conv["tipos"]}, key=lambda x: (-len(x), x))
    sufs = sorted({t["sufixo"] for t in conv["tipos"]}, key=lambda x: (-len(x), x))
    saida = []
    for fm in conv["formatos"]:
        rx, juntos = "", False
        for parte in _PH.split(fm):
            if parte == "{tipo}":
                rx += "(?P<tipo>" + "|".join(re.escape(x) for x in tipos) + ")"
            elif parte == "{sufixo}":
                rx += "(?P<sufixo>" + "|".join(re.escape(x) for x in sufs) + ")"
            elif parte == "{altura}{esforco}":
                rx, juntos = rx + r"(?P<digitos>\d+)", True
            elif parte == "{altura}":
                rx += r"(?P<altura>\d+)"
            elif parte == "{esforco}":
                rx += r"(?P<esforco>\d+)"
            else:
                rx += re.escape(parte)
        try:
            saida.append((re.compile("^" + rx + r"\Z", re.I | re.A), juntos))
        except re.error:
            continue
    _cache[chave] = saida
    return saida


def _tipo_de(conv, g):
    if g.get("tipo"):
        return g["tipo"].upper()
    suf = (g.get("sufixo") or "").upper()
    for t in conv["tipos"]:
        if t["sufixo"] == suf:
            return t["codigo"]
    return None


def _sufixo_de(conv, tipo):
    for t in conv["tipos"]:
        if t["codigo"] == tipo:
            return t["sufixo"]
    return tipo


def _render(modelo, conv, tipo, altura, esforco):
    return modelo.replace("{tipo}", tipo).replace("{sufixo}", _sufixo_de(conv, tipo)).replace("{altura}", str(altura)).replace("{esforco}", str(esforco))


def _explicito(nome, conv):
    """(tipo, altura, esforço) de um nome com separador (DT11/600, R12/1000, P11-400); None se nenhum formato explícito casa."""
    t = _trim(nome).upper()
    for rx, juntos in _compilar(conv):
        if juntos:
            continue
        m = rx.match(t)
        if m:
            tipo = _tipo_de(conv, m.groupdict())
            if tipo:
                return tipo, m.group("altura"), m.group("esforco")
    return None


def _colado(token, conv):
    """(tipo, dígitos) de um token com altura e esforço colados (11600D); None se não é."""
    t = _trim(token)
    for rx, juntos in _compilar(conv):
        if not juntos:
            continue
        m = rx.match(t)
        if m:
            tipo = _tipo_de(conv, m.groupdict())
            if tipo:
                return tipo, m.group("digitos")
    return None


def compacto_para_poste(token, postes=None, conv=None):
    """Notação compacta → poste. None se o token não é compacto; senão {"ativo", "compacto", "erro"} (a base decide a ambiguidade; sem base, faixas da convenção)."""
    conv = normalizar(conv)
    t = _trim(token)
    c = _colado(t, conv)
    if not c:
        return None
    tipo, digitos = c
    (amin, amax), (dmin, dmax) = conv["altura"], conv["esforco_digitos"]
    cands = []
    for hl in (2, 1):
        h, e = digitos[:hl], digitos[hl:]
        if not dmin <= len(e) <= dmax or h.startswith("0") or e.startswith("0") or not h:
            continue
        cands.append((int(h), _render(conv["formatos"][0], conv, tipo, int(h), int(e))))
    r = {"ativo": None, "compacto": t.upper(), "erro": None}
    if postes:
        na_base = [n for _, n in cands if n in postes]
        if len(na_base) == 1:
            r["ativo"] = na_base[0]
            return r
    plausiveis = [n for h, n in cands if amin <= h <= amax]
    if len(plausiveis) == 1:
        r["ativo"] = plausiveis[0]
    else:
        r["erro"] = f"não foi possível separar altura e esforço de '{t.upper()}'" if not plausiveis else f"'{t.upper()}' é ambíguo ({' ou '.join(plausiveis)})"
    return r


def compacto_do_poste(nome, conv=None):
    """Nome de poste → notação compacta, ou None (DT11/600 → 11600D; o próprio compacto vale quando se resolve)."""
    conv = normalizar(conv)
    t = _trim(nome).upper()
    e = _explicito(t, conv)
    if e:
        return _render(conv["compacto"], conv, e[0], e[1], e[2])
    if _colado(t, conv):
        c = compacto_para_poste(t, None, conv)
        if c is not None and not c["erro"]:
            return t
    return None


def eh_poste(nome, conv=None) -> bool:
    return compacto_do_poste(nome, conv) is not None


def testar(nome, conv=None, postes=None) -> dict:
    """"Testar um nome" da tela: é poste? nome oficial, compacto e o motivo quando não resolve."""
    conv = normalizar(conv)
    t = _trim(nome).upper()
    e = _explicito(t, conv)
    if e:
        return {"poste": True, "tipo": e[0], "ativo": _render(conv["formatos"][0], conv, e[0], e[1], e[2]), "compacto": _render(conv["compacto"], conv, e[0], e[1], e[2]), "erro": None}
    c = compacto_para_poste(t, postes, conv)
    if c is None:
        return {"poste": False, "tipo": None, "ativo": None, "compacto": None, "erro": None}
    tipo = _colado(t, conv)[0]
    return {"poste": not c["erro"], "tipo": tipo, "ativo": c["ativo"], "compacto": t if not c["erro"] else None, "erro": c["erro"]}
