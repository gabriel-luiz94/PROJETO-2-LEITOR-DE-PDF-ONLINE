"""
services/dicionario_cabos.py — Dicionário de cabos por projeto (TASK-173, ADR-007).

O que o cabo É (família, valor do tipocabo, seção, nível, fases, apelidos e neutro) vira tabela editável do projeto, guardada na configuração da
vinculação (`dicionario_cabos`). Sem dicionário (lista vazia), tudo funciona como antes: as regras fixas de Pernambuco continuam valendo.
A Tabela de Cabos da vinculação continua dizendo EM QUAIS ESTRUTURAS o cabo entra; o dicionário diz o que ele é.

Linha (a primeira que casa vence):
  {"padrao": "^A(\\d+)P$", "familia": "PROTEGIDO", "valor_cabo": "X", "secao": "$1" | 185 , "awg": {"4": 21.2, "1/0": 53.5},
   "nivel": "MT"|"BT"|"", "fases_padrao": 1, "notacao_solta": true (o código sozinho, sem Nx/tensão, já é notação: "ALU-50 30 m"), "apelidos": ["^(?:RAMA|PB)\\d+$", {"padrao": "^(\\d{2,3})$", "codigo": "C$1"}], "neutro": "CAZ{valor}", "ativa": true}
Mesma lógica de static/dicionario_cabos.js (paridade testada em tests/test_dicionario_cabos_173.py).
"""
import re

_A = re.ASCII | re.I
CAMPOS = ("padrao", "familia", "valor_cabo", "secao", "awg", "nivel", "fases_padrao", "apelidos", "neutro", "notacao_solta", "ativa", "obs")
NEUTRO_PADRAO = "CAZ{valor}"


def _trim(t) -> str:
    return str(t if t is not None else "").strip()


def _re(p):
    try:
        return re.compile(_trim(p), _A)
    except re.error:
        return None


def normalizar(dic) -> list:
    return [dict(r) for r in dic if isinstance(r, dict)] if isinstance(dic, list) else []


def validar(dic) -> list:
    """Mensagens de erro (lista vazia = ok)."""
    erros = []
    if dic is None:
        return erros
    if not isinstance(dic, list):
        return ["'dicionario_cabos' precisa ser uma lista de linhas."]
    for n, r in enumerate(dic, 1):
        p = f"Dicionário de cabos, linha {n}"
        if not isinstance(r, dict):
            erros.append(f"{p}: precisa ser um objeto.")
            continue
        if not _trim(r.get("padrao")):
            erros.append(f"{p}: falta o 'padrao' (regex do código do cabo).")
        elif _re(r["padrao"]) is None:
            erros.append(f"{p}: 'padrao' não é uma regex válida.")
        if _trim(r.get("nivel")) and _trim(r["nivel"]).upper() not in ("MT", "BT"):
            erros.append(f"{p}: 'nivel' precisa ser MT, BT ou vazio.")
        fp = r.get("fases_padrao")
        if fp not in (None, "") and not (isinstance(fp, int) and not isinstance(fp, bool) and fp >= 1):
            erros.append(f"{p}: 'fases_padrao' precisa ser um número inteiro maior que zero.")
        s = r.get("secao")
        if s not in (None, "") and not (isinstance(s, (int, float)) and not isinstance(s, bool)) and not re.fullmatch(r"\$[0-9]", _trim(s)):
            erros.append(f"{p}: 'secao' precisa ser um número ou $1…$9 (grupo da regex).")
        ap = r.get("apelidos")
        if ap is not None and not isinstance(ap, list):
            erros.append(f"{p}: 'apelidos' precisa ser uma lista.")
        for a in ap if isinstance(ap, list) else []:
            pa = a.get("padrao") if isinstance(a, dict) else a
            if not _trim(pa) or _re(pa) is None:
                erros.append(f"{p}: apelido '{pa}' não é uma regex válida.")
        if r.get("notacao_solta") is not None and not isinstance(r["notacao_solta"], bool):
            erros.append(f"{p}: 'notacao_solta' precisa ser true ou false.")
        nt = r.get("neutro")
        if nt not in (None, "") and "{valor}" not in _trim(nt):
            erros.append(f"{p}: 'neutro' precisa conter {{valor}} (ex.: CAZ{{valor}}).")
        aw = r.get("awg")
        if aw not in (None, "") and not (isinstance(aw, dict) and all(isinstance(v, (int, float)) for v in aw.values())):
            erros.append(f"{p}: 'awg' precisa ser {{\"4\": 21.2, \"1/0\": 53.5}}.")
    return erros


def linha(dic, codigo):
    """(índice, linha, match) da primeira linha ativa cujo `padrao` casa o código; (None, None, None) se nenhuma."""
    cod = _trim(codigo).upper()
    if not cod:
        return None, None, None
    for i, r in enumerate(normalizar(dic)):
        if r.get("ativa") is False:
            continue
        rx = _re(r.get("padrao"))
        m = rx.search(cod) if rx else None
        if m:
            return i, r, m
    return None, None, None


def secao(dic, codigo):
    """Seção em mm² pelo dicionário, ou None (a regra fixa do código decide)."""
    _, r, m = linha(dic, codigo)
    if r is None:
        return None
    s = r.get("secao")
    if s in (None, ""):
        return None
    if isinstance(s, (int, float)) and not isinstance(s, bool):
        return float(s)
    g = re.fullmatch(r"\$([0-9])", _trim(s))
    if not g or int(g.group(1)) > len(m.groups()) or m.group(int(g.group(1))) is None:
        return None
    txt = m.group(int(g.group(1))).upper()
    aw = r.get("awg") if isinstance(r.get("awg"), dict) else {}
    if txt in aw:
        return float(aw[txt])
    try:
        return float(txt.replace(",", "."))
    except ValueError:
        return None


def valor_cabo(dic, codigo):
    _, r, _ = linha(dic, codigo)
    return (_trim(r.get("valor_cabo")).upper() or None) if r else None


def nivel_padrao(dic, codigo):
    _, r, _ = linha(dic, codigo)
    n = _trim(r.get("nivel")).upper() if r else ""
    return n if n in ("MT", "BT") else None


def fases_padrao(dic, codigo):
    _, r, _ = linha(dic, codigo)
    fp = r.get("fases_padrao") if r else None
    return fp if isinstance(fp, int) and not isinstance(fp, bool) and fp >= 1 else None


def neutro_template(dic, codigo) -> str:
    _, r, _ = linha(dic, codigo)
    t = _trim(r.get("neutro")) if r else ""
    return t if t and "{valor}" in t else NEUTRO_PADRAO


def apelido(dic, rotulo):
    """Se o texto é um apelido do dicionário: (True, código) — o código é o do apelido com objeto {"padrao","codigo": "C$1"} ou o próprio texto. Senão (False, None)."""
    t = _trim(rotulo).upper().replace(" ", "")
    if not t:
        return False, None
    for r in normalizar(dic):
        if r.get("ativa") is False or not isinstance(r.get("apelidos"), list):
            continue
        for a in r["apelidos"]:
            pa = a.get("padrao") if isinstance(a, dict) else a
            rx = _re(pa)
            m = rx.search(t) if rx else None
            if not m:
                continue
            cod = _trim(a.get("codigo")).upper() if isinstance(a, dict) else ""
            if not cod:
                return True, t
            return True, re.sub(r"\$([0-9])", lambda g: (m.group(int(g.group(1))) or "") if int(g.group(1)) <= len(m.groups()) else "", cod).upper()
    return False, None


def reconhece_codigo(dic, codigo, solta=False) -> bool:
    """O código casa uma linha do dicionário? Com `solta`, só linhas marcadas `notacao_solta: true` (o código sozinho já indica a notação do LABEL)."""
    _, r, _ = linha(dic, codigo)
    return r is not None and (not solta or r.get("notacao_solta") is True)


def ordem_maior_padrao(dic):
    """Ordem do "maior primeiro" vinda do dicionário: os valor_cabo distintos na ordem das linhas (vazio = sem sugestão)."""
    ordem = []
    for r in normalizar(dic):
        v = _trim(r.get("valor_cabo")).upper()
        if v and v not in ordem:
            ordem.append(v)
    return ordem


def testar(dic, texto) -> dict:
    """"Testar um texto" da tela: o que o dicionário entende de uma linha de Cabos (código, fases, neutro, nível, seção, valor_cabo e a linha usada)."""
    from services import anotacao_cabos as ac                                                              # import tardio: anotacao_cabos importa este módulo
    r = ac.interpretar(texto, dic)
    principais = [i for i in r["itens"] if not i["neutro"]] or r["itens"]
    p = principais[0] if principais else None
    if not p:
        return {"formato": r["formato"], "codigo": None, "fases": None, "neutro": None, "nivel": None, "tensao": r["tensao"], "secao": None, "valor_cabo": None, "linha": None, "erros": r["erros"]}
    neutros = [i["ativo"] for i in r["itens"] if i["neutro"]]
    tens = float(r["tensao"].replace(",", ".")) if r["tensao"] else None
    nivel = ("MT" if tens > 1 else "BT") if tens is not None else (nivel_padrao(dic, p["ativo"]) or "BT")
    i, _, _ = linha(dic, p["ativo"])
    return {"formato": r["formato"], "codigo": p["ativo"], "fases": p["fases"], "neutro": neutros[0] if neutros else None, "nivel": nivel, "tensao": r["tensao"],
            "secao": secao(dic, p["ativo"]), "valor_cabo": valor_cabo(dic, p["ativo"]), "linha": i, "erros": r["erros"]}
