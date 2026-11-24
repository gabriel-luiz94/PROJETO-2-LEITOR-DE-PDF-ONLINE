"""
services/transporte.py — TASK-245: transporte de material (TRANSP / TRANSP_KM) igual à aba TRANSPORT da ICP.

A parte `transporte` do pacote do projeto traz os dados (faixas, pesos dos materiais, km por UTD, município/bairro → UTD).
`calcular` é a tradução célula a célula da aba (anexo .ai/tasks/anexos/TASK-245/transporte_icp.py), SEM "consertar" as bordas da
ICP (4,001 t; carga de 25 t exata; os dois postes que pesam em dobro; km da 1ª linha da UTD). static/transporte.js tem a mesma
função (paridade testada em tests/test_transporte_245.py).

Entrada de `calcular`:
  qtd_odi: {codigo_material: quantidade na ODI (instalação)}
  extra:   {codigo_material: quantidade extra digitada (coluna F)}
  km:      KMx2 da UTD da obra ou None quando o município/UTD não foi informado
  pesos:   lista [(codigo, descricao, kg)] na ordem da aba (com as repetições)
Saída: {'toneladas', 'km', 'servicos': [(codigo_servico, quantidade 2 casas) | ('#VALUE!', '#VALUE!')]} na ordem das linhas 14..17.
"""
import re

FORMATO = "transporte-material"


class _Erro(str):
    """#VALOR! do Excel: propaga em contas, SOMASE e comparações."""


ERRO = _Erro('#VALUE!')


def _err(*xs):
    return any(isinstance(x, _Erro) for x in xs)


BANDAS_ATE50 = ['SDETU5001SI', 'SDETU5003SI', 'SDETU5005SI', 'SDETU5007SI', 'SDETU5009SI']
BANDAS_MAIS50 = ['SDETU5002SI', 'SDETU5004SI', 'SDETU5006SI', 'SDETU5008SI', 'SDETU5010SI']


def _gt(a, b):
    """Comparação do Excel: texto é maior que qualquer número; "" (vazio de fórmula) também é texto."""
    if _err(a, b):
        raise _Propaga()
    if isinstance(a, str) and not isinstance(b, str):
        return True
    if not isinstance(a, str) and isinstance(b, str):
        return False
    return a > b


def _ge(a, b):
    return _gt(a, b) or a == b


def _le(a, b):
    return not _gt(a, b)


def _lt(a, b):
    return not _ge(a, b)


class _Propaga(Exception):
    pass


def arred(x, casas=2):
    """ARRED do Excel: meio para longe do zero, sobre o número com 15 algarismos significativos (0,185 -> 0,19)."""
    from decimal import Decimal, ROUND_HALF_UP
    return float(Decimal(format(x, '.15g')).quantize(Decimal(1).scaleb(-casas), rounding=ROUND_HALF_UP))


def _num(x):
    return x if isinstance(x, (int, float)) else 0.0


def banda_primeira(o5):
    """T3 (carga total O5, até 50 km). Repare no 4.001: entre 4 e 4,001 t não sai serviço."""
    if not _gt(o5, 0.0001):
        return ''
    if _gt(o5, 0.001) and _le(o5, 0.6): return 'SDETU5001SI'
    if _gt(o5, 0.6) and _le(o5, 4): return 'SDETU5003SI'
    if _gt(o5, 4.001) and _le(o5, 8): return 'SDETU5005SI'
    if _gt(o5, 8) and _le(o5, 14): return 'SDETU5007SI'
    if _gt(o5, 14) and _le(o5, 25): return 'SDETU5009SI'
    if _gt(o5, 25): return 'SDETU5009SI'
    return False                                    # o IF mais interno sem "senão" devolve FALSO


def banda_resto(v):
    """W3/Z3/AC3/...: bandas das cargas seguintes. Aqui 25 exato não cai em nenhuma (V3<25 / V3>25)."""
    if v == '':
        return ''
    if not _gt(v, 0.001):
        return ''
    if _gt(v, 0.001) and _le(v, 0.6): return 'SDETU5001SI'
    if _gt(v, 0.6) and _le(v, 4): return 'SDETU5003SI'
    if _gt(v, 4) and _le(v, 8): return 'SDETU5005SI'
    if _gt(v, 8) and _le(v, 14): return 'SDETU5007SI'
    if _gt(v, 14) and _lt(v, 25): return 'SDETU5009SI'
    if _gt(v, 25): return 'SDETU5009SI'
    return False


def banda_km(x, r3, cond_v=None):
    """T4/W4/...: serviço acima de 50 km para a carga x. `cond_v` = célula usada no último ramo (>25)."""
    if cond_v is None:
        cond_v = x
    if _le(x, 0.6) and _ge(r3, 50.001): return 'SDETU5002SI'
    if _ge(x, 0.6001) and _le(x, 4) and _ge(r3, 50.001): return 'SDETU5004SI'
    if _gt(x, 4) and _le(x, 8) and _ge(r3, 50.001): return 'SDETU5006SI'
    if _gt(x, 8) and _le(x, 14) and _ge(r3, 50.001): return 'SDETU5008SI'
    if _gt(x, 14) and _le(x, 25) and _ge(r3, 50.001): return 'SDETU5010SI'
    if _gt(cond_v, 25) and _gt(r3, 50): return 'SDETU5010SI'
    return ''


def calcular(qtd_odi, extra, km, pesos, h15_preenchido=None):
    if h15_preenchido is None:                      # sem km = campo MUNICÍPIO vazio na ICP (H15)
        h15_preenchido = km is not None
    # ---- O5: toneladas (H = E*G + F*G; O5 = soma/1000) ----
    h = 0.0
    for cod, _d, kg in pesos:
        h += _num(qtd_odi.get(str(cod), 0)) * kg + _num(extra.get(str(cod), 0)) * kg
    o5 = h / 1000
    # ---- R3, R4, O6 ----
    r3 = km if (h15_preenchido and km is not None) else '-'
    s1 = (r3 - 50) if h15_preenchido and not isinstance(r3, str) else ('' if not h15_preenchido else '#VALOR')
    r4 = '-' if not h15_preenchido else (s1 if _gt(r3, 50) else '')
    o6 = r3
    # ---- cadeia de cargas de 25 t: V3, Y3, AB3, AE3, AH3, AK3, (AN3) ----
    def resto(prev):
        if prev == '':
            return ''
        return (prev - 25) if _gt(prev, 25) else ''
    v3 = (o5 - 25) if _gt(o5, 25) else ''
    y3 = resto(v3); ab3 = resto(y3); ae3 = resto(ab3); ah3 = resto(ae3); ak3 = resto(ah3); an3 = resto(ak3)
    t3 = banda_primeira(o5)
    w3, z3, ac3, af3, ai3, al3 = (banda_resto(v3), banda_resto(y3), banda_resto(ab3), banda_resto(ae3), banda_resto(ah3), banda_resto(ak3))
    t4 = banda_km(o5, r3)
    w4 = banda_km(v3, r3) if v3 != '' else ''
    z4 = banda_km(y3, r3, cond_v=v3) if y3 != '' else ''          # Z4 usa V3 no último ramo
    ac4 = banda_km(ab3, r3, cond_v=y3) if ab3 != '' else ''       # AC4 usa Y3
    af4 = banda_km(ae3, r3, cond_v=ab3) if ae3 != '' else ''      # AF4 usa AB3
    ai4 = banda_km(ah3, r3, cond_v=ae3) if ah3 != '' else ''      # AI4 usa AE3
    al4 = banda_km(ak3, r3, cond_v=ak3) if ak3 != '' else ''
    # ---- O10 (ton x km excedente) ----
    if not h15_preenchido:
        o10 = '-'
    elif _lt(o6, 50):
        o10 = 'DISTÂNCIA INFERIOR A 50 Km'
    elif _gt(r3, 50) or _gt(o5, 25):
        o10 = o5 * r4 if not isinstance(r4, str) else ERRO
    else:
        o10 = ''
    # ---- quantidades por carga (coluna V) ----
    def load_qty(x, prox):
        if x == '':
            return ''
        if _gt(x, 0.001) and _lt(x, 25):
            return x
        if _gt(x, 25):
            return x - prox if prox != '' else ERRO
        return ''
    v13 = o5 if _le(o5, 25) else (o5 - v3)
    v15 = load_qty(v3, y3); v17 = load_qty(y3, ab3); v19 = load_qty(ab3, ae3); v21 = load_qty(ae3, ah3)
    v23 = load_qty(ah3, ak3); v25 = load_qty(ak3, an3)

    def mult(a, b):
        if isinstance(a, str) or isinstance(b, str):
            return ERRO                                  # texto ("" ou "-") x número = #VALOR!
        return a * b
    # T10, V10, Y10 ...
    if not h15_preenchido:
        t10 = v10 = y10 = ab10 = af10 = ai10 = al10 = ''
    else:
        t10 = '' if _lt(o6, 50) else (mult(v13, r4) if v3 != '' else '')
        v10 = '' if _lt(o6, 50) else (mult(v15, r4) if v3 != '' else '')
        y10 = mult(v17, r4) if y3 != '' else ''
        ab10 = mult(v19, r4) if ab3 != '' else ''
        af10 = mult(v21, r4) if ae3 != '' else ''
        ai10 = mult(v23, r4) if ah3 != '' else ''
        al10 = mult(v25, r4) if ak3 != '' else ''
    linhas = []                                     # (U, V) linhas 13..26
    linhas.append((t3 if t3 != '' else '', v13))
    u14 = t4 if (t4 != '' and _gt(o5, 0.001)) else ''
    v14 = ((o10 if _lt(o5, 25) else t10) if t4 != '' else '')
    linhas.append((u14, v14))
    linhas.append((w3 if v3 != '' else '', v15))
    u16 = w4 if (w4 != '' and v3 != '' and _gt(v3, 0.001)) else ''
    linhas.append((u16, v10 if w4 != '' else ''))
    linhas.append((z3 if z3 != '' else '', v17)); linhas.append((z4, y10 if z4 != '' else ''))
    linhas.append((ac3 if ac3 != '' else '', v19)); linhas.append((ac4, ab10 if ac4 != '' else ''))
    linhas.append((af3 if af3 != '' else '', v21)); linhas.append((af4, af10 if af4 != '' else ''))
    linhas.append((ai3 if ai3 != '' else '', v23)); linhas.append((ai4, ai10 if ai4 != '' else ''))
    linhas.append((al3 if al3 != '' else '', v25)); linhas.append((al4, al10 if al4 != '' else ''))
    # ---- AF: SOMASE por código (erro numa célula que casa = erro) ----
    af = {}
    for cod in BANDAS_ATE50 + BANDAS_MAIS50:
        casam = [v for u, v in linhas if u == cod]
        af[cod] = ERRO if _err(*casam) else sum(v for v in casam if isinstance(v, (int, float)))
    ac23 = ERRO if _err(*(af[c] for c in BANDAS_ATE50)) else sum(af[c] for c in BANDAS_ATE50)

    def primeiro(ordem):                              # IF(AF>0,001;cod;IF(…)) com curto-circuito
        for c in ordem:
            if _err(af[c]):
                return ERRO
            if af[c] > 0.001:
                return c
        return ''

    def e(*conds):                                   # E(…) do Excel: erro em qualquer argumento = erro
        vals = []
        for f in conds:
            try:
                vals.append(f())
            except _Propaga:
                return ERRO
        return all(vals)
    ak13 = primeiro(BANDAS_ATE50)
    ak14 = '' if not h15_preenchido else primeiro(BANDAS_MAIS50)
    c15 = e(lambda: _gt(ac23, 25), lambda: _gt(af['SDETU5009SI'], 0.001), lambda: _err(ak13) and (_ for _ in ()).throw(_Propaga()) or ak13 != 'SDETU5009SI')
    ak15 = ERRO if _err(c15) else ('SDETU5009SI' if c15 else '')
    k14 = ak13
    k15 = ERRO if _err(ak14) else (ak15 if ak14 == '' else ak14)
    k16 = ERRO if _err(ak14) else ('' if ak14 == '' else (ERRO if _err(c15) else ('SDETU5009SI' if c15 else '')))
    c17 = e(lambda: _err(ak15) and (_ for _ in ()).throw(_Propaga()) or ak15 == 'SDETU5009SI',
            lambda: _gt(af['SDETU5010SI'], 0.001),
            lambda: _err(ak13) and (_ for _ in ()).throw(_Propaga()) or ak13 != 'SDETU5010SI')
    k17 = ERRO if _err(c17) else ('SDETU5010SI' if c17 else '')
    out = []
    for k in (k14, k15, k16, k17):
        if _err(k):
            out.append((ERRO, ERRO))
        elif k:
            out.append((k, ERRO if _err(af[k]) else arred(af[k], 2)))
    return {'toneladas': o5, 'km': r3, 'servicos': out}


# ── dados da parte `transporte` ───────────────────────────────────────────────────────────────
def validar(v) -> list:
    """Erros da parte `transporte` (lista vazia = ok)."""
    if not isinstance(v, dict):
        return ["a parte transporte precisa ser um objeto JSON"]
    erros = []
    if v.get("formato") != FORMATO:
        erros.append(f"formato deve ser '{FORMATO}'")
    at = v.get("ativos")
    if not isinstance(at, dict) or not str(at.get("ate_50km") or "").strip() or not str(at.get("acima_50km") or "").strip():
        erros.append("ativos precisa de ate_50km e acima_50km")
    faixas = v.get("faixas")
    if not isinstance(faixas, list) or not faixas or not all(isinstance(f, dict) and f.get("faixa") and f.get("ate_50km") and f.get("acima_50km") for f in faixas):
        erros.append("faixas precisa de {faixa, ate_50km, acima_50km} em cada item")
    else:
        conhecidos = set(BANDAS_ATE50 + BANDAS_MAIS50)
        fora = [c for f in faixas for c in (f["ate_50km"], f["acima_50km"]) if c not in conhecidos]
        if fora:
            erros.append("serviço de faixa desconhecido: " + ", ".join(fora[:5]))
    pesos = v.get("pesos")
    if not isinstance(pesos, list) or not pesos:
        erros.append("pesos precisa de pelo menos um material")
    else:
        for i, p in enumerate(pesos, 1):
            if not isinstance(p, dict) or not str(p.get("codigo") or "").strip() or not isinstance(p.get("kg"), (int, float)) or isinstance(p.get("kg"), bool):
                erros.append(f"pesos, item {i}: precisa de codigo e kg numérico")
                break
    if not isinstance(v.get("km_por_utd", {}), dict):
        erros.append("km_por_utd precisa ser um objeto {UTD: km}")
    if not isinstance(v.get("municipio_utd", []), list):
        erros.append("municipio_utd precisa ser uma lista")
    return erros


def dados_do_projeto(projeto) -> dict:
    """A parte `transporte` do projeto (nome ou código): a importada no pacote ou, na falta, a do pacote de partida; {} se não há."""
    chave = str(projeto or "").strip()
    if not chave:
        return {}
    from services import pacote_projeto as pp
    proj = pp.projeto_do_cadastro(chave)
    v = (pp.parte_guardada(proj["codigo"], "transporte") if proj else None) or pp.parte_inicial(chave, "transporte")
    return v if isinstance(v, dict) and not validar(v) else {}


def _norm(t) -> str:
    from services.obra_params import normalizar
    return normalizar(t).replace("_", " ")


def resolver_utd(obra, dados) -> dict:
    """UTD da obra → {utd, km, origem}: `obra.utd`; senão a do bairro (1ª linha com o bairro); senão a do município. utd '' = não resolvida."""
    obra = obra or {}
    km_utd = {_norm(k): (k, v) for k, v in (dados.get("km_por_utd") or {}).items()}
    linhas = dados.get("municipio_utd") or []

    def com_km(utd, origem):
        k = km_utd.get(_norm(utd))
        return {"utd": k[0], "km": k[1], "origem": origem} if k else None
    if str(obra.get("utd") or "").strip():
        r = com_km(obra["utd"], "UTD da obra")
        if r:
            return r
    for campo, rotulo in (("bairro", "bairro"), ("municipio", "município")):
        alvo = _norm(obra.get(campo))
        if not alvo:
            continue
        for l in linhas:
            if _norm(l.get(campo)) == alvo and (campo != "bairro" or not _norm(obra.get("municipio")) or _norm(l.get("municipio")) == _norm(obra.get("municipio"))):
                r = com_km(l.get("utd"), f"{rotulo} {l.get(campo)}")
                if r:
                    return r
    return {"utd": "", "km": None, "origem": ""}


_RE_LINHA = re.compile(r"^\s*\*?[0-9]+(?:[.,][0-9]+)?[-xX]\s*([^\s(]+)")


def eh_linha_transporte(ativo, dados) -> bool:
    """A linha de Outros é um TRANSP/TRANSP_KM (de um cálculo anterior)?"""
    nomes = {str(v).upper() for v in (dados.get("ativos") or {}).values()}
    t = str(ativo or "").strip().upper()
    m = _RE_LINHA.match(t)
    nome = m.group(1) if m else t.split("(")[0].strip()
    return nome in nomes


def _qtd_texto(q) -> str:
    s = f"{q:.2f}".rstrip("0").rstrip(".")
    return s or "0"


def linhas_para_outros(servicos, dados) -> list:
    """[(texto da linha, erro?)] na ordem dos serviços: '37.6-TRANSP(faixa=25T)'; os #VALOR! vêm marcados e não são lançados."""
    ativos = dados.get("ativos") or {}
    por_codigo = {}
    for f in dados.get("faixas") or []:
        por_codigo[f["ate_50km"]] = (ativos.get("ate_50km"), f["faixa"])
        por_codigo[f["acima_50km"]] = (ativos.get("acima_50km"), f["faixa"])
    saida = []
    for cod, q in servicos:
        if _err(cod) or _err(q) or cod not in por_codigo:
            saida.append({"codigo": str(cod), "quantidade": None, "texto": "", "erro": True})
            continue
        ativo, faixa = por_codigo[cod]
        saida.append({"codigo": cod, "faixa": faixa, "quantidade": q, "texto": f"{_qtd_texto(q)}-{ativo}(faixa={faixa})", "erro": False})
    return saida
