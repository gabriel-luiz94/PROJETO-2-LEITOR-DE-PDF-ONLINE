"""Cálculo do transporte de material (aba TRANSPORT da ICP V1.62), reproduzindo as fórmulas célula a célula.

Entrada:
  qtd_odi: {codigo_material: quantidade na ODI (instalação)}
  extra:   {codigo_material: quantidade extra digitada (coluna F)}
  km:      KMx2 do UTD da obra (MUNICÍPIO_UTD!G, 1ª linha do UTD) ou None quando o município/depósito não foi informado
  pesos:   lista [(codigo, descricao, kg)] na ordem da aba (linhas 6..59, com as repetições)
Saída: lista [(codigo_servico, quantidade arredondada 2 casas)] na ordem das linhas 14..17 da aba.
"""

BANDAS_ATE50 = ['SDETU5001SI', 'SDETU5003SI', 'SDETU5005SI', 'SDETU5007SI', 'SDETU5009SI']
BANDAS_MAIS50 = ['SDETU5002SI', 'SDETU5004SI', 'SDETU5006SI', 'SDETU5008SI', 'SDETU5010SI']


def _gt(a, b):
    """Comparação do Excel: texto é maior que qualquer número; "" (vazio de fórmula) também é texto."""
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


def calcular(qtd_odi, extra, km, pesos, h15_preenchido=True):
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
        o10 = o5 * r4 if not isinstance(r4, str) else '#VALOR'
    else:
        o10 = ''
    # ---- quantidades por carga (coluna V) ----
    def load_qty(x, prox):
        if x == '':
            return ''
        if _gt(x, 0.001) and _lt(x, 25):
            return x
        if _gt(x, 25):
            return x - prox if prox != '' else '#VALOR'
        return ''
    v13 = o5 if _le(o5, 25) else (o5 - v3)
    v15 = load_qty(v3, y3); v17 = load_qty(y3, ab3); v19 = load_qty(ab3, ae3); v21 = load_qty(ae3, ah3)
    v23 = load_qty(ah3, ak3); v25 = load_qty(ak3, an3)

    def mult(a, b):
        if isinstance(a, str) or isinstance(b, str):
            return '#VALOR' if (a != '' and b != '') else '#VALOR'
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
    # ---- AF: SUMIF por código (SUMIF ignora texto) ----
    af = {}
    for cod in BANDAS_ATE50 + BANDAS_MAIS50:
        af[cod] = sum(v for u, v in linhas if u == cod and isinstance(v, (int, float)))
    ac23 = sum(af[c] for c in BANDAS_ATE50)
    # ---- AK13..AK16 e saída K14..K17 ----
    ak13 = next((c for c in BANDAS_ATE50 if af[c] > 0.001), '')
    ak14 = '' if not h15_preenchido else next((c for c in BANDAS_MAIS50 if af[c] > 0.001), '')
    ak15 = 'SDETU5009SI' if (ac23 > 25 and af['SDETU5009SI'] > 0.001 and ak13 != 'SDETU5009SI') else ''
    k14 = ak13
    k15 = ak15 if ak14 == '' else ak14
    k16 = '' if ak14 == '' else ('SDETU5009SI' if (ac23 > 25 and af['SDETU5009SI'] > 0.001 and ak13 != 'SDETU5009SI') else '')
    k17 = 'SDETU5010SI' if (ak15 == 'SDETU5009SI' and af['SDETU5010SI'] > 0.001 and ak13 != 'SDETU5010SI') else ''
    out = []
    for k in (k14, k15, k16, k17):
        if k:
            out.append((k, round(af[k] + 0.0, 2)))
    return {'toneladas': o5, 'km': r3, 'servicos': out}
