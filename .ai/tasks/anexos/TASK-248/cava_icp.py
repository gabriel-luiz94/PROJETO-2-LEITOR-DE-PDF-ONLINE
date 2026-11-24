"""Cava em rocha (aba CAVA_ROCHA da ICP V1.62), célula a célula.

ICP:
  CAVA_ROCHA!F<r> = SE(A<r> = DASHBORD!C4; Mat_Base_2!G2007 + Mat_Base_2!G2068; 0)
      G2007 = quantidade de SDEMU3002II (instalar estai em solo) na ODI
      G2068 = quantidade de SDETU5011SI = SOMA(G1995:G1996) = SDEMU1004II + SDEMU1005II na ODI
  CAVA_ROCHA!G<r> = SE(ARREDMULTB(MARRED(F*B; 0,1); 1) > F; F; ARREDMULTB(MARRED(F*B; 0,1); 1))
  Mat_Base_2!G2066 (SDEMU3007II CAVA EM ROCHA) = SOMA(CAVA_ROCHA!G3:G123)
"""
from decimal import Decimal, ROUND_HALF_UP, ROUND_FLOOR


def _d(x):
    return Decimal(format(x, '.15g'))                    # o Excel trabalha com 15 algarismos significativos


def mround(x, m):
    """MARRED: múltiplo de m mais próximo, meio para longe do zero."""
    if x == 0:
        return 0.0
    q = (_d(x) / _d(m)).quantize(Decimal(1), rounding=ROUND_HALF_UP)
    return float(q * _d(m))


def floor1(x):
    """ARREDMULTB(x; 1) para x >= 0."""
    return float(_d(x).quantize(Decimal(1), rounding=ROUND_FLOOR))


def cavas(postes_estais, indice):
    """Quantidade de cavas em rocha de UMA localidade (linha da aba)."""
    f = postes_estais
    g = floor1(mround(f * indice, 0.1))
    return f if g > f else g


def calcular(localidade, qtd_servicos, tabela):
    """localidade = DASHBORD!C4; qtd_servicos = {codigo: quantidade na ODI (instalação)}; tabela = [(localidade, indice), ...] na ordem da aba.
    Devolve a quantidade de SDEMU3007II (0 quando a localidade não está na tabela)."""
    f = (qtd_servicos.get('SDEMU3002II', 0) or 0) + (qtd_servicos.get('SDEMU1004II', 0) or 0) + (qtd_servicos.get('SDEMU1005II', 0) or 0)
    total = 0.0
    for loc, indice in tabela:
        if loc == localidade:                            # a aba soma todas as linhas iguais à localidade (na tabela não há repetidas)
            total += cavas(f, indice)
    return total
