"""
services/orcamento_arred.py — Arredondamento por linha da base de orçamento (TASK-107).

A coluna `arred` da base diz como a quantidade que ESTA linha gera é arredondada, a cada uso (quantidade do pedido x fator):
"NORMAL" (padrão: não arredonda cada uso; o total do código sai com 2 casas, como sempre), "INTEIRO" (arredondamento comum, 0 casas,
meio para cima), "CIMA" (próximo inteiro acima) e "BAIXO" (inteiro abaixo). Vazio/ausente = "NORMAL" (bases antigas continuam iguais).
"""
import math

VALORES_ARRED = ("NORMAL", "INTEIRO", "CIMA", "BAIXO")
_SINONIMOS = {"PARA CIMA": "CIMA", "PARA BAIXO": "BAIXO"}      # os nomes da Tabela de Regras da Totalizadora também valem na importação


def normalizar_arred(valor) -> str:
    """Normaliza o valor lido de CSV/tela/nuvem. Vazio/None = "NORMAL". Valor desconhecido → ValueError (a importação recusa)."""
    v = " ".join(str(valor if valor is not None else "").strip().upper().split())
    if not v:
        return "NORMAL"
    v = _SINONIMOS.get(v, v)
    if v not in VALORES_ARRED:
        raise ValueError(f"ARRED inválido '{str(valor).strip()}': use NORMAL, INTEIRO, CIMA ou BAIXO.")
    return v


def arred_tolerante(valor) -> str:
    """Para o cálculo: nunca falha — valor desconhecido na base vale como "NORMAL"."""
    try:
        return normalizar_arred(valor)
    except ValueError:
        return "NORMAL"


def aplicar_arred(x: float, modo) -> float:
    """Arredonda a quantidade de UM uso da linha. NORMAL devolve o valor intacto (o total é arredondado em 2 casas no fim)."""
    m = arred_tolerante(modo)
    if m == "NORMAL" or not math.isfinite(x):
        return x
    x = round(x, 9)                                 # 30 * 0,1 = 3,0000000000000004 não pode subir para 4
    if m == "CIMA":
        return float(math.ceil(x))
    if m == "BAIXO":
        return float(math.floor(x))
    return float(math.floor(x + 0.5))               # INTEIRO: meio para cima (como o Math.round do JavaScript)
