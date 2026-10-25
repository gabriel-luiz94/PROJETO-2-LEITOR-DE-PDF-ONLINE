"""
services/orcamento_lv.py — Linha viva x linha morta na base de orçamento (TASK-078).

A coluna `lv` da base tem três valores: "LV" (só para operações em linha viva, as que começam com `*`: *I, *R, *M),
"LM" (só para operações sem `*`) e "A" (ambas). Vazio/ausente = "A" (bases antigas continuam iguais).
"""
VALORES_LV = ("LV", "LM", "A")


def normalizar_lv(valor) -> str:
    """Normaliza o valor lido de CSV/tela/nuvem. Vazio/None = "A". Valor desconhecido → ValueError (a importação recusa)."""
    v = str(valor if valor is not None else "").strip().upper()
    if not v:
        return "A"
    if v not in VALORES_LV:
        raise ValueError(f"LV inválido '{str(valor).strip()}': use LV, LM ou A.")
    return v


def lv_tolerante(valor) -> str:
    """Para o cálculo: nunca falha — valor desconhecido na base vale como "A"."""
    try:
        return normalizar_lv(valor)
    except ValueError:
        return "A"


def operacao_viva(operacao) -> bool:
    return str(operacao or "").strip().startswith("*")


def passa_lv(lv_linha, operacao) -> bool:
    """A linha da base (lv) vale para esta operação? A: sempre; LV: só com `*`; LM: só sem `*`."""
    lv = lv_tolerante(lv_linha)
    if lv == "A":
        return True
    return operacao_viva(operacao) == (lv == "LV")
