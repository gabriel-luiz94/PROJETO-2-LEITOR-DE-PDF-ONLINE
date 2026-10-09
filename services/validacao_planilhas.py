"""
services/validacao_planilhas.py — Camada 1 da validação das planilhas Cabos e Outros (contrato).

Função pura: sem I/O e sem IA. Regras NÃO editáveis — são os formatos que o cálculo lê
(RULES Regra 5; ADR-004; catálogo em .ai/VALIDACOES.md).

O parser do backend é tolerante e falha em silêncio (assume 1, descarta token sobrando); estas
checagens expõem esses casos antes do cálculo.
"""
import re

from services.orcamento_calc import (
    extrair_ativo_cabo,
    extrair_ativos_outros,
    processar_calculo,
    tokenizar_outros,
)

OPERACOES_VALIDAS = {"I", "*I", "R", "*R", "M", "*M"}

_RE_QTD_HIFEN_ATIVO = re.compile(r"^\*?\d+([.,]\d+)?-\S")
_RE_TERMINA_EM_METROS = re.compile(r"\d+([.,]\d+)?\s*M\s*$", re.IGNORECASE)
_RE_NUMERO = re.compile(r"^[\d.,]+$")

# ATENÇÃO (Regra 4): a normalização abaixo espelha a de static/resumo.js
# (calcularQtdAtivos / extrairPrefixoEComprimentoCabo). A tabela Cabos da tela usa "CAA 2 ABC 35 m"
# (com espaço no nome), formato que o parser do backend NÃO lê — o backend só recebe o payload já
# convertido ("CAA2 1 35"). Se a normalização do frontend mudar, ajustar aqui também.
_PREFIXOS_CABO = (
    (re.compile(r"^CAA\s+(\d)", re.IGNORECASE), r"CAA\1"),
    (re.compile(r"^CA\s+(\d)", re.IGNORECASE), r"CA\1"),
    (re.compile(r"^CU\s+(\d)", re.IGNORECASE), r"CU\1"),
    (re.compile(r"^CAZ\s+(\d)", re.IGNORECASE), r"CAZ\1"),
    (re.compile(r"^P\s+(\d)", re.IGNORECASE), r"P\1"),
)


def _achado(linha_id, tabela, regra_id, severidade, mensagem):
    return {
        "linha_id": linha_id,
        "tabela": tabela,
        "regra_id": regra_id,
        "severidade": severidade,
        "mensagem": mensagem,
    }


def _para_float(token: str):
    try:
        return float(token.replace(",", "."))
    except ValueError:
        return None


def _tokens_cabo_tabela(ativo: str) -> list:
    """Tokens de uma linha da tabela Cabos, com a mesma normalização do frontend."""
    txt = ativo.strip().upper().replace("/", "")
    for regex, repl in _PREFIXOS_CABO:
        txt = regex.sub(repl, txt)
    txt = re.sub(r"\s+M\s*$", "", txt).strip()
    return txt.split()


def _validar_operacao(linha, linha_id, tabela, achados):
    operacao = (linha.get("operacao") or "").strip().upper()
    if operacao not in OPERACOES_VALIDAS:
        achados.append(_achado(
            linha_id, tabela, "C1-OP", "erro",
            f"Operação '{linha.get('operacao') or ''}' inválida. Use I, *I, R, *R, M ou *M.",
        ))


def _variavel_sem_valor(ativo, tabela, linha_id, achados) -> bool:
    """TASK-058: `V` de modelo que sobrou na linha (modelo usado sem gerar a obra)."""
    from services.modelos_obra import tem_variavel   # import tardio: modelos_obra usa esta camada
    if not tem_variavel(ativo, tabela):
        return False
    achados.append(_achado(linha_id, tabela, "C1-VAR", "erro",
                           f"'{ativo}' ainda tem variável de modelo (V ou X(...)) sem valor. Gere a obra pelo modelo ou troque a variável por um valor."))
    return True


def _validar_linha_cabo(linha, linha_id, achados):
    ativo = (linha.get("ativo") or "").strip()
    if not ativo:
        return  # o cálculo ignora linha sem ativo
    if _variavel_sem_valor(ativo, "cabos", linha_id, achados):
        return
    _validar_operacao(linha, linha_id, "cabos", achados)

    if _RE_QTD_HIFEN_ATIVO.match(ativo):
        achados.append(_achado(
            linha_id, "cabos", "C1-CABO-PARTE", "erro",
            f"'{ativo}' está no formato <qtd>-<ativo> (Outros). Cabos usa [ATIVO] [FASE] [COMPRIMENTO] m.",
        ))
        return

    tokens = _tokens_cabo_tabela(ativo)
    if not tokens:
        return
    if len(tokens) == 1:
        return  # linha standalone: herda a fase da linha seguinte (RN-04)

    ultimo = _para_float(tokens[-1]) if _RE_NUMERO.match(tokens[-1]) else None
    if ultimo is None:
        achados.append(_achado(
            linha_id, "cabos", "C1-CABO-NUM", "erro",
            f"Comprimento não numérico em '{ativo}'. Formato: [ATIVO] [FASE] [COMPRIMENTO] m.",
        ))
    elif ultimo <= 0:
        achados.append(_achado(
            linha_id, "cabos", "C1-QTD-ZERO", "aviso",
            f"Comprimento zero em '{ativo}': a linha não entra no orçamento.",
        ))


def _validar_linha_outros(linha, linha_id, achados):
    ativo = (linha.get("ativo") or "").strip()
    if not ativo:
        return
    if _variavel_sem_valor(ativo, "outros", linha_id, achados):
        return
    _validar_operacao(linha, linha_id, "outros", achados)

    if _RE_TERMINA_EM_METROS.search(ativo):
        achados.append(_achado(
            linha_id, "outros", "C1-OUT-CABO", "erro",
            f"'{ativo}' termina em metros (formato de cabo). Outros usa <qtd>-<ativo>.",
        ))
        return

    tokens = tokenizar_outros(ativo)
    if len(tokens) % 2 == 1:
        achados.append(_achado(
            linha_id, "outros", "C1-OUT-ORFAO", "erro",
            f"Token sem par em '{ativo}': '{tokens[-1]}' seria descartado no cálculo. Use <qtd>-<ativo>.",
        ))

    for i in range(0, len(tokens) - 1, 2):
        if _para_float(tokens[i]) is None:
            achados.append(_achado(
                linha_id, "outros", "C1-OUT-NUM", "erro",
                f"Quantidade '{tokens[i]}' não numérica antes de '{tokens[i + 1]}': o cálculo assumiria 1.",
            ))

    itens = extrair_ativos_outros({"ativo": ativo, "operacao": ""})
    vistos = {}
    for item in itens:
        vistos[item["ativo"]] = vistos.get(item["ativo"], 0) + 1
        if item["qtd"] == 0:
            achados.append(_achado(
                linha_id, "outros", "C1-QTD-ZERO", "aviso",
                f"Quantidade zero para '{item['ativo']}': o item não entra no orçamento.",
            ))
    for nome, vezes in vistos.items():
        if vezes > 1:
            achados.append(_achado(
                linha_id, "outros", "C1-DUP", "aviso",
                f"'{nome}' aparece {vezes} vezes na mesma linha; some as quantidades.",
            ))


def validar_planilhas(cabos: list, outros: list) -> list:
    """Camada 1 sobre as tabelas Cabos e Outros como aparecem na tela.

    Cada linha: {"id": opcional, "ativo": str, "operacao": str}. Sem `id`, usa CABOS-<i>/OUTROS-<i>
    (índice de 0, o mesmo do contexto enviado ao chat).
    """
    achados = []
    for i, linha in enumerate(cabos or []):
        _validar_linha_cabo(linha, linha.get("id") or f"CABOS-{i}", achados)
    for i, linha in enumerate(outros or []):
        _validar_linha_outros(linha, linha.get("id") or f"OUTROS-{i}", achados)
    return achados


def validar_base(cabos_payload: list, outros_payload: list, projeto, base_rows: list) -> list:
    """C1-BASE sobre o payload de cálculo (Totalizadora, já após as regras de conversão).

    Roda depois das regras de conversão (ADR-003) de propósito: uma regra pode gerar ou trocar um
    ativo, então checar a tabela crua daria falso positivo. Usa a mesma cascata RN-06 do cálculo.
    """
    achados = []
    for tabela, linhas, extrair, chamar in (
        ("cabos", cabos_payload or [], lambda it: [a for a in [extrair_ativo_cabo(it)] if a],
         lambda itens: processar_calculo(itens, [], projeto, base_rows)),
        ("outros", outros_payload or [], extrair_ativos_outros,
         lambda itens: processar_calculo([], itens, projeto, base_rows)),
    ):
        nao_encontrados = set(chamar(linhas)["nao_encontrados"])
        if not nao_encontrados:
            continue
        for i, linha in enumerate(linhas):
            linha_id = linha.get("id") or f"{tabela.upper()}-{i}"
            for av in extrair(linha):
                if av["ativo"] in nao_encontrados:
                    achados.append(_achado(
                        linha_id, tabela, "C1-BASE", "aviso",
                        f"Ativo '{av['ativo']}' não encontrado na base técnica"
                        f"{' (projeto ' + str(projeto) + ')' if projeto else ''}.",
                    ))
    return achados


def resumir(achados: list) -> dict:
    resumo = {"erro": 0, "aviso": 0, "info": 0}
    for a in achados:
        resumo[a["severidade"]] = resumo.get(a["severidade"], 0) + 1
    return resumo
