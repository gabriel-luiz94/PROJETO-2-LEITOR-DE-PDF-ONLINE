"""
services/obras_contexto.py — Contexto das obras salvas para a IA do chat (TASK-028).

Funções puras (sem banco/rede). O contexto só é montado quando o pedido menciona obras (`menciona_obra`), para que as
conversas normais fiquem exatamente como eram (mesmos tokens, mesma latência). O índice é enxuto (id | nome | data, sem
`dados_json`); o conteúdo de uma obra só vai quando o usuário pergunta por ela, e só dela, com teto de linhas.
Não registrar conteúdo de obras em log (RULES Regra 11).
"""
import json
import re
import unicodedata

LIMITE_INDICE = 50          # obras listadas no índice
LIMITE_OBRAS_CONTEUDO = 2   # obras cujo conteúdo vai junto
LIMITE_LINHAS = 40          # linhas por tabela no resumo de uma obra

_RE_OBRA = re.compile(r"\bobras?\b")
_RE_PEDE_CONTEUDO = re.compile(
    r"\b(o que (tem|ha|contem)|conteudo|detalh\w*|mostre|mostra|compar\w*|quantos|quantas|resum\w*|descreva|liste os? (itens|ativos|postes))\b")


def normalizar(texto: str) -> str:
    """Minúsculas, sem acentos e com espaços simples — para comparar nomes e palavras-chave."""
    sem = unicodedata.normalize("NFD", texto or "")
    sem = "".join(c for c in sem if unicodedata.category(c) != "Mn")
    return re.sub(r"\s+", " ", sem.lower()).strip()


def menciona_obra(prompt: str) -> bool:
    """O pedido fala de obras? (palavra "obra"/"obras", sem distinguir acentos/maiúsculas)."""
    return bool(_RE_OBRA.search(normalizar(prompt)))


def montar_indice(obras: list) -> str:
    """Índice enxuto das obras (já na ordem desejada, mais recentes primeiro)."""
    cab = "OBRAS SALVAS NESTE PROJETO (mais recentes primeiro):"
    if not obras:
        return cab + "\n(nenhuma obra salva neste projeto)"
    linhas = [f'- id={o["id"]} | nome="{o["nome"]}" | data={o.get("data", "")}' for o in obras[:LIMITE_INDICE]]
    if len(obras) > LIMITE_INDICE:
        linhas.append(f"(+{len(obras) - LIMITE_INDICE} obras mais antigas não listadas)")
    return cab + "\n" + "\n".join(linhas)


def obras_citadas(prompt: str, obras: list) -> list:
    """Obras cujo nome aparece no pedido, quando o pedido pede o CONTEÚDO (não para comandos de adicionar/subtrair/carregar)."""
    p = normalizar(prompt)
    if not _RE_PEDE_CONTEUDO.search(p):
        return []
    achadas = [o for o in obras if len(normalizar(o["nome"])) >= 2 and normalizar(o["nome"]) in p]
    achadas.sort(key=lambda o: -len(normalizar(o["nome"])))   # o nome mais específico primeiro ("Obra 2" antes de "Obra")
    return achadas[:LIMITE_OBRAS_CONTEUDO]


def _linhas(tabela) -> list:
    dados = tabela.get("data") if isinstance(tabela, dict) else None
    return [r for r in (dados or []) if isinstance(r, dict)]


def contar_obra(dados_json) -> tuple:
    """(linhas em Cabos, linhas em Outros) de um `dados_json` salvo; (0, 0) se ilegível."""
    try:
        snap = json.loads(dados_json) if isinstance(dados_json, str) else (dados_json or {})
    except (TypeError, ValueError):
        return 0, 0
    if not isinstance(snap, dict):
        return 0, 0
    return len(_linhas(snap.get("cabos"))), len(_linhas(snap.get("outros")))


def resumir_obra(obra: dict) -> str:
    """Texto com o conteúdo de UMA obra (teto de linhas por tabela)."""
    try:
        snap = json.loads(obra.get("dados_json") or "{}")
    except (TypeError, ValueError):
        snap = {}
    saida = [f'CONTEÚDO DA OBRA "{obra["nome"]}" (id={obra["id"]}):']
    for rotulo, chave in (("CABOS", "cabos"), ("OUTROS", "outros")):
        linhas = _linhas(snap.get(chave) if isinstance(snap, dict) else None)
        saida.append(f"{rotulo} ({len(linhas)} linhas):")
        saida += [f"  {(r.get('operacao') or '-')} | {r.get('ativo', '')}" for r in linhas[:LIMITE_LINHAS]]
        if len(linhas) > LIMITE_LINHAS:
            saida.append(f"  (+{len(linhas) - LIMITE_LINHAS} linhas não mostradas)")
    return "\n".join(saida)
