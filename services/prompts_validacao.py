"""
services/prompts_validacao.py — Leitura e validação dos prompts de validação (camada 3, ADR-004).

Um prompt é um texto único: cabeçalho entre linhas `---` (chave: valor) + corpo com placeholders
`{{NOME}}`. Sementes em data/validacoes/*.md; edição pelo admin via routers/validacao_prompts.py.
Função pura: sem I/O de banco.
"""
import os
import re

ESCOPOS_VALIDOS = {"cabos", "outros", "ambos"}
MODOS_VALIDOS = {"checar", "corrigir"}
SAIDAS_VALIDAS = {"json"}
# Único conjunto que o backend sabe preencher (TASK-014). Placeholder fora daqui nunca seria trocado.
PLACEHOLDERS_CONHECIDOS = {"LINHAS_CABOS", "LINHAS_OUTROS", "ACHADOS_PREVIOS", "ACHADOS"}
CAMPOS_OBRIGATORIOS = ("id", "escopo", "modo", "ordem", "modelo", "ativo", "saida")

_RE_PLACEHOLDER = re.compile(r"\{\{\s*([A-Z_]+)\s*\}\}")
_RE_ID = re.compile(r"^[a-z0-9][a-z0-9-]*$")


def parse_prompt(texto: str):
    """Devolve (meta: dict, corpo: str). Levanta ValueError se não houver cabeçalho bem formado."""
    linhas = (texto or "").replace("\r\n", "\n").split("\n")
    if not linhas or linhas[0].strip() != "---":
        raise ValueError("O prompt precisa começar com o cabeçalho '---'.")
    try:
        fim = next(i for i in range(1, len(linhas)) if linhas[i].strip() == "---")
    except StopIteration:
        raise ValueError("Cabeçalho sem fechamento: falta a segunda linha '---'.")
    meta = {}
    for n, linha in enumerate(linhas[1:fim], start=2):
        if not linha.strip():
            continue
        if ":" not in linha:
            raise ValueError(f"Cabeçalho, linha {n}: esperado 'chave: valor'.")
        chave, valor = linha.split(":", 1)
        meta[chave.strip()] = valor.strip()
    return meta, "\n".join(linhas[fim + 1:]).strip()


def validar_prompt(texto: str, prompt_id_esperado: str | None = None) -> list:
    """Lista de erros (vazia = válido). Nunca aceita o que o backend não conseguiria montar."""
    try:
        meta, corpo = parse_prompt(texto)
    except ValueError as e:
        return [str(e)]

    erros = []
    for campo in CAMPOS_OBRIGATORIOS:
        if not meta.get(campo):
            erros.append(f"Cabeçalho: campo '{campo}' é obrigatório.")
    if erros:
        return erros

    if not _RE_ID.match(meta["id"]):
        erros.append("Cabeçalho: 'id' deve usar só minúsculas, números e hífen.")
    if prompt_id_esperado and meta["id"] != prompt_id_esperado:
        erros.append(f"Cabeçalho: 'id' ({meta['id']}) não pode ser alterado (esperado: {prompt_id_esperado}).")
    if meta["escopo"] not in ESCOPOS_VALIDOS:
        erros.append(f"Cabeçalho: 'escopo' precisa ser um de {sorted(ESCOPOS_VALIDOS)}.")
    if meta["modo"] not in MODOS_VALIDOS:
        erros.append(f"Cabeçalho: 'modo' precisa ser um de {sorted(MODOS_VALIDOS)}.")
    if meta["saida"] not in SAIDAS_VALIDAS:
        erros.append(f"Cabeçalho: 'saida' precisa ser um de {sorted(SAIDAS_VALIDAS)}.")
    if meta["ativo"].lower() not in ("true", "false"):
        erros.append("Cabeçalho: 'ativo' precisa ser true ou false.")
    try:
        float(meta["ordem"])
    except ValueError:
        erros.append("Cabeçalho: 'ordem' precisa ser numérico.")
    if meta.get("temperatura"):
        try:
            if not 0 <= float(meta["temperatura"]) <= 2:
                erros.append("Cabeçalho: 'temperatura' precisa estar entre 0 e 2.")
        except ValueError:
            erros.append("Cabeçalho: 'temperatura' precisa ser numérico.")

    if not corpo:
        erros.append("O corpo do prompt está vazio.")
        return erros

    declarados = {p.strip() for p in meta.get("placeholders", "").split(",") if p.strip()}
    usados = set(_RE_PLACEHOLDER.findall(corpo))
    for p in sorted(declarados - PLACEHOLDERS_CONHECIDOS):
        erros.append(f"Cabeçalho: placeholder '{p}' desconhecido (válidos: {sorted(PLACEHOLDERS_CONHECIDOS)}).")
    for p in sorted(usados - PLACEHOLDERS_CONHECIDOS):
        erros.append(f"Corpo: '{{{{{p}}}}}' não é um placeholder conhecido e nunca seria preenchido.")
    for p in sorted(declarados - usados):
        erros.append(f"Placeholder '{p}' declarado no cabeçalho, mas não usado no corpo.")
    for p in sorted((usados & PLACEHOLDERS_CONHECIDOS) - declarados):
        erros.append(f"Placeholder '{p}' usado no corpo, mas não declarado no cabeçalho.")
    if "{{" in _RE_PLACEHOLDER.sub("", corpo):
        erros.append("Corpo: há '{{' que não forma um placeholder válido (ex.: {{LINHAS_CABOS}}).")
    return erros


def resumo_meta(texto: str) -> dict:
    """Metadados para listagem; em prompt inválido devolve o que der para ler."""
    try:
        meta, _ = parse_prompt(texto)
    except ValueError:
        return {}
    return {
        "id": meta.get("id"),
        "escopo": meta.get("escopo"),
        "modo": meta.get("modo"),
        "ordem": meta.get("ordem"),
        "modelo": meta.get("modelo"),
        "ativo": (meta.get("ativo") or "").lower() == "true",
    }


def ler_sementes(diretorio: str):
    """Devolve ({prompt_id: texto}, [(arquivo, motivo)]) dos .md do diretório; inválido é ignorado."""
    sementes, ignorados = {}, []
    if not os.path.isdir(diretorio):
        return sementes, ignorados
    for nome in sorted(os.listdir(diretorio)):
        if not nome.endswith(".md"):
            continue
        with open(os.path.join(diretorio, nome), "r", encoding="utf-8") as f:
            texto = f.read()
        erros = validar_prompt(texto)
        if erros:
            ignorados.append((nome, erros[0]))
            continue
        sementes[parse_prompt(texto)[0]["id"]] = texto
    return sementes, ignorados
