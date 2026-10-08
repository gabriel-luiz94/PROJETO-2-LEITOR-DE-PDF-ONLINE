"""
services/validacao_ia.py — Camada 3 da validação (ADR-004): revisão das planilhas por IA com prompt salvo.

Só monta o prompt, divide em lotes e interpreta a resposta; a chamada ao modelo é injetada
(`chamar`), o que mantém este módulo testável sem rede. A IA só aponta o que o código não verifica
(coerência semântica) — recebe os achados das camadas 1 e 2 para não repeti-los — e NUNCA corrige:
correção é outro prompt (`corrigir-planilhas`, TASK-015) e sempre exige aceite do usuário.

Não registrar o conteúdo das planilhas em log (RULES Regra 11).
"""
import asyncio
import json
import re

from services.prompts_validacao import parse_prompt

SEVERIDADES = {"erro", "aviso", "info"}
# Mesma ordem de reserva do chat (routers/ai_chat.py); o modelo do cabeçalho do prompt vem primeiro
# (só para Gemini — o campo "modelo" dos prompts salvos sempre guarda um ID Gemini; ver
# routers/validacao.py:_preparar_ia). gemini-1.5-flash/pro: desligados desde 2025 (404).
# gemini-2.5-flash: desligamento anunciado pela Google para 16/10/2026 — removidos (TASK-054).
MODELOS_RESERVA = ["gemini-3.1-flash-lite", "gemini-3.6-flash", "gemini-3.5-flash"]
# Claude com claude-haiku-5-5 primeiro (rápido/barato, padrão pedido pelo usuário) e
# claude-sonnet-5-5 como reserva de qualidade (TASK-055).
MODELOS_RESERVA_OPENAI = ["gpt-4o-mini", "gpt-4o"]
MODELOS_RESERVA_CLAUDE = ["claude-haiku-5-5", "claude-sonnet-5-5"]
# "invalid_argument"/"multiturn" cobrem o erro "Multiturn chat is not enabled for this model" (e
# variações parecidas) como defesa em profundidade (TASK-054) — `chamar_gemini` já usa chamada
# única (sem sessão), então não tem a causa raiz desse erro, mas um modelo futuro pode ter outra
# restrição parecida.
_ERROS_DE_FALLBACK = ("429", "quota", "exhausted", "not found", "404", "unavailable", "invalid_argument", "multiturn")
TIMEOUT_S = 60
TAMANHO_LOTE = 40      # linhas por chamada ao modelo
LIMITE_LINHAS = 400    # acima disso a IA revisa só as primeiras (resposta indica `truncado`)

_RE_CERCA = re.compile(r"^```(?:json)?\s*|\s*```$", re.IGNORECASE)


def _linhas_validas(linhas, prefixo):
    """[(id, operacao, ativo)] das linhas com ativo, com o id que a tela usa (CABOS-<i>/OUTROS-<i>)."""
    saida = []
    for i, linha in enumerate(linhas or []):
        ativo = (linha.get("ativo") or "").strip()
        if ativo:
            saida.append((linha.get("id") or f"{prefixo}-{i}", (linha.get("operacao") or "").strip(), ativo))
    return saida


def dividir_em_lotes(cabos, outros, tamanho=TAMANHO_LOTE, limite=LIMITE_LINHAS):
    """Devolve (lotes, truncado). Cada lote: {"cabos": [...], "outros": [...]} de (id, operacao, ativo)."""
    todas = [("cabos", x) for x in _linhas_validas(cabos, "CABOS")] + \
            [("outros", x) for x in _linhas_validas(outros, "OUTROS")]
    truncado = len(todas) > limite
    todas = todas[:limite]
    lotes = []
    for inicio in range(0, len(todas), tamanho):
        lote = {"cabos": [], "outros": []}
        for tabela, item in todas[inicio:inicio + tamanho]:
            lote[tabela].append(item)
        lotes.append(lote)
    return lotes, truncado


def _formatar(itens):
    return "\n".join(f"[{i}] | {op or '-'} | {ativo}" for i, op, ativo in itens) or "(nenhuma linha)"


def _formatar_achados(achados, ids_do_lote):
    linhas = [f"{a['linha_id']} | {a['regra_id']} | {a['mensagem']}"
              + (f" | sugestão: {a['sugestao']}" if a.get("sugestao") else "")
              for a in achados or [] if a.get("linha_id") in ids_do_lote or a.get("linha_id") == "GERAL"]
    return "\n".join(linhas) or "(nenhum)"


def montar_prompt(texto_prompt, lote, achados_previos):
    """Troca os placeholders do corpo do prompt salvo. Levanta ValueError se o prompt for ilegível."""
    _, corpo = parse_prompt(texto_prompt)
    ids = {i for i, _, _ in lote["cabos"] + lote["outros"]}
    valores = {
        "LINHAS_CABOS": _formatar(lote["cabos"]),
        "LINHAS_OUTROS": _formatar(lote["outros"]),
        "ACHADOS_PREVIOS": _formatar_achados(achados_previos, ids),
        "ACHADOS": _formatar_achados(achados_previos, ids),
    }
    for nome, valor in valores.items():
        corpo = re.sub(r"\{\{\s*" + nome + r"\s*\}\}", lambda _m, v=valor: v, corpo)
    return corpo


def interpretar_resposta(texto, ids_por_tabela):
    """(achados, descartados: lista de motivos). Tolera cerca ```json e descarta item malformado ou com linha_id inventado.

    ids_por_tabela: {"cabos": {ids}, "outros": {ids}} do lote enviado — a IA só pode apontar linhas que recebeu.
    Levanta ValueError se a resposta inteira não for um JSON com a chave "achados" em lista.
    """
    limpo = _RE_CERCA.sub("", (texto or "").strip())
    try:
        dados = json.loads(limpo)
    except json.JSONDecodeError as e:
        raise ValueError(f"resposta da IA não é JSON válido ({e.msg})")
    itens = dados.get("achados") if isinstance(dados, dict) else None
    if not isinstance(itens, list):
        raise ValueError('a resposta da IA não tem a lista "achados"')

    achados, descartados = [], []
    for item in itens:
        if not isinstance(item, dict):
            descartados.append({"linha_id": None, "motivo": "item não é um objeto"})
            continue
        linha_id = item.get("linha_id")
        tabela = next((t for t, ids in ids_por_tabela.items() if linha_id in ids), None)
        problema = item.get("problema")
        if tabela is None or not isinstance(problema, str) or not problema.strip():
            motivo = "linha não enviada à IA" if tabela is None else "sem descrição do problema"
            descartados.append({"linha_id": linha_id if isinstance(linha_id, str) else None, "motivo": motivo})
            continue
        severidade = item.get("severidade") if item.get("severidade") in SEVERIDADES else "info"
        regra = item.get("regra") if isinstance(item.get("regra"), str) and item["regra"].strip() else "ia"
        achado = {
            "linha_id": linha_id, "tabela": tabela, "regra_id": f"IA:{regra.strip()}",
            "severidade": severidade, "mensagem": problema.strip(), "origem": "ia",
        }
        if isinstance(item.get("sugestao"), str) and item["sugestao"].strip():
            achado["sugestao"] = item["sugestao"].strip()
        achados.append(achado)
    return achados, descartados


async def executar_em_lotes(chamar, texto_prompt, cabos, outros, achados, tratar):
    """Laço comum da revisão e da correção: divide em lotes, monta o prompt, chama o modelo e passa a resposta a
    `tratar(resposta, lote) -> (itens, descartados)`. Nunca levanta: falha de um lote (JSON ilegível, rede, cota,
    modelo) vira mensagem e não derruba os demais nem as outras camadas da validação.

    Devolve {"status": "ok"|"parcial"|"erro", "itens": [...], "descartados": [...], "mensagem": str,
             "lotes": n, "truncado": bool}. "parcial" = alguns lotes falharam.
    """
    lotes, truncado = dividir_em_lotes(cabos, outros)
    resultado = {"status": "ok", "itens": [], "descartados": [], "mensagem": "", "lotes": len(lotes),
                 "truncado": truncado}
    falhas = []
    for n, lote in enumerate(lotes, start=1):
        try:
            resposta = await chamar(montar_prompt(texto_prompt, lote, achados))
            itens, descartados = tratar(resposta, lote)
            resultado["itens"] += itens
            resultado["descartados"] += descartados
        except Exception as e:
            falhas.append(f"lote {n}: {str(e)[:200]}")
    if falhas:
        resultado["status"] = "erro" if len(falhas) == len(lotes) else "parcial"
        resultado["mensagem"] = "; ".join(falhas[:3])
    return resultado


async def validar_com_ia(chamar, texto_prompt, cabos, outros, achados_previos):
    """Revisa as planilhas. `chamar(texto) -> str` é assíncrona.

    Resposta: {"status", "achados": [...], "mensagem", "lotes", "truncado", "descartados": n}.
    """
    def tratar(resposta, lote):
        return interpretar_resposta(resposta, {
            "cabos": {i for i, _, _ in lote["cabos"]}, "outros": {i for i, _, _ in lote["outros"]},
        })

    r = await executar_em_lotes(chamar, texto_prompt, cabos, outros, achados_previos, tratar)
    return {"status": r["status"], "achados": r["itens"], "mensagem": r["mensagem"], "lotes": r["lotes"],
            "truncado": r["truncado"], "descartados": len(r["descartados"])}


async def chamar_gemini(api_key, modelos, texto, temperatura=0.0):
    """Uma chamada não-streaming pedindo JSON. Tenta os modelos em ordem só para erros de cota/indisponibilidade."""
    from google import genai
    from google.genai import types
    client = genai.Client(api_key=api_key)
    ultimo = None
    for modelo in modelos:
        try:
            resposta = await asyncio.wait_for(
                client.aio.models.generate_content(
                    model=modelo, contents=texto,
                    config=types.GenerateContentConfig(temperature=temperatura, response_mime_type="application/json"),
                ),
                timeout=TIMEOUT_S,
            )
            return resposta.text or ""
        except Exception as e:
            if any(k in str(e).lower() for k in _ERROS_DE_FALLBACK):
                ultimo = e
                continue
            raise
    raise ultimo or RuntimeError("Nenhum modelo Gemini disponível.")


async def chamar_openai(api_key, modelos, texto, temperatura=0.0):
    """Mesmo contrato de `chamar_gemini` (TASK-055), usando o modo JSON nativo da OpenAI."""
    from openai import AsyncOpenAI
    client = AsyncOpenAI(api_key=api_key)
    ultimo = None
    for modelo in modelos:
        try:
            resposta = await asyncio.wait_for(
                client.chat.completions.create(
                    model=modelo, messages=[{"role": "user", "content": texto}],
                    temperature=temperatura, response_format={"type": "json_object"},
                ),
                timeout=TIMEOUT_S,
            )
            return resposta.choices[0].message.content or ""
        except Exception as e:
            if any(k in str(e).lower() for k in _ERROS_DE_FALLBACK):
                ultimo = e
                continue
            raise
    raise ultimo or RuntimeError("Nenhum modelo OpenAI disponível.")


async def chamar_claude(api_key, modelos, texto, temperatura=0.0):
    """Mesmo contrato de `chamar_gemini` (TASK-055). A Anthropic não tem um modo JSON dedicado — os
    prompts salvos já instruem o modelo a responder só com JSON no próprio texto (mesmo mecanismo
    que já sustenta isso para o Gemini hoje; `response_mime_type` é reforço, não a única garantia)."""
    from anthropic import AsyncAnthropic
    client = AsyncAnthropic(api_key=api_key)
    ultimo = None
    for modelo in modelos:
        try:
            resposta = await asyncio.wait_for(
                client.messages.create(
                    model=modelo, max_tokens=4096, temperature=temperatura,
                    messages=[{"role": "user", "content": texto}],
                ),
                timeout=TIMEOUT_S,
            )
            return "".join(b.text for b in resposta.content if getattr(b, "type", None) == "text")
        except Exception as e:
            if any(k in str(e).lower() for k in _ERROS_DE_FALLBACK):
                ultimo = e
                continue
            raise
    raise ultimo or RuntimeError("Nenhum modelo Claude disponível.")
