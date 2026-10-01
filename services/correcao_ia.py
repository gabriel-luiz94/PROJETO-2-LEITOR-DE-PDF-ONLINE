"""
services/correcao_ia.py — Correção assistida das planilhas por IA (TASK-015, ADR-004 camada 3).

A IA só PROPÕE: nada é aplicado aqui nem pelo endpoint. Cada proposta é reconferida antes de ir para a tela:
- a linha precisa ser uma das enviadas (a IA não pode inventar linha);
- o "antes" mostrado é sempre o texto real da linha, nunca o que a IA disser;
- proposta que não muda nada, ou que repete a linha, é descartada;
- a proposta passa de novo pela camada 1 (formato de ativo): se ainda reprova, é descartada com o motivo.
A correção altera só o texto do ativo; a operação nunca é tocada.

Não registrar o conteúdo das planilhas em log (RULES Regra 11).
"""
import json
import re

from services.ajustes_planilhas import ajustar, descrever_acao, validar_acoes
from services.validacao_ia import _RE_CERCA, executar_em_lotes
from services.validacao_planilhas import validar_planilhas


def _com_id(linhas, prefixo):
    return [{**linha, "id": linha.get("id") or f"{prefixo}-{i}"} for i, linha in enumerate(linhas or [])]


def linhas_citadas(cabos, outros, achados):
    """Só as linhas de Cabos/Outros citadas nos achados (com id da tela), na ordem da tabela.

    Achados `GERAL` e da Totalizadora não apontam uma linha editável e ficam de fora.
    """
    ids = {a.get("linha_id") for a in achados or []}
    return ([l for l in _com_id(cabos, "CABOS") if l["id"] in ids and (l.get("ativo") or "").strip()],
            [l for l in _com_id(outros, "OUTROS") if l["id"] in ids and (l.get("ativo") or "").strip()])


def _normaliza(texto):
    return re.sub(r"\s+", " ", (texto or "").strip()).upper()


def _reprovado_na_camada_1(tabela, depois):
    """Erros de formato da camada 1 para o texto proposto (a operação é fixada em I para não interferir)."""
    linha = [{"id": "X", "ativo": depois, "operacao": "I"}]
    achados = validar_planilhas(linha, []) if tabela == "cabos" else validar_planilhas([], linha)
    return [a for a in achados if a["severidade"] == "erro"], [a for a in achados if a["severidade"] != "erro"]


def interpretar_correcoes(texto, reais):
    """(correcoes, descartadas). `reais`: {linha_id: (tabela, operacao, ativo_atual)} das linhas enviadas.

    Levanta ValueError se a resposta inteira não for um JSON com a lista "correcoes".
    """
    limpo = _RE_CERCA.sub("", (texto or "").strip())
    try:
        dados = json.loads(limpo)
    except json.JSONDecodeError as e:
        raise ValueError(f"resposta da IA não é JSON válido ({e.msg})")
    itens = dados.get("correcoes") if isinstance(dados, dict) else None
    if not isinstance(itens, list):
        raise ValueError('a resposta da IA não tem a lista "correcoes"')

    correcoes, descartadas, vistos = [], [], set()
    for item in itens:
        if not isinstance(item, dict):
            descartadas.append({"linha_id": None, "motivo": "item não é um objeto"})
            continue
        linha_id = item.get("linha_id")
        if not isinstance(linha_id, str) or linha_id not in reais:
            descartadas.append({"linha_id": linha_id if isinstance(linha_id, str) else None,
                                "motivo": "linha não enviada à IA"})
            continue
        tabela, _operacao, atual = reais[linha_id]
        depois = item.get("depois")
        if not isinstance(depois, str) or not depois.strip():
            descartadas.append({"linha_id": linha_id, "motivo": "sem o texto corrigido"})
            continue
        depois = depois.strip()
        if _normaliza(depois) == _normaliza(atual):
            descartadas.append({"linha_id": linha_id, "motivo": "a correção não muda a linha"})
            continue
        if linha_id in vistos:
            descartadas.append({"linha_id": linha_id, "motivo": "mais de uma proposta para a mesma linha"})
            continue
        erros, avisos = _reprovado_na_camada_1(tabela, depois)
        if erros:
            descartadas.append({"linha_id": linha_id,
                                "motivo": "o texto proposto ainda reprova na validação: " + "; ".join(e["mensagem"] for e in erros)})
            continue
        vistos.add(linha_id)
        motivo = item.get("motivo")
        correcoes.append({
            "linha_id": linha_id, "tabela": tabela,
            "antes": atual,  # sempre o texto real, não o "antes" que a IA devolveu
            "depois": depois,
            "motivo": motivo.strip() if isinstance(motivo, str) else "",
            "avisos": [a["mensagem"] for a in avisos],
        })
    return correcoes, descartadas


async def corrigir_com_ia(chamar, texto_prompt, cabos, outros, achados):
    """Pede correções só para as linhas citadas em `achados`. Nunca levanta e nunca aplica nada.

    Resposta: {"status", "correcoes": [...], "descartadas": [...], "mensagem", "lotes", "truncado"}.
    """
    cabos_c, outros_c = linhas_citadas(cabos, outros, achados)

    def tratar(resposta, lote):
        reais = {i: ("cabos", op, ativo) for i, op, ativo in lote["cabos"]}
        reais.update({i: ("outros", op, ativo) for i, op, ativo in lote["outros"]})
        return interpretar_correcoes(resposta, reais)

    r = await executar_em_lotes(chamar, texto_prompt, cabos_c, outros_c, achados, tratar)
    return {"status": r["status"], "correcoes": r["itens"], "descartadas": r["descartados"],
            "mensagem": r["mensagem"], "lotes": r["lotes"], "truncado": r["truncado"]}


# ═══════════════════════════════════════════════════════════════════════════
# AJUSTES ESTRUTURADOS (TASK-025, ADR-006): a IA propõe AÇÕES; o motor determinístico gera o diff
# ═══════════════════════════════════════════════════════════════════════════
LIMITE_ACOES_IA = 10
ACOES_DESTRUTIVAS = {"excluir_linhas", "remover_ativo"}


def interpretar_acoes(texto):
    """(propostas, descartadas) da resposta da IA. `propostas` = [{"acao": dict, "motivo": str}], ainda sem validar o schema
    (isso é feito uma vez, com os grupos do projeto, em `ajustar_com_ia`). Levanta ValueError se a resposta não for um
    JSON com a lista "acoes"."""
    limpo = _RE_CERCA.sub("", (texto or "").strip())
    try:
        dados = json.loads(limpo)
    except json.JSONDecodeError as e:
        raise ValueError(f"resposta da IA não é JSON válido ({e.msg})")
    itens = dados.get("acoes") if isinstance(dados, dict) else None
    if not isinstance(itens, list):
        raise ValueError('a resposta da IA não tem a lista "acoes"')
    propostas, descartadas = [], []
    for item in itens:
        if not isinstance(item, dict):
            descartadas.append({"motivo": "item não é um objeto"})
            continue
        acao = {k: v for k, v in item.items() if k != "motivo"}
        motivo = item.get("motivo")
        propostas.append({"acao": acao, "motivo": motivo.strip() if isinstance(motivo, str) else ""})
    return propostas, descartadas


async def ajustar_com_ia(chamar, texto_prompt, cabos, outros, achados, grupos=None):
    """Pede AÇÕES de ajuste à IA (só para as linhas citadas em `achados`) e devolve o DIFF calculado pelo motor.

    Nada é aplicado. Cada ação é validada pelo schema do motor; a IA nunca muda a operação da linha (`operacao_nova` é
    recusado); a camada 1 continua sendo a guarda de cada linha. Nunca levanta: falha vira `status`.
    Resposta: {"status", "acoes": [{"indice","acao","frase","motivo","destrutiva"}], "diff": {...}|None, "destrutivo": bool,
               "descartadas": [...], "mensagem", "lotes", "truncado"}.
    """
    cabos_c, outros_c = linhas_citadas(cabos, outros, achados)

    def tratar(resposta, lote):
        propostas, descartadas = interpretar_acoes(resposta)
        return propostas, descartadas

    r = await executar_em_lotes(chamar, texto_prompt, cabos_c, outros_c, achados, tratar)
    validas, descartadas, vistos = [], list(r["descartados"]), set()
    for p in r["itens"]:
        chave = json.dumps(p["acao"], sort_keys=True, ensure_ascii=False)
        if chave in vistos:
            continue  # o mesmo ajuste proposto por mais de um lote
        vistos.add(chave)
        erros = validar_acoes([p["acao"]], grupos)
        if "operacao_nova" in p["acao"]:
            erros.append("a IA não pode mudar a operação da linha")
        if erros:
            descartadas.append({"acao": p["acao"].get("acao"), "motivo": "; ".join(erros)})
        elif len(validas) >= LIMITE_ACOES_IA:
            descartadas.append({"acao": p["acao"].get("acao"), "motivo": f"mais de {LIMITE_ACOES_IA} ações propostas"})
        else:
            validas.append(p)
    resposta = {"status": r["status"], "acoes": [], "diff": None, "destrutivo": False, "descartadas": descartadas,
                "mensagem": r["mensagem"], "lotes": r["lotes"], "truncado": r["truncado"]}
    if not validas:
        return resposta
    acoes = [p["acao"] for p in validas]
    try:
        diff = ajustar(acoes, cabos, outros, grupos)
    except ValueError as e:
        resposta["status"], resposta["mensagem"] = "erro", "; ".join(e.args[0])[:200]
        return resposta
    resposta["acoes"] = [{"indice": i, "acao": a, "frase": descrever_acao(a), "motivo": validas[i]["motivo"],
                          "destrutiva": a["acao"] in ACOES_DESTRUTIVAS} for i, a in enumerate(acoes)]
    resposta["diff"] = diff
    resposta["destrutivo"] = any(x["destrutiva"] for x in resposta["acoes"]) or diff["resumo"]["excluir"] > 0
    return resposta
