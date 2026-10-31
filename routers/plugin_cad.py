"""
routers/plugin_cad.py — API para o plugin do nanoCAD (plugin_nanocad/, TASK-104).

O plugin lê o desenho dentro do CAD com as regras do projeto. Com a aplicação aberta (localhost ou online), ele passa a usar as
regras GRAVADAS AQUI e a conferir a obra contra a Base de Orçamento, sem repetir a lógica do cálculo:

    GET  /api/plugin/versao?projeto=081    impressão digital do que muda o resultado (regras do leitor, vínculo, conversão, base).
                                           O plugin consulta a cada poucos segundos; mudou → baixa de novo e reconfere.
    GET  /api/plugin/regras?projeto=081    regras do leitor (processamento + classificação) e de vínculo do projeto, numa chamada.
    POST /api/plugin/conferir              {"projeto", "dados": {"cabos", "outros"}} (as linhas do .obra.json) → para cada linha, os
                                           ativos que a Totalizadora gera e se cada um existe na Base de Orçamento do projeto
                                           (mesmo caminho do modo autônomo: montar_totalizadora → processar_calculo, RN-06).

Só leitura: nada é gravado. Exige login (o middleware protege /api/*). Não registrar o conteúdo das obras em log (RULES Regra 11).
"""
import hashlib
import json
import math

from fastapi import APIRouter, HTTPException, Request
from pydantic import BaseModel
from typing import Any, Dict, Optional

import database
from services import obras_arquivo

router = APIRouter(prefix="/api/plugin", tags=["plugin-cad"])

LIMITE_LINHAS = obras_arquivo.LIMITE_LINHAS


def _projeto(codigo: str) -> dict:
    conn = database.get_row_connection()
    try:
        r = conn.execute("SELECT nome, codigo FROM projetos WHERE codigo = ?", ((codigo or "").strip(),)).fetchone()
    finally:
        conn.close()
    if not r:
        raise HTTPException(status_code=404, detail=f"Projeto '{codigo}' não cadastrado.")
    return {"codigo": r["codigo"], "nome": r["nome"]}


def _user_id(request: Request):
    user = getattr(request.state, "user", None)
    return user["user_id"] if user else None


def _fontes(codigo: str, user_id) -> dict:
    """Tudo que muda o resultado do plugin, das MESMAS fontes que a tela e o modo autônomo usam."""
    from routers.regras import get_regras_conversao
    from routers.regras_leitor import _get_regras
    from routers.vinculacao_config import obter_config_com_origem
    from services.sync_service import get_merged_orcamento
    from services.autonomo.totalizadora import REGRA_PADRAO
    cfg, origem = obter_config_com_origem(codigo)
    return {
        "processamento": _get_regras("processamento", codigo),
        "classificacao": _get_regras("classificacao", codigo),
        "vinculacao": cfg,
        "vinculacao_origem": origem,
        "conversao": (get_regras_conversao(codigo) or {}).get("regras") or [dict(REGRA_PADRAO)],
        "base": get_merged_orcamento(user_id),
    }


def _hash(valor) -> str:
    return hashlib.sha256(json.dumps(valor, ensure_ascii=False, sort_keys=True, default=str).encode("utf-8")).hexdigest()[:16]


def _versoes(f: dict) -> dict:
    partes = {"regras": _hash([f["processamento"], f["classificacao"], f["vinculacao"]]),
              "conversao": _hash(f["conversao"]), "base": _hash(f["base"])}
    partes["geral"] = _hash(partes)
    return partes


@router.get("/versao")
def versao(request: Request, projeto: str):
    """Leve o bastante para o plugin consultar a cada poucos segundos: só devolve as impressões digitais."""
    p = _projeto(projeto)
    return {"projeto": p, "versao": _versoes(_fontes(p["codigo"], _user_id(request)))}


@router.get("/regras")
def regras(request: Request, projeto: str):
    p = _projeto(projeto)
    f = _fontes(p["codigo"], _user_id(request))
    return {"projeto": p, "versao": _versoes(f), "processamento": f["processamento"], "classificacao": f["classificacao"],
            "vinculacao": f["vinculacao"], "vinculacao_origem": f["vinculacao_origem"]}


class ConferirPayload(BaseModel):
    projeto: str
    dados: Dict[str, Any] = {}
    incluir_totalizadora: Optional[bool] = False


def _qtd(q):
    if isinstance(q, float) and (math.isnan(q) or math.isinf(q)):
        return None
    return q


@router.post("/conferir")
def conferir(payload: ConferirPayload, request: Request):
    """Para cada linha de Cabos/Outros: os ativos que ela gera na Totalizadora e quais NÃO existem na Base de Orçamento do projeto."""
    from services.autonomo.totalizadora import montar_totalizadora, normalizar_para_json, payload_calculo
    from services.autonomo.vinculacao import itens_vinculo_para_totalizadora
    from services.orcamento_calc import processar_calculo
    from services.vinculacao_regras import normalizar_config

    p = _projeto(payload.projeto)
    dados = obras_arquivo.normalizar_dados(payload.dados if isinstance(payload.dados, dict) else {})
    cabos, outros = dados["cabos"], dados["outros"]
    if len(cabos) > LIMITE_LINHAS or len(outros) > LIMITE_LINHAS:
        raise HTTPException(status_code=413, detail=f"Linhas demais (máx. {LIMITE_LINHAS} por tabela).")
    f = _fontes(p["codigo"], _user_id(request))
    tot = montar_totalizadora(cabos, outros, f["conversao"], f["base"], p["nome"],
                              itens_vinculo=itens_vinculo_para_totalizadora(cabos, outros), cfg_vinculo=normalizar_config(f["vinculacao"]))
    calc = payload_calculo(tot)
    nao_encontrados = list(processar_calculo(calc["cabos"], calc["outros"], p["nome"], f["base"])["nao_encontrados"])

    def nao_achado(ativo):
        a = str(ativo or "")
        return any(n == a or n.startswith(a + " (") for n in nao_encontrados)

    # TOT-k → linha de origem: o contador da Totalizadora anda uma vez por linha não vazia, Cabos e depois Outros
    origem, k = {}, 1
    for tabela, linhas in (("cabos", cabos), ("outros", outros)):
        for i, linha in enumerate(linhas):
            if linha:
                origem[f"TOT-{k}"] = (tabela, i)
                k += 1
    por_linha = {}
    for item in tot:
        chave = origem.get(item.get("id") or item.get("baseId"))
        if chave is None:
            continue
        # naoEncontrado = o que a Totalizadora da tela pinta de vermelho (sem linha do projeto nem genérica na base);
        # foraDoCalculo = nem a cascata do cálculo (RN-06/RN-08, que ainda tenta outros projetos) achou — some do orçamento
        fora = nao_achado(item.get("ativo"))
        por_linha.setdefault(chave, []).append({"ativo": item.get("ativo"), "operacao": item.get("operacao"), "qtd": _qtd(item.get("qtd")),
                                                "origem": item.get("origem"), "naoEncontrado": bool(item.get("naoEncontrado")) or fora,
                                                "foraDoCalculo": fora, "desc": item.get("desc") or ""})
    linhas = [{"tabela": t, "indice": i, "itens": itens} for (t, i), itens in sorted(por_linha.items(), key=lambda x: (x[0][0] != "cabos", x[0][1]))]
    resposta = {"projeto": p, "versao": _versoes(f), "linhas": linhas, "nao_encontrados": sorted(set(nao_encontrados)),
                "resumo": {"linhas": len(linhas), "com_pendencia": sum(1 for l in linhas if any(x["naoEncontrado"] for x in l["itens"]))}}
    if payload.incluir_totalizadora:
        resposta["totalizadora"] = normalizar_para_json(tot)
    return resposta
