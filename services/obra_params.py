"""TASK-165 — parâmetros de obra (``OBRA.*``): valores vindos do cabeçalho da obra, do contexto e da tabela regional.

Parâmetro de ativo cujo nome começa com ``OBRA.`` é preenchido pelo programa, nunca pelo pedido. Lista fechada:
OBRA.MUNICIPIO, OBRA.BAIRRO, OBRA.UTD, OBRA.UTEP (cabeçalho), OBRA.REGIAO (derivado) e OBRA.CONTEXTO (contexto selecionado).
Este arquivo tem a MESMA regra de static/ativo_parametros.js (RULES Regra 4); a tabela regional vai ao navegador por /api/obra/regionais.
"""
import json
import re
import unicodedata

import config

PREFIXO = "OBRA."
NOMES = ("OBRA.MUNICIPIO", "OBRA.BAIRRO", "OBRA.UTD", "OBRA.UTEP", "OBRA.REGIAO", "OBRA.CONTEXTO")
_CAMPOS = {"municipio": "OBRA.MUNICIPIO", "bairro": "OBRA.BAIRRO", "utd": "OBRA.UTD", "utep": "OBRA.UTEP", "contexto": "OBRA.CONTEXTO"}
_TABELAS = {"PERNAMBUCO": config.REGIONAIS_PE_PATH}
_cache = {}


def eh_obra(nome) -> bool:
    return str(nome or "").upper().startswith(PREFIXO)


def normalizar(valor) -> str:
    """Maiúsculas, sem acento, espaços simples e espaço → ``_`` ("Centro Oeste" → CENTRO_OESTE, "34,5kV" → 34,5KV)."""
    t = unicodedata.normalize("NFD", str(valor or ""))
    t = "".join(c for c in t if unicodedata.category(c) != "Mn").upper()
    return re.sub(r"\s+", "_", t.strip())


def _chave(valor) -> str:
    """Chave da tabela: mesma normalização, mas com espaço (as chaves do JSON têm espaço simples)."""
    return normalizar(valor).replace("_", " ")


def tabela_regional(projeto) -> dict:
    arq = _TABELAS.get(str(projeto or "").strip().upper())
    if not arq:
        return {}
    if arq not in _cache:
        try:
            with open(arq, encoding="utf-8") as f:
                _cache[arq] = json.load(f)
        except (OSError, ValueError):
            _cache[arq] = {}
    return _cache[arq]


def regiao_de(obra, projeto="") -> str:
    """OBRA.REGIAO: utep do cabeçalho (se for uma região da tabela) → utd → município(+bairro) → "" (ausente)."""
    obra = obra or {}
    t = tabela_regional(projeto)
    regioes = set(t.get("regioes") or [])
    if not regioes:
        return ""
    utep = normalizar(obra.get("utep"))
    if utep in regioes:
        return utep
    utd = _chave(obra.get("utd"))
    if utd and utd in (t.get("por_utd") or {}):
        return t["por_utd"][utd]
    mun, bairro = _chave(obra.get("municipio")), _chave(obra.get("bairro"))
    if mun:
        if bairro:
            achou = (t.get("por_municipio_bairro") or {}).get(f"{mun}|{bairro}")
            if achou:
                return achou["regiao"]
        return (t.get("por_municipio") or {}).get(mun, "")
    return ""


def montar(obra, projeto="") -> dict:
    """{"OBRA.X": valor normalizado} só dos dados que a obra tem. ``obra`` = {municipio, bairro, utd, utep, contexto} (qualquer chave pode faltar)."""
    obra = obra if isinstance(obra, dict) else {}
    saida = {}
    for campo, nome in _CAMPOS.items():
        v = normalizar(obra.get(campo))
        if v:
            saida[nome] = v
    r = regiao_de(obra, projeto)
    if r:
        saida["OBRA.REGIAO"] = r
    return saida
