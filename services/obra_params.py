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


def limpar_cache():
    _cache.clear()


def tabela_regional(projeto) -> dict:
    """Tabela regional (município/UTD → região) do projeto (nome ou código): a importada no pacote do projeto ou, na falta, a do pacote de partida dele; {} se não há (TASK-177)."""
    chave = str(projeto or "").strip().upper()
    if not chave:
        return {}
    if chave not in _cache:
        tabela = {}
        try:
            from services import pacote_projeto as pp
            proj = pp.projeto_do_cadastro(chave)
            guardada = pp.parte_guardada(proj["codigo"], "regional") if proj else None
            tabela = guardada or pp.parte_inicial(chave, "regional") or {}
        except Exception:  # noqa: BLE001
            tabela = {}
        _cache[chave] = tabela if isinstance(tabela, dict) else {}
    return _cache[chave]


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


def nomes_do_projeto(projeto) -> tuple:
    """TASK-176: os parâmetros OBRA.* que a base deste projeto pode usar: a lista de sempre + as chaves do cabeçalho do perfil de exportação marcadas `obra` (OBRA.CONTRATO…)."""
    extras = ()
    try:
        from routers.vinculacao_config import perfil_do_projeto
        from services import perfil_exportacao as pe
        perfil = perfil_do_projeto(projeto) if str(projeto or "").strip() else None
        extras = tuple(n for n in pe.chaves_obra(perfil).values() if n not in NOMES) if perfil else ()
    except Exception:  # noqa: BLE001
        extras = ()
    return NOMES + extras


def montar(obra, projeto="", perfil=None) -> dict:
    """{"OBRA.X": valor normalizado} só dos dados que a obra tem. ``obra`` = {municipio, bairro, utd, utep, contexto} (qualquer chave pode faltar).
    `perfil` = perfil de exportação do projeto (TASK-176): as chaves do cabeçalho marcadas `obra` viram OBRA.<CHAVE> (padrão: municipio, bairro, utd, utep) + REGIAO e CONTEXTO."""
    from services import perfil_exportacao as pe
    obra = obra if isinstance(obra, dict) else {}
    saida = {}
    campos = dict(pe.chaves_obra(perfil) if perfil else _CAMPOS)
    campos["contexto"] = "OBRA.CONTEXTO"
    for campo, nome in campos.items():
        v = normalizar(obra.get(campo))
        if v:
            saida[nome] = v
    r = regiao_de(obra, projeto)
    if r:
        saida["OBRA.REGIAO"] = r
    return saida
