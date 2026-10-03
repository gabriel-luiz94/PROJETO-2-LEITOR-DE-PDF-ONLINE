"""
services/autonomo/montagem.py — Montagem das tabelas Cabos/Outros a partir da extração (TASK-031, fase A).

PORTE FIEL do que a tela faz em static/script.js (passo do Leitor) e static/resumo.js (passo do Resumo): `computeRowLogic`, a
separação em Cabos/Outros/Ramais, `calcularQtdAtivos`, `extrairFase`, `isStandaloneLine` e `recalcAllQtdAtivos`. A paridade é provada
em tests/test_autonomo_montagem.py executando o JS real. Se a tela mudar essas regras, o teste acusa — atualize os dois lados.

Cuidados de fidelidade com o JavaScript: `\\s`/`trim` do JS (inclui BOM, não inclui \\x1c-\\x1f), `\\d` só ASCII, `$` só no fim da
string e `length` em unidades UTF-16.
"""
import re

from services.autonomo.leitor_js import obter_leitor

_WS = "\t\n\v\f\r                  　﻿"
_S = "[" + _WS + "]"
_RE_FIM_M = re.compile(_S + "+M" + _S + r"*\Z", re.I)
_PREFIXOS = [re.compile(p, re.I) for p in (r"^CAA" + _S + r"+([0-9])", r"^CA" + _S + r"+([0-9])", r"^CU" + _S + r"+([0-9])",
                                          r"^CAZ" + _S + r"+([0-9])", r"^P" + _S + r"+([0-9])")]
_NOMES = ("CAA", "CA", "CU", "CAZ", "P")
_RE_NUM = re.compile(r"[0-9.,]+\Z")
_RE_ESPACOS = re.compile(_S + "+")


def _trim(s: str) -> str:
    return s.strip(_WS)


def _tamanho(s: str) -> int:
    return len(s.encode("utf-16-le")) // 2  # String.length do JS


def _tokens(ativo_texto: str) -> list:
    txt = _trim(ativo_texto).upper().replace("/", "")
    for nome, rx in zip(_NOMES, _PREFIXOS):
        txt = rx.sub(lambda m, n=nome: n + m.group(1), txt, count=1)
    cleaned = _trim(_RE_FIM_M.sub("", txt, count=1))
    return _RE_ESPACOS.split(cleaned)


def calcular_qtd_ativos(ativo_texto, fase_da_proxima=None) -> int:
    if not ativo_texto or not _trim(ativo_texto):
        return 0
    tokens = _tokens(ativo_texto)
    prefixo = tokens[0]
    if re.match(r"^M[0-9]", prefixo, re.I) or re.fullmatch(r"M", prefixo, re.I) or re.match(r"^CAZ", prefixo, re.I):
        return 1
    fase = None
    if len(tokens) >= 3:
        fase = "".join(tokens[1:len(tokens) - 1])
    elif len(tokens) == 1:
        if fase_da_proxima:
            fase = fase_da_proxima
        else:
            return 1
    elif len(tokens) == 2:
        segundo = tokens[1]
        if _RE_NUM.match(segundo):
            if fase_da_proxima:
                fase = fase_da_proxima
            else:
                return 1
        else:
            fase = segundo
    if not fase:
        return 1
    tam = _tamanho(fase)
    if re.match(r"^CAA2", prefixo, re.I) and tam == 1:
        return tam + 1
    if re.match(r"^CA4", prefixo, re.I):
        return tam + 1
    return tam or 1


def extrair_fase(ativo_texto):
    if not ativo_texto or not _trim(ativo_texto):
        return None
    tokens = _tokens(ativo_texto)
    if len(tokens) >= 3:
        return "".join(tokens[1:len(tokens) - 1])
    return None


def linha_standalone(ativo_texto) -> bool:
    if not ativo_texto or not _trim(ativo_texto):
        return False
    tokens = _tokens(ativo_texto)
    if len(tokens) == 1:
        return True
    return len(tokens) == 2 and bool(_RE_NUM.match(tokens[1]))


def recalcular_qtd_ativos(cabos: list) -> None:
    """recalcAllQtdAtivos: de baixo para cima (a linha 'standalone' herda a fase da seguinte). Altera `cabos` no lugar."""
    for i in range(len(cabos) - 1, -1, -1):
        row = cabos[i]
        if not row:
            continue
        fase_prox = None
        if linha_standalone(row.get("ativo")) and i + 1 < len(cabos):
            fase_prox = extrair_fase(cabos[i + 1].get("ativo"))
        atual = row.get("qtdAtivos")
        negativo = atual is not None and str(atual).strip().startswith("-")
        novo = calcular_qtd_ativos(row.get("ativo"), fase_prox)
        row["qtdAtivos"] = "-" + str(novo) if (negativo and novo > 0) else novo


def prefixo_cabo(ativo_texto):
    """Prefixo normalizado do ativo de uma linha Cabos (ex.: "CAA2" de "CAA 2 ABC 30 m") — usado pelo
    vínculo cabo<->estrutura/poste (TASK-036). Mesma normalização de calcular_qtd_ativos/_tokens."""
    if not ativo_texto or not _trim(ativo_texto):
        return None
    tokens = _tokens(ativo_texto)
    if len(tokens) < 2:
        return None
    return tokens[0]


def _item_exportado(item: dict, r: dict) -> dict:
    """O que o botão 'Processar' do Leitor grava (script.js): texto/cor/layer + entidade, operação e ativo.
    `_x`/`_y` (TASK-036): o autônomo controla os dois lados do pipeline (extração e montagem) sem passar
    pelo `static/script.js` do navegador, então pode levar a coordenada adiante sem a limitação que a
    tela tem hoje (o botão "Processar" do Leitor não repassa `_x`/`_y` para o Resumo)."""
    d = {"pagina": item.get("pagina"), "texto": item.get("texto"), "cor": item.get("cor"), "layer": item.get("layer") or "",
         "entidade": r["entidade"], "operacao": r["operacao"], "ativo": r["ativo"]}
    if item.get("_x") is not None and item.get("_y") is not None:
        d["_x"], d["_y"] = item["_x"], item["_y"]
    return d


def montar_tabelas(itens: list, regras_proc: list, regras_cls: list, leitor=None) -> dict:
    """Da extração (lista de {pagina,texto,cor,layer}) às tabelas: {cabos, outros, ramais} como o Resumo as monta."""
    leitor = leitor or obter_leitor()
    # passo 1 — tela do Leitor: o motor roda em todas as linhas
    exportados = [_item_exportado(it, r) for it, r in zip(itens, leitor.processar_lote(itens, regras_proc, regras_cls, "extracao"))]

    # passo 2 — Resumo (computeRowLogic): confia em entidade/ativo já definidos; senão roda o motor de novo
    confia = [bool(e["entidade"]) and e["entidade"] != "0" and bool(e["ativo"]) and _trim(e["ativo"]) != "" for e in exportados]
    refazer = [i for i, ok in enumerate(confia) if not ok]
    novos = dict(zip(refazer, leitor.processar_lote(exportados, regras_proc, regras_cls, "resumo", refazer))) if refazer else {}
    processados = []
    for i, e in enumerate(exportados):
        if confia[i]:
            processados.append({"entidade": e["entidade"], "operacao": e["operacao"] or "M", "ativo": e["ativo"], "_raw": e})
        else:
            r = novos[i]
            processados.append({"entidade": r["entidade"], "operacao": r["operacao"], "ativo": r["ativo"], "_raw": e})

    def linha(r):
        d = {"entidade": r["entidade"] or "0", "operacao": r["operacao"] or "M", "ativo": r["ativo"] or ""}
        raw = r.get("_raw") or {}
        if raw.get("_x") is not None and raw.get("_y") is not None:
            d["_x"], d["_y"] = raw["_x"], raw["_y"]
        return d

    cabos = [linha(r) for r in processados if r["entidade"] == "CABO"]
    outros = [linha(r) for r in processados if r["entidade"] not in ("CABO", "0", "RAMAIS")]
    ramais = []
    for r in processados:
        if r["entidade"] in ("RAMAIS", "IP"):
            ok = True
        elif r["entidade"] == "APOIO":
            txt = (r["_raw"].get("texto") or r["ativo"] or "").upper()
            ok = bool(re.search(r"REC.*CAL[CÇ]ADA", txt, re.I) or re.search(r"CONC.*BASE", txt, re.I))
        else:
            ok = False
        if ok:
            ramais.append({"entidade": r["entidade"], "texto": r["_raw"].get("texto") or r["ativo"] or "",
                           "pagina": r["_raw"]["pagina"] if r["_raw"].get("pagina") is not None else "-"})
    recalcular_qtd_ativos(cabos)
    return {"cabos": cabos, "outros": outros, "ramais": ramais}
