"""
services/importar_sap_pdf.py — lê o PDF SAP (ODI / ODD) e devolve as linhas (TASK-157).

Reconhece o PDF gerado pelo programa (`services/exportar_sap_pdf.py`) e o feito por fora no mesmo formato: cabeçalho com "TIPO DE ORDEM - ODI|ODD" e a
tabela TIPO | CÓDIGO | QTD | CENTRO | DESCRIÇÃO DO MATERIAL. Precisa de texto selecionável (PDF escaneado como imagem não é lido).
ODI → operação I (instalação); ODD → operação R (retirada). Funções puras: sem banco, sem rede. Não registrar o conteúdo em log.
"""
import re

import pymupdf

LIMITE_PAGINAS = 300
LIMITE_LINHAS = 20000
_HIFENS = "-‐‑‒–—−"                                      # o PDF traz hífen "bonito" (U+2010) no cabeçalho
_RE_TIPO_ORDEM = re.compile(r"TIPO\s+DE\s+ORDEM\s*[" + _HIFENS + r"]\s*(ODI|ODD)", re.I)
_RE_NOTA = re.compile(r"^\d{6,}$")
_RE_QTD = re.compile(r"^-?\d{1,3}(?:\.\d{3})*(?:,\d+)?$|^-?\d+(?:,\d+)?$")
_RE_COD = re.compile(r"^[A-Za-z0-9]{1,20}$")                 # material é numérico (2203008); mão de obra tem código com letras (SDEVU1434II)
_RE_CENTRO = re.compile(r"^\d{3,5}$")
_CABECALHO = {"TIPO", "CÓDIGO", "CODIGO", "QTD", "CENTRO", "DESCRIÇÃO DO MATERIAL", "DESCRICAO DO MATERIAL", "INTERNAL USE", "#INTERNAL USE"}
_MDO = {"M": "MATERIAL", "N": "MAO-DE-OBRA", "S": "MAO-DE-OBRA"}


class ErroSapPdf(ValueError):
    """PDF recusado; `args[0]` = mensagem em português."""


def _numero(texto: str) -> float:
    return float(texto.replace(".", "").replace(",", "."))


def _comeca_linha(t: list, j: int) -> bool:
    return (j + 3 < len(t) and re.fullmatch(r"[MNS]", t[j]) is not None and _RE_COD.match(t[j + 1]) is not None
            and _RE_QTD.match(t[j + 2]) is not None and _RE_CENTRO.match(t[j + 3]) is not None)


def _linhas_da_pagina(textos: list) -> list:
    saida, i = [], 0
    while i < len(textos):
        if not _comeca_linha(textos, i):
            i += 1
            continue
        j, desc = i + 4, []
        while j < len(textos) and not _comeca_linha(textos, j) and textos[j].upper() not in _CABECALHO:
            desc.append(textos[j])
            j += 1
        saida.append({"mdo": _MDO[textos[i]], "codigo": (textos[i + 1].lstrip("0") or "0") if textos[i + 1].isdigit() else textos[i + 1].upper(), "qtd": round(_numero(textos[i + 2]), 4),
                      "centro": textos[i + 3], "desc": re.sub(r"\s+", " ", " ".join(desc)).strip()})
        i = j
    return saida


def ler_pdf(conteudo: bytes) -> dict:
    """→ {nota, secoes: [{tipo: 'ODI'|'ODD', linhas: [{mdo, codigo, qtd, centro, desc}]}]}. Levanta ErroSapPdf.
    Um PDF pode trazer o ODI e o ODD juntos (como o relatório do ICP: ODI nas primeiras páginas e o ODD depois). Uma página que traz
    "TIPO DE ORDEM - ODI|ODD" abre uma seção nova quando o tipo muda; páginas sem cabeçalho (continuação da tabela) seguem na seção atual."""
    try:
        doc = pymupdf.open(stream=conteudo, filetype="pdf")
    except Exception:  # noqa: BLE001 — qualquer defeito do arquivo vira mensagem
        raise ErroSapPdf("O arquivo não é um PDF válido.")
    try:
        if doc.page_count > LIMITE_PAGINAS:
            raise ErroSapPdf(f"PDF com páginas demais (máx. {LIMITE_PAGINAS}).")
        nota, secoes, total = "", [], 0
        for pagina in doc:
            textos = [t.strip() for t in pagina.get_text("text").splitlines() if t.strip()]
            m = _RE_TIPO_ORDEM.search(" ".join(textos))
            if m and (not secoes or secoes[-1]["tipo"] != m.group(1).upper()):
                secoes.append({"tipo": m.group(1).upper(), "linhas": []})
            if not nota:
                for k, t in enumerate(textos):
                    if re.match(r"^N[º°o]\s*NOTA", t, re.I) and k + 1 < len(textos) and _RE_NOTA.match(textos[k + 1]):
                        nota = textos[k + 1]
                        break
            linhas = _linhas_da_pagina(textos)
            if linhas and secoes:
                secoes[-1]["linhas"] += linhas
            total += len(linhas)
            if total > LIMITE_LINHAS:
                raise ErroSapPdf(f"PDF com linhas demais (máx. {LIMITE_LINHAS}).")
    finally:
        doc.close()
    if not secoes:
        raise ErroSapPdf('Não encontrei "TIPO DE ORDEM - ODI/ODD" no PDF: não parece um PDF SAP.')
    secoes = [x for x in secoes if x["linhas"]]
    if not secoes:
        raise ErroSapPdf("Não encontrei linhas de material no PDF (se ele for uma imagem escaneada, não dá para ler).")
    return {"nota": nota, "secoes": secoes}


def para_resultado(lidos: list) -> list:
    """Vários PDFs lidos (cada um com uma ou duas seções) → linhas no formato do Resultado do orçamento: {operacao, mdo, codigo, desc_codigo, filtro, total}.
    ODI → I, ODD → R. Levanta ErroSapPdf se o mesmo tipo vier duas vezes (no mesmo PDF ou em PDFs diferentes)."""
    tipos = [sec["tipo"] for x in lidos for sec in x["secoes"]]
    if len(set(tipos)) != len(tipos):
        repetido = next(t for t in tipos if tipos.count(t) > 1)
        raise ErroSapPdf(f"Foram escolhidos dois PDFs {repetido}. Escolha um ODI e/ou um ODD.")
    saida = []
    for x in lidos:
        for sec in x["secoes"]:
            op = "I" if sec["tipo"] == "ODI" else "R"
            saida += [{"operacao": op, "mdo": l["mdo"], "codigo": l["codigo"], "desc_codigo": l["desc"], "filtro": "", "total": l["qtd"]} for l in sec["linhas"]]
    return saida
