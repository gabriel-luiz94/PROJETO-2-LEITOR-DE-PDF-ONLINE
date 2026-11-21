"""
services/exportar_sap_pdf.py — Relatórios ODI e ODD em PDF no formato da "NOTA ... ORÇAMENTO" do SAP (TASK-140).

TASK-161: ODI e ODD saem num PDF só (cada um começa em página nova, com o cabeçalho); as páginas de continuação não repetem o cabeçalho nem a linha verde
(como o relatório do ICP); 48 linhas na 1ª página; quantidades sempre com 2 casas.
Folha A4 retrato: bloco de cabeçalho (TÍTULO, Nº NOTA, BAIRRO, UTD, PROJETISTA | Nº SAP, MUNICIPIO, UTEP, DATA | ID INVESTIMENTO, TIPO DE ORDEM),
tabela TIPO · CÓDIGO · QTD · CENTRO · DESCRIÇÃO DO MATERIAL com o título em verde (#00A443, texto branco) e grade preta fina, Calibri 11 e o rodapé "Internal Use"
em verde. Medidas (em pontos) lidas do PDF de referência. Fonte: Calibri do Windows se existir; senão a Carlito empacotada (mesmas métricas, licença OFL).
"""
import os
import re
from datetime import date

import config
from services import perfil_exportacao as pe

LIMITE_LINHAS = 20000
LIMITE_TEXTO = 200
CAMPOS = ("titulo", "nota", "sap", "bairro", "municipio", "utd", "utep", "projetista", "data", "id_investimento", "id_investimento_odd")
CENTRO = {"M": "1600", "N": "1000"}
_MDO = {"M": "M", "MATERIAL": "M", "N": "N", "S": "N", "MAO-DE-OBRA": "N", "MÃO-DE-OBRA": "N", "": "N", "-": "N"}
_ILEGAIS = re.compile(r"[\x00-\x08\x0b-\x1f]")

PAGINA = (595.32, 841.92)
VERDE_TAB, VERDE_RODAPE, CINZA_TXT = (0.0, 0.6431, 0.2627), (0.0, 0.5, 0.0), (0.251, 0.251, 0.251)
TAM = 11.0
X_TAB = 16.5
COLUNAS = (16.5, 86.58, 183.3, 234.06, 298.5, 573.18)           # TIPO | CÓDIGO | QTD | CENTRO | DESCRIÇÃO
ALT_LINHA = 14.52
ALT_CAB_TAB = 14.64
LINHA_ESP = 0.12
TOPO_BLOCO = 31.44                                              # linha de base do 1º texto do cabeçalho
TOPO_TAB = 107.4
BASE_DENTRO = 10.92                                             # linha de base dentro da linha de dados
BASE_CAB = 11.04
RODAPE = "Internal Use"
LIMITE_Y = 819.5                                                # a tabela não passa daqui (o rodapé fica em 837): 48 linhas na 1ª página, como no relatório do ICP
TOPO_CONT = 14.7                                                # nas páginas de continuação a grade recomeça no topo da folha, SEM cabeçalho nem linha verde (TASK-161)


class ErroPdf(ValueError):
    pass


def _texto(v, limite=LIMITE_TEXTO) -> str:
    return _ILEGAIS.sub("", "" if v is None else str(v)).strip()[:limite]


def _fontes():
    """(caminho da regular, caminho da negrito): Calibri do Windows se houver, senão a Carlito empacotada."""
    cand = []
    win = os.environ.get("WINDIR") or os.environ.get("SystemRoot")
    if win:
        cand.append((os.path.join(win, "Fonts", "calibri.ttf"), os.path.join(win, "Fonts", "calibrib.ttf")))
    cand.append((config._get_resource_path(os.path.join("data", "fontes", "Carlito-Regular.ttf")), config._get_resource_path(os.path.join("data", "fontes", "Carlito-Bold.ttf"))))
    for a, b in cand:
        if os.path.exists(a) and os.path.exists(b):
            return a, b
    raise ErroPdf("Fontes do relatório não encontradas (data/fontes).")


def _qtd(v) -> float:
    if isinstance(v, bool):
        return 0.0
    if isinstance(v, (int, float)):
        return float(v)
    s = str(v or "").strip()
    if not s:
        return 0.0
    try:
        return float(s.replace(".", "").replace(",", ".")) if "," in s else float(s)
    except ValueError:
        return 0.0


def formatar_qtd(v) -> str:
    """3 → '3,00'; 1234.5 → '1.234,50' (Excel pt-BR, 2 casas)."""
    t = f"{round(_qtd(v), 2):,.2f}"
    return t.replace(",", "X").replace(".", ",").replace("X", ".")


def _linhas(tipo, linhas, perfil=None):
    if not isinstance(linhas, list):
        raise ErroPdf(f"{tipo}: precisa ser uma lista.")
    if len(linhas) > LIMITE_LINHAS:
        raise ErroPdf(f"{tipo}: linhas demais (máx. {LIMITE_LINHAS}).")
    saida = []
    for i, x in enumerate(linhas, 1):
        if not isinstance(x, dict):
            raise ErroPdf(f"{tipo} linha {i}: precisa ser um objeto.")
        mdo = _MDO.get(_texto(x.get("mdo")).upper(), _texto(x.get("mdo")).upper()[:1] or "N")
        saida.append((mdo, _texto(x.get("codigo"), 40), formatar_qtd(x.get("qtd")), pe.centro_de(mdo, perfil), _texto(x.get("desc"), 300)))
    return saida


def _cabecalho(c, tipo, perfil=None):
    """Valores do cabeçalho: os campos do perfil (padrão: título, nota, SAP, bairro, município, UTD, UTEP, projetista, data) e o `id` do investimento (ODD, se houver, senão ODI)."""
    c = c if isinstance(c, dict) else {}
    g = lambda k: _texto(c.get(k), 120)
    idi = g("id_investimento_odd") if tipo == "ODD" and g("id_investimento_odd") else g("id_investimento")
    saida = {x["chave"]: g(x["chave"]) for x in pe.normalizar(perfil)["cabecalho"] if x["coluna"] != "id"}
    saida["data"] = g("data") or date.today().strftime("%d/%m/%Y")
    saida["id"] = idi
    return saida


def gerar_pdf(cabecalho: dict, tipo: str, linhas: list, perfil=None) -> bytes:
    """PDF (bytes) do relatório `tipo` ('ODI'|'ODD'). `linhas`: [{mdo, codigo, desc, qtd}]. O cabeçalho vem do diálogo."""
    return gerar_relatorio(cabecalho, [(tipo, linhas)], perfil)


def gerar_pdf_unico(cabecalho: dict, odi: list, odd: list, perfil=None) -> bytes:
    """TASK-161: um PDF só com o ODI e depois o ODD (a seção sem itens é omitida; é preciso ter itens em pelo menos uma)."""
    secoes = [(t, l) for t, l in (("ODI", odi), ("ODD", odd)) if isinstance(l, list) and l]
    if not secoes:
        raise ErroPdf("Não há itens no ODI nem no ODD para gerar o relatório.")
    return gerar_relatorio(cabecalho, secoes, perfil)


def _cor(hexcor, padrao):
    try:
        return tuple(round(int(hexcor[i:i + 2], 16) / 255, 4) for i in (1, 3, 5))
    except Exception:  # noqa: BLE001
        return padrao


def gerar_relatorio(cabecalho: dict, secoes: list, perfil=None) -> bytes:
    """PDF com uma ou mais seções [(tipo, linhas)], cada uma começando em página nova com o cabeçalho completo (TIPO DE ORDEM - ODI|ODD).
    `perfil` = perfil de exportação do projeto (TASK-176): campos do cabeçalho, centro, cor do título e rodapé; sem perfil, o de Neoenergia (saída idêntica à de sempre)."""
    perfil = pe.normalizar(perfil)
    cor_tab = VERDE_TAB if perfil["pdf"]["cor_titulo"].upper() == pe.PADRAO["pdf"]["cor_titulo"].upper() else _cor(perfil["pdf"]["cor_titulo"], VERDE_TAB)
    preparadas = []
    for tipo, linhas in secoes:
        tipo = str(tipo).upper()
        if tipo not in ("ODI", "ODD"):
            raise ErroPdf("Tipo precisa ser ODI ou ODD.")
        preparadas.append((tipo, _linhas(tipo, linhas, perfil), _cabecalho(cabecalho, tipo, perfil)))
    if not preparadas:
        raise ErroPdf("Informe ao menos uma seção (ODI ou ODD).")
    import pymupdf as fitz

    reg, neg = _fontes()
    doc = fitz.open()
    pg = [None]
    f_reg, f_neg = fitz.Font(fontfile=reg), fitz.Font(fontfile=neg)

    def nova_pagina():
        pg[0] = doc.new_page(width=PAGINA[0], height=PAGINA[1])
        p = pg[0]
        p.insert_font(fontname="cal", fontfile=reg)
        p.insert_font(fontname="calb", fontfile=neg)
        rod = perfil["pdf"]["rodape"]
        w = f_reg.text_length(rod, fontsize=12)
        p.insert_text((PAGINA[0] / 2 - w / 2, 837.72), rod, fontname="cal", fontsize=12, color=VERDE_RODAPE)
        return p

    def texto(p, x, y, t, negrito=False, cor=(0, 0, 0), tam=TAM):
        p.insert_text((x, y), t, fontname="calb" if negrito else "cal", fontsize=tam, color=cor)

    def centrado(p, x0, x1, y, t, negrito=False, cor=(0, 0, 0)):
        f = f_neg if negrito else f_reg
        w = f.text_length(t, fontsize=TAM)
        texto(p, (x0 + x1) / 2 - w / 2, y, t, negrito, cor)

    def ajustado(p, x, x1, y, t):
        """Texto à esquerda; se não couber na coluna, reduz a fonte (mínimo 7) em vez de cortar."""
        tam = TAM
        while tam > 7 and f_reg.text_length(t, fontsize=tam) > (x1 - x - 2):
            tam -= 0.5
        texto(p, x, y, t, tam=tam)

    def cabecalho_tabela(p, topo):
        p.draw_rect(fitz.Rect(X_TAB + 0.4, topo, COLUNAS[-1], topo + ALT_CAB_TAB), color=None, fill=cor_tab)
        b = topo + BASE_CAB
        branco = (1, 1, 1)
        centrado(p, COLUNAS[0], COLUNAS[1], b, "TIPO", True, branco)
        centrado(p, COLUNAS[1], COLUNAS[2], b, "CÓDIGO", True, branco)
        texto(p, COLUNAS[2] + 1.8, b, "QTD", True, branco)
        texto(p, COLUNAS[3] + 1.8, b, "CENTRO", True, branco)
        texto(p, COLUNAS[4] + 1.8, b, "DESCRIÇÃO DO MATERIAL", True, branco)

    def grade(p, y0, n, topo_cab=None):
        """Grade das `n` linhas que começam em `y0`; com `topo_cab` inclui também a linha do título (1ª página da seção)."""
        base = y0 + n * ALT_LINHA
        ys = ([topo_cab] if topo_cab is not None else []) + [y0 + i * ALT_LINHA for i in range(n + 1)]
        for y in ys:
            p.draw_rect(fitz.Rect(X_TAB, y - LINHA_ESP / 2, COLUNAS[-1] + LINHA_ESP, y + LINHA_ESP / 2), color=None, fill=(0, 0, 0))
        for x in COLUNAS:
            p.draw_rect(fitz.Rect(x - LINHA_ESP / 2, topo_cab if topo_cab is not None else y0, x + LINHA_ESP / 2, base), color=None, fill=(0, 0, 0))

    for tipo, dados, cab in preparadas:
        p = nova_pagina()
        esq, dir_ = 18.4, 88.4
        l = [TOPO_BLOCO + i * ALT_LINHA for i in range(5)]
        esquerda = [x for x in perfil["cabecalho"] if x["coluna"] == "esq"][:5]                              # 1ª coluna: a partir da 1ª linha
        meio = [x for x in perfil["cabecalho"] if x["coluna"] == "meio"][:4]                                 # 2ª coluna: a partir da 2ª linha
        for k, x in enumerate(esquerda):
            texto(p, esq, l[k], x["rotulo"], True)
            texto(p, dir_, l[k], cab.get(x["chave"], ""))
        for k, x in enumerate(meio):
            texto(p, 235.9, l[k + 1], x["rotulo"], True)
            texto(p, 300.4, l[k + 1], cab.get(x["chave"], ""))
        texto(p, 443.5, l[1], f"{perfil['pdf']['rotulo_id']} - {cab['id']}", True, CINZA_TXT)
        texto(p, 443.5, l[2], f"{perfil['pdf']['rotulo_ordem']}  -  {tipo}", True, CINZA_TXT)

        i, primeira = 0, True
        while True:
            if primeira:
                cabecalho_tabela(p, TOPO_TAB)
                y0 = TOPO_TAB + ALT_CAB_TAB
            else:
                y0 = TOPO_CONT                                              # continuação: só a grade, a partir do topo da folha
            cabem = max(1, int((LIMITE_Y - y0) // ALT_LINHA))
            parte = dados[i:i + cabem]
            for k, (mdo, cod, qtd, centro, desc) in enumerate(parte):
                b = y0 + k * ALT_LINHA + BASE_DENTRO
                centrado(p, COLUNAS[0], COLUNAS[1], b, mdo)
                centrado(p, COLUNAS[1], COLUNAS[2], b, cod)
                centrado(p, COLUNAS[2], COLUNAS[3], b, qtd)
                centrado(p, COLUNAS[3], COLUNAS[4], b, centro)
                ajustado(p, COLUNAS[4] + 1.8, COLUNAS[5], b - 0.12, desc)
            grade(p, y0, len(parte), TOPO_TAB if primeira else None)
            i += len(parte)
            if i >= len(dados):
                break
            p = nova_pagina()
            primeira = False
    out = doc.tobytes(garbage=3, deflate=True)
    doc.close()
    return out
