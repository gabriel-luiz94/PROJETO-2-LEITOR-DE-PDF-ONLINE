"""
services/pdf_geometria.py (TASK-236) — geometria do PDF vetorial (exportado do CAD) para o mapa de vinculação.

O PDF não tem layers nem blocos, mas tem:
  - a posição de cada texto (pymupdf: caixa do trecho), em pontos da folha (1 pt = 1/72 pol.), y para baixo;
  - as linhas desenhadas (`page.get_drawings()`): a rede é reconhecida pela COR do traço (mesma regra de cor do projeto do DXF:
    vermelho = MT, verde = BT; `services.leitura_desenho.tipo_pela_cor`). Cada caminho do PDF vira uma polilinha (o CAD exporta um vão
    por polilinha); as pontas das polilinhas da rede valem como postes (sem bloco de poste no PDF);
  - linhas curtas de outra cor podem ser linhas-guia (mesma regra do DXF).
Depois associa cada texto à rede/poste exatamente como no DXF (`services.dxf_geometria.associar_textos`).

ESCALA: a folha está em pontos, não em metros. Só existem duas escalas de projeto — 1:500 e 1:1000. A escala é escolhida pelas metragens
anotadas nos cabos («… 38 m»): a mediana de «metros anotados ÷ comprimento desenhado» fica mais perto de uma das duas. Sem cabo com
metragem ligado à rede, vale 1:500 (não confirmada). Tudo do mapa sai em METROS, com o y para cima e as páginas lado a lado:
`_geo`, `_poste` (ponta de rede perto do texto) e `_posteAprox` (posição do próprio texto). `_x/_y` continuam em pontos da página
(as regras do Leitor e o vínculo por coordenada do PDF já usam essa unidade).
"""
import math
import re
import statistics

from config import logger
from services import leitura_desenho as ld
from services.dxf_geometria import associar_textos, _comprimento, _simplificar

PT_EM_M = 0.0254 / 72                        # 1 pt da folha em metros de papel
ESCALAS = (500, 1000)                        # as duas escalas de projeto
FOLGA_PAGINAS_PT = 72                        # páginas lado a lado com 1 pol. de folga
_RE_METROS = re.compile(r"(\d+(?:[.,]\d+)?)\s*m\b", re.IGNORECASE)
TOL_PONTA_PT = 1.0


def _hex(cor):
    if not cor:
        return None
    try:
        return "#%02x%02x%02x" % tuple(max(0, min(255, int(round(c * 255)))) for c in cor[:3])
    except (TypeError, ValueError):
        return None


def _polilinhas(caminho):
    """Pontos encadeados de um caminho do PDF (linhas e curvas — a curva vale pela corda). Quebra onde o traço não continua."""
    saida, atual = [], None
    for it in caminho.get("items") or []:
        op = it[0]
        if op == "l":
            a, b = (it[1].x, it[1].y), (it[2].x, it[2].y)
        elif op == "c":
            a, b = (it[1].x, it[1].y), (it[4].x, it[4].y)
        else:
            if atual and len(atual) >= 2:
                saida.append(atual)
            atual = None
            continue
        if atual and math.dist(atual[-1], a) <= 0.01:
            atual.append(b)
        else:
            if atual and len(atual) >= 2:
                saida.append(atual)
            atual = [a, b]
    if atual and len(atual) >= 2:
        saida.append(atual)
    return [_sem_volta(pts) for pts in saida]


def _area2(pts):
    return abs(sum(pts[i][0] * pts[i + 1][1] - pts[i + 1][0] * pts[i][1] for i in range(len(pts) - 1)))


def _sem_volta(pts):
    """Caminho «fechado» que só volta pelo mesmo traço (A→B→A, comum no PDF de uma linha): fica só a ida."""
    if len(pts) >= 3 and math.dist(pts[-1], pts[0]) <= 0.01 and _area2(pts) <= 1e-6 * max(1.0, _comprimento(pts)) ** 2:
        return pts[:len(pts) // 2 + 1] if len(pts) % 2 == 1 else pts[:-1]
    return pts


def coletar_geometria_pagina(page, cfg=None) -> dict:
    """{'redes': [{tipo, pts}], 'postes': [(x, y)] (pontas da rede), 'guias': [[(x, y)…]]} em pontos da página."""
    redes, guias = [], []
    for caminho in page.get_drawings():
        try:
            if caminho.get("fill") is not None and caminho.get("color") is None:
                continue                                                              # área preenchida (hachura, símbolo cheio)
            tipo = ld.tipo_pela_cor(_hex(caminho.get("color")), cfg)
            for pts in _polilinhas(caminho):
                if _comprimento(pts) <= 0:
                    continue
                if tipo:
                    redes.append({"tipo": tipo, "layer": "", "pts": pts})
                else:
                    guias.append(pts)
        except Exception as e:                                                        # um desenho estranho não derruba a leitura
            logger.debug(f"pdf: desenho ignorado ({e})")
    postes = []
    for r in redes:
        for p in (r["pts"][0], r["pts"][-1]):
            if not any(math.dist(p, q) <= TOL_PONTA_PT for q in postes):
                postes.append(p)
    return {"redes": redes, "postes": postes, "guias": guias}


def escolher_escala(amostras):
    """`amostras` = [metros anotados ÷ comprimento desenhado em pt]. Devolve {denominador, metros_por_pt, amostras, desvio, confirmada}."""
    validas = [a for a in amostras if a > 0 and math.isfinite(a)]
    if not validas:
        return {"denominador": ESCALAS[0], "metros_por_pt": PT_EM_M * ESCALAS[0], "amostras": 0, "desvio": None, "confirmada": False}
    med = statistics.median(validas)
    melhor = min(ESCALAS, key=lambda d: abs(math.log(med / (PT_EM_M * d))))
    k = PT_EM_M * melhor
    desvio = statistics.median(abs(a / k - 1) for a in validas)
    return {"denominador": melhor, "metros_por_pt": k, "amostras": len(validas), "desvio": round(desvio, 3),
            "confirmada": len(validas) >= 3 and abs(math.log(med / k)) <= math.log(1.25)}


def _centro(item):
    """Centro aproximado do texto (pymupdf dá o canto superior esquerdo): meia altura para baixo, um pouco para a direita."""
    h = float(item.get("tamanho") or 0)
    return item["_x"] + h * 0.3 * max(1, len(str(item.get("texto") or ""))) / 2, item["_y"] + h / 2


def enriquecer(doc, itens, cfg=None):
    """Acrescenta `_geo`/`_geoTipo`/`_poste`/`_posteAprox` (metros) aos itens do PDF (in place). Devolve (geometria_de_fundo, escala)."""
    paginas = {}
    for i, page in enumerate(doc):
        paginas[i + 1] = {"geo": coletar_geometria_pagina(page, cfg), "largura": float(page.rect.width)}
    # associação por página, em pontos (mesma regra do DXF)
    associados = []
    for n, pg in paginas.items():
        proxies = []
        for it in itens:
            if it.get("pagina") != n or "_x" not in it:
                continue
            cx, cy = _centro(it)
            proxies.append((it, {"_x": cx, "_y": cy, "tamanho": it.get("tamanho")}))
        associar_textos([p for _, p in proxies], pg["geo"])
        associados.extend(proxies)
    # escala pelas metragens dos cabos ligados à rede
    amostras = []
    for it, p in associados:
        if "_geo" not in p:
            continue
        m = _RE_METROS.findall(str(it.get("texto") or ""))
        if not m:
            continue
        rede = next((r for r in paginas[it["pagina"]]["geo"]["redes"] if _simplificar(r["pts"]) == p["_geo"]), None)
        comp = _comprimento(rede["pts"]) if rede else _comprimento(p["_geo"])
        if comp > 0:
            amostras.append(float(m[-1].replace(",", ".")) / comp)
    escala = escolher_escala(amostras)
    k = escala["metros_por_pt"]
    deslocamento, x0 = {}, 0.0
    for n in sorted(paginas):
        deslocamento[n] = x0
        x0 += paginas[n]["largura"] + FOLGA_PAGINAS_PT
    para_m = lambda n, x, y: [round((x + deslocamento[n]) * k, 2), round(-y * k, 2)]
    for it, p in associados:
        n = it["pagina"]
        if "_geo" in p:
            it["_geo"] = [para_m(n, x, y) for x, y in p["_geo"]]
            it["_geoTipo"] = p["_geoTipo"]
        if "_poste" in p:
            it["_poste"] = para_m(n, *p["_poste"])
        it["_posteAprox"] = para_m(n, p["_x"], p["_y"])
    redes, postes, total = [], [], 0
    for n, pg in paginas.items():
        for r in pg["geo"]["redes"]:
            pts = [para_m(n, x, y) for x, y in r["pts"]]
            total += len(pts)
            if total > 40000:
                break
            redes.append({"t": r["tipo"], "p": pts})
        postes.extend(para_m(n, x, y) for x, y in pg["geo"]["postes"])
    fundo = {"redes": redes, "postes": [[round(x, 1), round(y, 1)] for x, y in postes[:20000]]} if (redes or postes) else None
    return fundo, escala


def extrair_pdf_com_geometria(doc, cfg=None):
    """(itens, geometria_de_fundo, escala). Falha na geometria nunca derruba a leitura dos textos."""
    from services.pdf_service import extract_pdf_content
    itens = extract_pdf_content(doc)
    try:
        fundo, escala = enriquecer(doc, itens, cfg)
    except Exception as e:
        logger.warning(f"Geometria do PDF ignorada: {e}")
        fundo, escala = None, None
    return itens, fundo, escala
