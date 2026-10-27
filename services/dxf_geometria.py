"""
services/dxf_geometria.py (TASK-091) — geometria implícita do DXF para o mapa do vínculo (Pernambuco).

Lê, SEM criar linhas na tabela do Leitor (que continua só com textos):
  - redes: LINE/LWPOLYLINE/POLYLINE nas layers REDE_MT e REDE_BT. O tipo vem da COR (vermelho = MT, verde = BT) e a
    LAYER confirma: cor e layer divergentes (ex.: vermelha em REDE_BT) descartam a linha;
  - postes: INSERT cujo nome de bloco contém "poste" (sem distinguir maiúsculas);
  - guias: LEADER e linhas curtas FORA das layers de rede (a "linha-guia" que liga a anotação à rede).
Depois associa cada texto: com guia, segue a guia até a rede/poste; sem guia, usa o mais próximo dentro do raio
(RAIO_ALTURAS x a altura do texto). O texto ganha `_geo` (pontos da linha da rede), `_geoTipo` ('MT'|'BT') e `_poste` ([x, y]).
A associação é por proximidade para qualquer texto; quem decide se o texto é cabo (usa `_geo`) ou estrutura (usa `_poste`) é o Resumo.
"""
import math
from collections import defaultdict
from config import logger
from services.dxf_service import _get_entity_base_color

RAIO_ALTURAS = 4          # raio de busca = 4 x altura do texto (pedido: entre 3 e 5)
GUIA_MAX_ALTURAS = 40     # linha fora da rede só vale como guia se tiver até 40 x altura típica do texto
MAX_PTS_GEO = 60
_PROF_BLOCOS = 3


def _tipo_pela_cor(hexcor: str):
    try:
        r, g, b = int(hexcor[1:3], 16), int(hexcor[3:5], 16), int(hexcor[5:7], 16)
    except Exception:
        return None
    if r >= 170 and g <= 100 and b <= 100:
        return "MT"
    if g >= 140 and r <= 120 and b <= 120:
        return "BT"
    return None


def _tipo_pela_layer(layer: str):
    u = (layer or "").upper()
    if "REDE_MT" in u:
        return "MT"
    if "REDE_BT" in u:
        return "BT"
    return None


def classificar_rede(layer: str, hexcor: str):
    """A cor decide o tipo e a layer confirma. Cor fora de vermelho/verde: vale a layer. Divergência: None (descarta)."""
    da_layer = _tipo_pela_layer(layer)
    if da_layer is None:
        return None
    da_cor = _tipo_pela_cor(hexcor)
    if da_cor is None or da_cor == da_layer:
        return da_layer
    return None


def _pontos(ent):
    tp = ent.dxftype()
    if tp == "LINE":
        s, e = ent.dxf.start, ent.dxf.end
        return [(float(s.x), float(s.y)), (float(e.x), float(e.y))]
    if tp == "LWPOLYLINE":
        pts = [(float(p[0]), float(p[1])) for p in ent.get_points("xy")]
        if ent.closed and pts:
            pts.append(pts[0])
        return pts
    if tp == "POLYLINE":
        pts = [(float(v.dxf.location.x), float(v.dxf.location.y)) for v in ent.vertices]
        if ent.is_closed and pts:
            pts.append(pts[0])
        return pts
    if tp == "LEADER":
        return [(float(v[0]), float(v[1])) for v in ent.vertices]
    return []


def _comprimento(pts):
    return sum(math.dist(pts[i], pts[i + 1]) for i in range(len(pts) - 1))


def coletar_geometria(doc) -> dict:
    """{'redes': [{tipo, layer, pts}], 'postes': [(x, y)], 'guias': [[(x, y), ...]]} — guias ainda sem filtro de tamanho."""
    redes, postes, guias = [], [], []

    def varrer(entidades, prof):
        for ent in entidades:
            try:
                tp = ent.dxftype()
                if tp == "INSERT":
                    nome = str(ent.dxf.name or "")
                    if "poste" in nome.lower():
                        p = ent.dxf.insert
                        postes.append((float(p.x), float(p.y)))
                    elif prof < _PROF_BLOCOS:
                        varrer(ent.virtual_entities(), prof + 1)
                elif tp in ("LINE", "LWPOLYLINE", "POLYLINE", "LEADER"):
                    pts = _pontos(ent)
                    if len(pts) < 2:
                        continue
                    layer = getattr(ent.dxf, "layer", "") or ""
                    tipo = classificar_rede(layer, _get_entity_base_color(ent, doc)) if tp != "LEADER" else None
                    if tipo:
                        redes.append({"tipo": tipo, "layer": layer, "pts": pts})
                    elif _tipo_pela_layer(layer) is None:
                        guias.append(pts)
            except Exception as e:     # uma entidade estranha não pode derrubar a extração de texto
                logger.debug(f"geometria: entidade ignorada ({e})")

    varrer(doc.modelspace(), 0)
    return {"redes": redes, "postes": postes, "guias": guias}


# ── associação ──

def _dist_ponto_segmento(p, a, b):
    ax, ay, bx, by = a[0], a[1], b[0], b[1]
    dx, dy = bx - ax, by - ay
    t = 0.0 if dx == 0 and dy == 0 else max(0.0, min(1.0, ((p[0] - ax) * dx + (p[1] - ay) * dy) / (dx * dx + dy * dy)))
    return math.hypot(p[0] - (ax + t * dx), p[1] - (ay + t * dy))


class _GradeSegmentos:
    """Índice espacial simples: cada segmento entra nas células que sua caixa cobre (segmentos muito longos ficam à parte)."""

    def __init__(self, itens, celula):
        self.celula = max(celula, 1e-6)
        self.grade = defaultdict(list)
        self.grandes = []
        for idx, pts in itens:
            for i in range(len(pts) - 1):
                a, b = pts[i], pts[i + 1]
                c0x, c1x = sorted((int(a[0] // self.celula), int(b[0] // self.celula)))
                c0y, c1y = sorted((int(a[1] // self.celula), int(b[1] // self.celula)))
                if (c1x - c0x + 1) * (c1y - c0y + 1) > 400:
                    self.grandes.append((idx, a, b))
                    continue
                for cx in range(c0x, c1x + 1):
                    for cy in range(c0y, c1y + 1):
                        self.grade[(cx, cy)].append((idx, a, b))

    def mais_proximo(self, p, raio):
        """(idx, distância) do segmento mais próximo dentro do raio, ou None."""
        n = int(raio // self.celula) + 1
        cx, cy = int(p[0] // self.celula), int(p[1] // self.celula)
        melhor = None
        vistos = set()
        cand = list(self.grandes)
        for i in range(cx - n, cx + n + 1):
            for j in range(cy - n, cy + n + 1):
                cand.extend(self.grade.get((i, j), ()))
        for idx, a, b in cand:
            chave = (idx, a, b)
            if chave in vistos:
                continue
            vistos.add(chave)
            d = _dist_ponto_segmento(p, a, b)
            if d <= raio and (melhor is None or d < melhor[1]):
                melhor = (idx, d)
        return melhor


class _GradePontos:
    def __init__(self, pontos, celula):
        self.celula = max(celula, 1e-6)
        self.grade = defaultdict(list)
        for i, p in enumerate(pontos):
            self.grade[(int(p[0] // self.celula), int(p[1] // self.celula))].append((i, p))

    def mais_proximo(self, p, raio):
        n = int(raio // self.celula) + 1
        cx, cy = int(p[0] // self.celula), int(p[1] // self.celula)
        melhor = None
        for i in range(cx - n, cx + n + 1):
            for j in range(cy - n, cy + n + 1):
                for idx, q in self.grade.get((i, j), ()):
                    d = math.dist(p, q)
                    if d <= raio and (melhor is None or d < melhor[1]):
                        melhor = (idx, d)
        return melhor


def _simplificar(pts):
    pts = [[round(x, 2), round(y, 2)] for x, y in pts]
    if len(pts) > MAX_PTS_GEO:
        passo = (len(pts) - 1) / (MAX_PTS_GEO - 1)
        pts = [pts[round(i * passo)] for i in range(MAX_PTS_GEO)]
    return pts


def associar_textos(rows: list, geo: dict) -> None:
    """Acrescenta `_geo`/`_geoTipo`/`_poste` às linhas de texto (in place). Não cria nem remove linhas."""
    redes, postes = geo["redes"], geo["postes"]
    if not redes and not postes:
        return
    alturas = sorted(r["tamanho"] for r in rows if r.get("tamanho"))
    if not alturas:
        return
    h_tipica = alturas[len(alturas) // 2]
    celula = h_tipica * RAIO_ALTURAS * 2
    grade_redes = _GradeSegmentos([(i, r["pts"]) for i, r in enumerate(redes)], celula) if redes else None
    grade_postes = _GradePontos(postes, celula) if postes else None
    guias = [g for g in geo["guias"] if _comprimento(g) <= GUIA_MAX_ALTURAS * h_tipica]
    pontas = [g[0] for g in guias] + [g[-1] for g in guias]
    grade_guias = _GradePontos(pontas, celula) if pontas else None

    def ancoras(p, raio):
        """Pontos de partida da busca: a outra ponta de cada guia que toca o texto (do mais perto ao mais longe), depois o próprio texto."""
        saida = []
        if grade_guias:
            n = int(raio // grade_guias.celula) + 1
            cx, cy = int(p[0] // grade_guias.celula), int(p[1] // grade_guias.celula)
            achados = []
            for i in range(cx - n, cx + n + 1):
                for j in range(cy - n, cy + n + 1):
                    for idx, q in grade_guias.grade.get((i, j), ()):
                        d = math.dist(p, q)
                        if d <= raio:
                            achados.append((d, idx))
            for _, idx in sorted(achados):
                g = guias[idx % len(guias)]
                saida.append(g[-1] if idx < len(guias) else g[0])
        saida.append(p)
        return saida

    for r in rows:
        h = r.get("tamanho") or 0
        if not h or "_x" not in r:
            continue
        raio = RAIO_ALTURAS * h
        p = (r["_x"], r["_y"])
        pontos_busca = ancoras(p, raio)
        if grade_redes:
            for q in pontos_busca:
                achado = grade_redes.mais_proximo(q, raio)
                if achado:
                    rede = redes[achado[0]]
                    r["_geo"] = _simplificar(rede["pts"])
                    r["_geoTipo"] = rede["tipo"]
                    break
        if grade_postes:
            for q in pontos_busca:
                achado = grade_postes.mais_proximo(q, raio)
                if achado:
                    r["_poste"] = [round(postes[achado[0]][0], 2), round(postes[achado[0]][1], 2)]
                    break


def geometria_de_fundo(geo: dict, max_pontos: int = 40000) -> dict:
    """Versão enxuta da rede e dos postes para o mapa desenhar ao fundo (esmaecida)."""
    redes, total = [], 0
    for r in geo["redes"]:
        pts = _simplificar(r["pts"])
        total += len(pts)
        if total > max_pontos:
            break
        redes.append({"t": r["tipo"], "p": pts})
    return {"redes": redes, "postes": [[round(x, 1), round(y, 1)] for x, y in geo["postes"][:20000]]}
