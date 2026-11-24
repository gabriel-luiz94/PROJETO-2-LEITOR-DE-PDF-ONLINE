"""
services/dxf_geometria.py (TASK-091) — geometria implícita do DXF para o mapa do vínculo (Pernambuco).

Lê, SEM criar linhas na tabela do Leitor (que continua só com textos):
  - redes: LINE/LWPOLYLINE/POLYLINE nas layers REDE_MT e REDE_BT. O tipo vem da COR (vermelho = MT, verde = BT) e a
    LAYER confirma: cor e layer divergentes (ex.: vermelha em REDE_BT) descartam a linha;
  - postes: INSERT cujo nome de bloco contém "poste" (sem distinguir maiúsculas);
  - guias: LEADER e linhas curtas FORA das layers de rede (a "linha-guia" que liga a anotação à rede).
Depois associa cada texto: com guia, segue a guia até a rede/poste; sem guia, usa o mais próximo dentro do raio
(RAIO_ALTURAS x a altura do texto). TASK-239: a anotação emoldurada (retângulo fechado em volta do texto) procura a guia pela BORDA
da moldura; o texto de cabo que diz a tensão (kV, A35P → MT; PB, RAMA, C16( → BT) só vai para a rede do seu tipo; marcas de "X" (dois
traços iguais cruzados pelo meio, riscando uma anotação) não são rede. O texto ganha `_geo` (pontos da linha da rede), `_geoTipo` ('MT'|'BT') e `_poste` ([x, y]).
A associação é por proximidade para qualquer texto; quem decide se o texto é cabo (usa `_geo`) ou estrutura (usa `_poste`) é o Resumo.
"""
import math
import re
from collections import defaultdict
from config import logger
from services import leitura_desenho as ld
from services.dxf_service import _get_entity_base_color

RAIO_ALTURAS = 4          # raio de busca = 4 x altura do texto (pedido: entre 3 e 5)
GUIA_MAX_ALTURAS = 40     # linha fora da rede só vale como guia se tiver até 40 x altura típica do texto
MAX_PTS_GEO = 60
MOLDURA_MAX_ALTURAS = 50  # moldura de anotação: polilinha fechada de 4 a 12 vértices com largura + altura até 50 x a altura típica
_RE_MT = re.compile(r"\d\s*KV\b|\bA\d+P\b|\dXA\d+P", re.I)
_RE_BT = re.compile(r"\bPB\s*\d|\bRAMA|(\b|\dX)C\d+\s*\(", re.I)


def nivel_do_texto(texto: str):
    """TASK-239: MT | BT quando o texto do cabo diz a tensão (kV / cabo protegido A..P → MT; multiplex PB, RAMA, concêntrico C16( → BT), senão None."""
    t = str(texto or "")
    if _RE_MT.search(t):
        return "MT"
    if _RE_BT.search(t):
        return "BT"
    return None
_PROF_BLOCOS = 3


def _tipo_pela_cor(hexcor: str, cfg=None):
    return ld.tipo_pela_cor(hexcor, cfg)


def _tipo_pela_layer(layer: str, cfg=None):
    return ld.tipo_pela_layer(layer, cfg)


def classificar_rede(layer: str, hexcor: str, cfg=None):
    """Tipo da rede (MT|BT) ou None. Sem `cfg`: a cor decide o tipo e a layer confirma (cor fora de vermelho/verde: vale a layer; divergência: None, descarta).
    Com `cfg` (TASK-175, configuração de leitura do projeto): layers, cores e regra (COR_MANDA | LAYER_MANDA | SO_LAYER) vêm dela — ver services/leitura_desenho.py."""
    return ld.classificar_rede(layer, hexcor, cfg)


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


def coletar_geometria(doc, cfg=None) -> dict:
    """`cfg` = configuração de leitura do projeto (TASK-175; sem ela, o padrão de Pernambuco). {'redes': [{tipo, layer, pts}], 'postes': [(x, y)], 'guias': [[(x, y), ...]]} — guias ainda sem filtro de tamanho."""
    redes, postes, guias, molduras = [], [], [], []

    def varrer(entidades, prof):
        for ent in entidades:
            try:
                tp = ent.dxftype()
                if tp == "INSERT":
                    nome = str(ent.dxf.name or "")
                    if ld.eh_bloco_poste(nome, getattr(ent.dxf, "layer", "") or "", cfg):
                        p = ent.dxf.insert
                        postes.append((float(p.x), float(p.y)))
                    elif prof < _PROF_BLOCOS:
                        varrer(ent.virtual_entities(), prof + 1)
                elif tp == "MULTILEADER":
                    for chamada in ent.context.leaders:
                        for linha in chamada.lines:
                            pts = [(float(v.x), float(v.y)) for v in linha.vertices]
                            if pts:
                                ponta = chamada.last_leader_point
                                guias.append(pts + [(float(ponta.x), float(ponta.y))])
                elif tp in ("LINE", "LWPOLYLINE", "POLYLINE", "LEADER"):
                    pts = _pontos(ent)
                    if len(pts) < 2:
                        continue
                    layer = getattr(ent.dxf, "layer", "") or ""
                    tipo = classificar_rede(layer, _get_entity_base_color(ent, doc), cfg) if tp != "LEADER" else None
                    if tipo:
                        redes.append({"tipo": tipo, "layer": layer, "pts": pts})
                    elif _eh_moldura(ent, pts):
                        xs, ys = [q[0] for q in pts], [q[1] for q in pts]
                        molduras.append((min(xs), min(ys), max(xs), max(ys)))          # TASK-239: moldura da anotação (não é guia)
                    elif ld.pode_ser_guia(layer, cfg):
                        guias.append(pts)
            except Exception as e:     # uma entidade estranha não pode derrubar a extração de texto
                logger.debug(f"geometria: entidade ignorada ({e})")

    varrer(doc.modelspace(), 0)
    return {"redes": sem_marcas_x(redes), "postes": postes, "guias": guias, "molduras": molduras}


def _eh_moldura(ent, pts) -> bool:
    fechada = (ent.dxftype() == "LWPOLYLINE" and ent.closed) or (ent.dxftype() == "POLYLINE" and ent.is_closed) or (len(pts) > 3 and math.dist(pts[0], pts[-1]) < 1e-6)
    return fechada and 4 <= len(pts) <= 13


def sem_marcas_x(redes: list) -> list:
    """TASK-239: tira da rede as marcas de "X" — dois traços retos da mesma layer, de comprimento igual (±5%), com o meio no mesmo lugar."""
    dois = [i for i, r in enumerate(redes) if len(r["pts"]) == 2 and math.dist(*r["pts"]) > 0]
    if len(dois) < 2:
        return redes
    celula = max(1e-6, sorted(math.dist(*redes[i]["pts"]) for i in dois)[len(dois) // 2])
    grade = defaultdict(list)
    meio = lambda r: ((r["pts"][0][0] + r["pts"][1][0]) / 2, (r["pts"][0][1] + r["pts"][1][1]) / 2)
    for i in dois:
        m = meio(redes[i])
        grade[(int(m[0] // celula), int(m[1] // celula))].append(i)
    fora = set()
    for i in dois:
        a = redes[i]
        ma, la = meio(a), math.dist(*a["pts"])
        cx, cy = int(ma[0] // celula), int(ma[1] // celula)
        for gx in (cx - 1, cx, cx + 1):
            for gy in (cy - 1, cy, cy + 1):
                for j in grade.get((gx, gy), ()):
                    if j <= i or redes[j]["layer"] != a["layer"]:
                        continue
                    b = redes[j]
                    lb = math.dist(*b["pts"])
                    if abs(la - lb) <= 0.05 * max(la, lb) and math.dist(ma, meio(b)) <= 0.1 * la and _cruzam(a["pts"], b["pts"]):
                        fora.update((i, j))
    return [r for k, r in enumerate(redes) if k not in fora] if fora else redes


def _cruzam(p, q) -> bool:
    (ax, ay), (bx, by) = p
    (cx, cy), (dx, dy) = q
    d = (bx - ax) * (dy - cy) - (by - ay) * (dx - cx)
    if abs(d) < 1e-9:
        return False                                     # paralelos: não é X
    t = ((cx - ax) * (dy - cy) - (cy - ay) * (dx - cx)) / d
    u = ((cx - ax) * (by - ay) - (cy - ay) * (bx - ax)) / d
    return 0 <= t <= 1 and 0 <= u <= 1


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
    grade_tipo = {t: _GradeSegmentos([(i, r["pts"]) for i, r in enumerate(redes) if r["tipo"] == t], celula) for t in ("MT", "BT")} if redes else {}
    grade_postes = _GradePontos(postes, celula) if postes else None
    guias = [g for g in geo["guias"] if _comprimento(g) <= GUIA_MAX_ALTURAS * h_tipica]
    pontas = [g[0] for g in guias] + [g[-1] for g in guias]
    grade_guias = _GradePontos(pontas, celula) if pontas else None
    molduras = [m for m in geo.get("molduras") or [] if (m[2] - m[0]) + (m[3] - m[1]) <= MOLDURA_MAX_ALTURAS * h_tipica]
    grade_molduras = defaultdict(list)
    for k, m in enumerate(molduras):
        for i in range(int(m[0] // celula), int(m[2] // celula) + 1):
            for j in range(int(m[1] // celula), int(m[3] // celula) + 1):
                grade_molduras[(i, j)].append(k)

    def moldura_do_texto(p, h):
        """A moldura que contém o ponto do texto (com folga de 1 altura); a menor, se houver mais de uma."""
        dentro = [molduras[k] for k in grade_molduras.get((int(p[0] // celula), int(p[1] // celula)), ())
                  if molduras[k][0] - h <= p[0] <= molduras[k][2] + h and molduras[k][1] - h <= p[1] <= molduras[k][3] + h]
        return min(dentro, key=lambda m: (m[2] - m[0]) * (m[3] - m[1])) if dentro else None

    def ancoras(p, raio, moldura=None):
        """Pontos de partida da busca: a outra ponta de cada guia que toca o texto — ou a BORDA da moldura do texto (TASK-239) —
        do mais perto ao mais longe, depois o próprio texto."""
        saida = []
        if grade_guias:
            x0, y0, x1, y1 = moldura if moldura else (p[0], p[1], p[0], p[1])
            dist = lambda q: math.hypot(max(x0 - q[0], 0, q[0] - x1), max(y0 - q[1], 0, q[1] - y1))
            achados = []
            for i in range(int((x0 - raio) // grade_guias.celula), int((x1 + raio) // grade_guias.celula) + 1):
                for j in range(int((y0 - raio) // grade_guias.celula), int((y1 + raio) // grade_guias.celula) + 1):
                    for idx, q in grade_guias.grade.get((i, j), ()):
                        d = dist(q)
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
        pontos_busca = ancoras(p, raio, moldura_do_texto(p, h) if molduras else None)
        if grade_redes:
            nivel = nivel_do_texto(r.get("texto"))
            for q in pontos_busca:
                # TASK-239: o cabo que diz a tensão só vai para a rede do seu tipo (MT nunca na linha BT, e vice-versa), como no plugin
                achado = grade_tipo[nivel].mais_proximo(q, raio) if nivel else grade_redes.mais_proximo(q, raio)
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
