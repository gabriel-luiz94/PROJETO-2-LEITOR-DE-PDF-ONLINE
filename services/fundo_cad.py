"""TASK-233: fundo do CAD para o mapa de vinculação, a partir da seção "cad" da obra do plugin do nanoCAD.

O plugin guarda em "cad" todos os blocos de poste (`postes`, inclusive os sem anotação) e todos os trechos da rede (`vaos`, inclusive
os sem anotação). O mapa mostra ao fundo, esmaecido, só o que NÃO virou linha da obra: os trechos de rede sem cabo anotado e os blocos de
poste livres (sem anotação) — é ali que falta montar a rede. Mesmo formato de `services.dxf_geometria.geometria_de_fundo`
({"redes": [{"t", "p"}], "postes": [[x, y]]}), para o mapa tratar as duas origens igual.
"""
from __future__ import annotations

import math

MAX_PONTOS = 40000
MAX_POSTES = 20000


def _num(v):
    try:
        n = float(v)
    except (TypeError, ValueError):
        return None
    return n if math.isfinite(n) else None


def _ponto(p):
    if not isinstance(p, (list, tuple)) or len(p) < 2:
        return None
    x, y = _num(p[0]), _num(p[1])
    return None if x is None or y is None else [round(x, 1), round(y, 1)]


def _anotado(e: dict) -> bool:
    return bool(str(e.get("ativoLido") or "").strip()) and not e.get("excluido")


def fundo_do_cad(cad) -> dict | None:
    """{"redes": [{"t": "MT"|"BT"|"", "p": [[x, y], ...]}], "postes": [[x, y], ...]} ou None quando não há nada a mostrar."""
    if not isinstance(cad, dict):
        return None
    redes, total, vistos = [], 0, set()
    for v in cad.get("vaos") or []:
        if not isinstance(v, dict) or _anotado(v):
            continue
        pts = [q for q in (_ponto(p) for p in (v.get("pontos") or [])) if q]
        if len(pts) < 2:
            continue
        chave = (str(v.get("linha") or ""), v.get("trecho"))
        if chave[0] and chave in vistos:
            continue
        vistos.add(chave)
        total += len(pts)
        if total > MAX_PONTOS:
            break
        tipo = str(v.get("tipo") or "").upper()
        redes.append({"t": tipo if tipo in ("MT", "BT") else "", "p": pts})
    # blocos livres: com handle e sem anotação em nenhum poste no mesmo lugar (troca de poste deixa dois blocos juntos)
    anotados = set()
    for p in cad.get("postes") or []:
        if isinstance(p, dict) and _anotado(p):
            q = _ponto([p.get("x"), p.get("y")])
            if q and str(p.get("handle") or ""):
                anotados.add((round(q[0]), round(q[1])))
    postes, vistos = [], set()
    for p in cad.get("postes") or []:
        if not isinstance(p, dict) or _anotado(p) or p.get("excluido") or not str(p.get("handle") or ""):
            continue
        q = _ponto([p.get("x"), p.get("y")])
        if not q:
            continue
        chave = (round(q[0]), round(q[1]))
        if chave in anotados or chave in vistos:
            continue
        vistos.add(chave)
        postes.append(q)
        if len(postes) >= MAX_POSTES:
            break
    if not redes and not postes:
        return None
    return {"redes": redes, "postes": postes}
