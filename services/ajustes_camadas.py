"""
services/ajustes_camadas.py — Ajustes (receitas) em camadas: padrão (DEFAULT) + ajustes do projeto (TASK-023, ADR-006).

Mesmo modelo de services/regras_camadas.py, para a lista de receitas (sem grupos próprios: os grupos de ativos são os das
regras). O padrão guarda {"versao": 1, "ajustes": [receita, ...]}; o projeto guarda só o overlay:
    {"versao": 1, "overlay": true, "adicionadas": [receita], "sobrescritas": {id: {campo: valor|None}}, "ocultas": [id]}
Efetivo = padrão com o overlay por cima (vale a sobrescrita do projeto); cada receita sai com `origem`
("padrao" | "sobrescrita" | "projeto") e `oculta`. Receita = {id, nome, descricao?, ativa, acoes: [...], regras?: [ids]}.
Funções puras, sem I/O.
"""
import copy

CAMPOS_META = {"origem", "oculta", "frases"}
ORIGEM_PADRAO, ORIGEM_SOBRESCRITA, ORIGEM_PROJETO = "padrao", "sobrescrita", "projeto"


def overlay_vazio() -> dict:
    return {"versao": 1, "overlay": True, "adicionadas": [], "sobrescritas": {}, "ocultas": []}


def eh_overlay(bruto) -> bool:
    return isinstance(bruto, dict) and bool(bruto.get("overlay"))


def overlay_vazio_de(o: dict) -> bool:
    return not (o["adicionadas"] or o["sobrescritas"] or o["ocultas"])


def _limpa(r: dict) -> dict:
    return {k: v for k, v in r.items() if k not in CAMPOS_META}


def lista_de_bruto(bruto) -> list:
    """Container {"ajustes": [...]} ou lista pura → lista de receitas."""
    if isinstance(bruto, list):
        return [dict(r) for r in bruto if isinstance(r, dict)]
    if isinstance(bruto, dict):
        return [dict(r) for r in bruto.get("ajustes") or [] if isinstance(r, dict)]
    return []


def normalizar_overlay(bruto) -> dict:
    o = overlay_vazio()
    if isinstance(bruto, dict):
        o["adicionadas"] = [_limpa(r) for r in bruto.get("adicionadas") or [] if isinstance(r, dict)]
        o["sobrescritas"] = {k: dict(v) for k, v in (bruto.get("sobrescritas") or {}).items() if isinstance(v, dict)}
        o["ocultas"] = list(dict.fromkeys(bruto.get("ocultas") or []))
    return o


def efetivo(padrao: list, overlay: dict):
    """(receitas, avisos). `padrao` = lista de receitas do DEFAULT."""
    overlay = normalizar_overlay(overlay)
    ids_padrao = {r["id"] for r in padrao}
    saida, avisos = [], []
    for r in padrao:
        base, origem = _limpa(copy.deepcopy(r)), ORIGEM_PADRAO
        ajuste = overlay["sobrescritas"].get(r["id"])
        if ajuste:
            for campo, valor in ajuste.items():
                if valor is None:
                    base.pop(campo, None)
                else:
                    base[campo] = copy.deepcopy(valor)
            origem = ORIGEM_SOBRESCRITA
        saida.append({**base, "origem": origem, "oculta": r["id"] in overlay["ocultas"]})
    for r in overlay["adicionadas"]:
        if r.get("id") in ids_padrao:
            avisos.append(f"O ajuste próprio '{r['id']}' repete o id de um ajuste do padrão e foi ignorado.")
            continue
        saida.append({**copy.deepcopy(r), "origem": ORIGEM_PROJETO, "oculta": False})
    for rid in list(overlay["sobrescritas"]) + list(overlay["ocultas"]):
        if rid not in ids_padrao:
            avisos.append(f"O ajuste da receita '{rid}' foi ignorado: ela não existe mais no padrão.")
    return saida, list(dict.fromkeys(avisos))


def overlay_de_efetivo(padrao: list, editadas: list) -> dict:
    """Diferença entre a lista editada e o padrão → overlay. Receita do padrão ausente da lista = oculta."""
    o = overlay_vazio()
    por_id = {r["id"]: r for r in padrao}
    vistos = set()
    for bruta in editadas:
        r = dict(bruta)
        rid = r["id"]
        vistos.add(rid)
        limpa = _limpa(r)
        if rid not in por_id:
            o["adicionadas"].append(limpa)
            continue
        base = _limpa(por_id[rid])
        ajuste = {c: v for c, v in limpa.items() if base.get(c) != v}
        ajuste.update({c: None for c in base if c not in limpa})
        if ajuste:
            o["sobrescritas"][rid] = ajuste
        if r.get("oculta"):
            o["ocultas"].append(rid)
    o["ocultas"] += [rid for rid in por_id if rid not in vistos]
    return o


def overlay_de_bruto(padrao: list, bruto) -> dict:
    """Overlay guardado ou container/lista completos (cópia integral) → overlay."""
    if eh_overlay(bruto):
        return normalizar_overlay(bruto)
    if isinstance(bruto, (list, dict)):
        return overlay_de_efetivo(padrao, lista_de_bruto(bruto))
    return overlay_vazio()
