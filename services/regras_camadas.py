"""
services/regras_camadas.py — Regras de domínio em camadas: padrão (DEFAULT) + ajustes do projeto (TASK-018, ADR-005).

Funções puras (sem banco/rede). O projeto guarda só o *overlay*:
    {"versao": 2, "overlay": true,
     "grupos":      {NOME: [...]},          # grupos novos ou redefinidos só no projeto
     "adicionadas": [regra, ...],           # regras exclusivas do projeto
     "sobrescritas": {id: {campo: valor}},  # campos trocados numa regra herdada (valor None = remove o campo)
     "ocultas":     [id, ...]}              # regras herdadas desligadas/recolhidas no projeto (botão −)

O efetivo = padrão com o overlay por cima; onde o projeto sobrescreve, vale o do projeto. Cada regra sai anotada com
`origem` ("padrao" | "sobrescrita" | "projeto") e `oculta` (bool) — metadados que o motor ignora.
"""
import copy

from services.regras_dominio import CAMPOS_META, normalizar_regra

ORIGEM_PADRAO, ORIGEM_SOBRESCRITA, ORIGEM_PROJETO = "padrao", "sobrescrita", "projeto"


def overlay_vazio() -> dict:
    return {"versao": 2, "overlay": True, "grupos": {}, "adicionadas": [], "sobrescritas": {}, "ocultas": []}


def eh_overlay(bruto) -> bool:
    return isinstance(bruto, dict) and bool(bruto.get("overlay"))


def overlay_vazio_de(overlay: dict) -> bool:
    return not (overlay["grupos"] or overlay["adicionadas"] or overlay["sobrescritas"] or overlay["ocultas"])


def normalizar_overlay(bruto) -> dict:
    o = overlay_vazio()
    if isinstance(bruto, dict):
        o["grupos"] = dict(bruto.get("grupos") or {})
        o["adicionadas"] = [_limpa(normalizar_regra(r)) for r in bruto.get("adicionadas") or []]
        o["sobrescritas"] = {k: dict(v) for k, v in (bruto.get("sobrescritas") or {}).items()}
        o["ocultas"] = list(dict.fromkeys(bruto.get("ocultas") or []))
    return o


def _limpa(regra: dict) -> dict:
    return {k: v for k, v in regra.items() if k not in CAMPOS_META}


def efetivo(padrao: dict, overlay: dict):
    """(regras, grupos, grupos_origem, avisos) do padrão + overlay. `padrao` = {"grupos", "regras"} já normalizado."""
    overlay = normalizar_overlay(overlay)
    grupos = {**padrao["grupos"], **overlay["grupos"]}
    grupos_origem = {n: (ORIGEM_SOBRESCRITA if n in padrao["grupos"] else ORIGEM_PROJETO) if n in overlay["grupos"]
                     else ORIGEM_PADRAO for n in grupos}
    ids_padrao = {r["id"] for r in padrao["regras"]}
    regras, avisos = [], []
    for r in padrao["regras"]:
        base = _limpa(copy.deepcopy(r))
        ajuste = overlay["sobrescritas"].get(r["id"])
        origem = ORIGEM_PADRAO
        if ajuste:
            for campo, valor in ajuste.items():
                if valor is None:
                    base.pop(campo, None)
                else:
                    base[campo] = copy.deepcopy(valor)
            origem = ORIGEM_SOBRESCRITA
        regras.append({**normalizar_regra(base), "origem": origem, "oculta": r["id"] in overlay["ocultas"]})
    for r in overlay["adicionadas"]:
        if r["id"] in ids_padrao:
            avisos.append(f"A regra própria '{r['id']}' repete o id de uma regra do padrão e foi ignorada.")
            continue
        regras.append({**normalizar_regra(copy.deepcopy(r)), "origem": ORIGEM_PROJETO, "oculta": False})
    for rid in list(overlay["sobrescritas"]) + list(overlay["ocultas"]):
        if rid not in ids_padrao:
            avisos.append(f"O ajuste da regra '{rid}' foi ignorado: ela não existe mais no padrão.")
    return regras, grupos, grupos_origem, list(dict.fromkeys(avisos))


def overlay_de_efetivo(padrao: dict, regras: list, grupos: dict) -> dict:
    """Diferença entre uma lista editada (efetivo) e o padrão → overlay. Regra do padrão ausente da lista = oculta."""
    o = overlay_vazio()
    por_id = {r["id"]: r for r in padrao["regras"]}
    vistos = set()
    for bruta in regras:
        r = normalizar_regra(bruta)
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
    o["grupos"] = {n: v for n, v in (grupos or {}).items() if padrao["grupos"].get(n) != v}
    return o


def overlay_de_bruto(padrao: dict, bruto) -> dict:
    """Qualquer formato guardado (overlay, container v2 ou lista v1 = cópia integral antiga) → overlay."""
    if eh_overlay(bruto):
        return normalizar_overlay(bruto)
    if isinstance(bruto, list):
        return overlay_de_efetivo(padrao, bruto, padrao["grupos"])
    if isinstance(bruto, dict):
        return overlay_de_efetivo(padrao, bruto.get("regras") or [], bruto.get("grupos") or {})
    return overlay_vazio()
