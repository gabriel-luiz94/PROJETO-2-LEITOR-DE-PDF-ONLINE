"""
services/leitura_desenho.py — Leitura do desenho (DXF) configurável por projeto: layers e cores de rede (TASK-175, ADR-007).

Os MESMOS nomes de campo do plugin do nanoCAD (plugin_nanocad/src/Nucleo/ConfigLeitura.cs): `layers_rede_mt`, `layers_rede_bt`, `cor_confirma`, `blocos_poste`,
`layers_poste`, `layers_guia` — mais `cor_mt`, `cor_bt` e `regra`. A mesma configuração serve aos dois. Sem configuração: o padrão de Pernambuco (REDE_MT/REDE_BT,
vermelho = MT, verde = BT, a cor manda e a layer confirma), idêntico ao que era fixo em services/dxf_geometria.py.

  * trecho de layer: o nome CONTÉM o trecho (sem diferença de maiúsculas); "=NOME" exige o nome exato; com `*`/`?` é curinga e vale para o nome inteiro ("MT_*").
  * `cor_mt`/`cor_bt`: vazio = vermelho/verde de sempre; número = cor ACI (1 = vermelho); "#RRGGBB" = cor exata; {"r": [mín, máx], "g": [...], "b": [...]} = faixa RGB; ou lista deles.
  * `regra`: COR_MANDA (padrão: a layer precisa ser de rede e uma cor reconhecida que contradiz a layer descarta a linha) | LAYER_MANDA (a layer decide e a cor não descarta;
    em layer que não é de rede, uma cor MT/BT reconhecida vale) | SO_LAYER (só a layer; a cor é ignorada). `cor_confirma: false` (plugin) equivale a SO_LAYER.
"""
import fnmatch
import re

REGRAS = ("COR_MANDA", "LAYER_MANDA", "SO_LAYER")
_COR_MT = {"r": [170, 255], "g": [0, 100], "b": [0, 100]}
_COR_BT = {"r": [0, 120], "g": [140, 255], "b": [0, 120]}
PADRAO = {"layers_rede_mt": ["REDE_MT"], "layers_rede_bt": ["REDE_BT"], "cor_mt": None, "cor_bt": None, "regra": "COR_MANDA",
          "blocos_poste": ["poste"], "layers_poste": [], "layers_guia": []}


def _trim(t) -> str:
    return str(t if t is not None else "").strip()


def _lista(v, padrao):
    return [_trim(x) for x in v if _trim(x)] if isinstance(v, list) else list(padrao)


def normalizar(cfg) -> dict:
    """Configuração completa (campos que faltam usam o padrão de Pernambuco)."""
    cfg = cfg if isinstance(cfg, dict) else {}
    c = {"layers_rede_mt": _lista(cfg.get("layers_rede_mt"), PADRAO["layers_rede_mt"]), "layers_rede_bt": _lista(cfg.get("layers_rede_bt"), PADRAO["layers_rede_bt"]),
         "cor_mt": cfg.get("cor_mt") if cfg.get("cor_mt") not in ("", []) else None, "cor_bt": cfg.get("cor_bt") if cfg.get("cor_bt") not in ("", []) else None,
         "blocos_poste": _lista(cfg.get("blocos_poste"), PADRAO["blocos_poste"]), "layers_poste": _lista(cfg.get("layers_poste"), []), "layers_guia": _lista(cfg.get("layers_guia"), [])}
    r = _trim(cfg.get("regra")).upper()
    c["regra"] = r if r in REGRAS else ("SO_LAYER" if cfg.get("cor_confirma") is False else "COR_MANDA")
    return c


def contem(trechos, nome) -> bool:
    n = _trim(nome).upper()
    for t in trechos or []:
        x = _trim(t).upper()
        if not x:
            continue
        if x.startswith("="):
            if n == x[1:].strip():
                return True
        elif "*" in x or "?" in x:
            if fnmatch.fnmatchcase(n, x):
                return True
        elif x in n:
            return True
    return False


def tipo_pela_layer(layer, cfg=None):
    c = normalizar(cfg)
    if contem(c["layers_rede_mt"], layer):
        return "MT"
    if contem(c["layers_rede_bt"], layer):
        return "BT"
    return None


def _rgb(hexcor):
    try:
        return int(hexcor[1:3], 16), int(hexcor[3:5], 16), int(hexcor[5:7], 16)
    except Exception:  # noqa: BLE001
        return None


def _aci_rgb(aci):
    try:
        from ezdxf.colors import aci2rgb
        return tuple(aci2rgb(int(aci)))
    except Exception:  # noqa: BLE001
        return None


def _casa_cor(spec, rgb) -> bool:
    if isinstance(spec, list):
        return any(_casa_cor(s, rgb) for s in spec)
    if isinstance(spec, bool):
        return False
    if isinstance(spec, int):
        return _aci_rgb(spec) == rgb
    if isinstance(spec, str):
        t = _trim(spec)
        if re.fullmatch(r"#[0-9A-Fa-f]{6}", t):
            return _rgb(t) == rgb
        return t.isdigit() and _aci_rgb(int(t)) == rgb
    if isinstance(spec, dict):
        return all(isinstance(spec.get(k), list) and len(spec[k]) == 2 and spec[k][0] <= v <= spec[k][1] for k, v in zip("rgb", rgb))
    return False


def tipo_pela_cor(hexcor, cfg=None):
    c = normalizar(cfg)
    rgb = _rgb(hexcor or "")
    if rgb is None:
        return None
    if _casa_cor(c["cor_mt"] if c["cor_mt"] is not None else _COR_MT, rgb):
        return "MT"
    if _casa_cor(c["cor_bt"] if c["cor_bt"] is not None else _COR_BT, rgb):
        return "BT"
    return None


def classificar_rede(layer, hexcor, cfg=None):
    """MT | BT | None (não é rede ou foi descartada), conforme a regra do projeto."""
    c = normalizar(cfg)
    da_layer = tipo_pela_layer(layer, c)
    if c["regra"] == "SO_LAYER":
        return da_layer
    da_cor = tipo_pela_cor(hexcor, c)
    if c["regra"] == "LAYER_MANDA":
        return da_layer or da_cor
    if da_layer is None:
        return None
    return da_layer if da_cor is None or da_cor == da_layer else None


def eh_bloco_poste(nome_bloco, layer=None, cfg=None) -> bool:
    c = normalizar(cfg)
    return contem(c["blocos_poste"], nome_bloco) or contem(c["layers_poste"], layer)


def pode_ser_guia(layer, cfg=None) -> bool:
    c = normalizar(cfg)
    if tipo_pela_layer(layer, c) is not None:
        return False
    return not c["layers_guia"] or contem(c["layers_guia"], layer)


def validar(cfg) -> list:
    erros = []
    if cfg is None:
        return erros
    if not isinstance(cfg, dict):
        return ["A configuração de leitura tem de ser um objeto."]
    if cfg.get("regra") not in (None, "") and _trim(cfg["regra"]).upper() not in REGRAS:
        erros.append(f"regra: precisa ser um de {list(REGRAS)}")
    for k in ("cor_mt", "cor_bt"):
        v = cfg.get(k)
        for s in (v if isinstance(v, list) else [v]):
            if s in (None, ""):
                continue
            if isinstance(s, bool) or not (isinstance(s, int) and 0 <= s <= 255 or isinstance(s, str) and (re.fullmatch(r"#[0-9A-Fa-f]{6}", _trim(s)) or _trim(s).isdigit())
                                           or isinstance(s, dict) and all(isinstance(s.get(x), list) and len(s[x]) == 2 and all(isinstance(n, int) for n in s[x]) for x in "rgb")):
                erros.append(f"{k}: use um número ACI (1 = vermelho), \"#RRGGBB\" ou {{\"r\": [mín, máx], \"g\": [mín, máx], \"b\": [mín, máx]}}")
                break
    return erros


def testar(cfg, layer, cor) -> dict:
    """"Testar" da tela: layer + cor (#RRGGBB ou número ACI) → MT | BT | descartada | não é rede, com o motivo."""
    c = normalizar(cfg)
    hexcor = cor
    if isinstance(cor, int) or _trim(cor).isdigit():
        rgb = _aci_rgb(int(cor))
        hexcor = "#%02x%02x%02x" % rgb if rgb else ""
    da_layer, da_cor = tipo_pela_layer(layer, c), tipo_pela_cor(hexcor, c)
    tipo = classificar_rede(layer, hexcor, c)
    if tipo:
        resultado, motivo = tipo, f"layer {da_layer or 'sem tipo'}, cor {da_cor or 'neutra'}, regra {c['regra']}"
    elif da_layer is None and not (c["regra"] == "LAYER_MANDA" and da_cor):
        resultado, motivo = "NAO_E_REDE", "a layer não é de rede"
    else:
        resultado, motivo = "DESCARTADA", f"a layer diz {da_layer} e a cor diz {da_cor}"
    return {"resultado": resultado, "tipo": tipo, "da_layer": da_layer, "da_cor": da_cor, "regra": c["regra"], "motivo": motivo}
