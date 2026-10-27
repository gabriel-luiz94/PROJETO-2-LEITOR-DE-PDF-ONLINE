"""
services/vinculacao_regras.py — Regras de vinculação cabo <-> estrutura (TASK-084): tabela de CABOS e tabela de ESTRUTURAS por projeto.

Lógica pura (sem I/O) — a MESMA de static/vinculacao_regras.js (paridade testada em tests/test_vinculacao_regras.py). Usada pela tela (modal de
vinculação), pela Totalizadora (parâmetro `cabo` das estruturas), pela validação e pelo modo autônomo.

Configuração (um JSON por projeto):
  {"versao": 2,
   "cabos":      [{"padrao", "grupo", "valor_cabo", "operacoes", "pontas": {"I": 2, "M": 1}, "estruturas", "vincula", "obs", "ativa"}],
   "estruturas": [{"padrao", "qtd_cabos", "lados": "IGUAIS"|"LIVRE", "amarracao", "aceita_i", "cabos", "parametro", "se_diferem": "BLOQUEAR"|"MAIOR"|"AVISAR", "obs", "ativa"}],
   "ordem_maior": ["C", "P", "A"], "scripts": [{"nome", "script"}]}
`padrao` = regex (JS/Python simples) OU lista separada por vírgula ("N1, N2") OU lista JSON; vazio = qualquer. A primeira linha ativa que casa vale.
Estrutura = código do primeiro item que não é poste da linha de Outros (ou o próprio poste se a linha só tem poste); cabo = código do cabo principal
(sem 3x, sem tensão). Cabo fictício = código guardado na linha de Outros (`cabosFicticios`): conta como um cabo da estrutura, sem metros.
"""
import re

ORDEM_MAIOR_PADRAO = ["C", "P", "A"]
SE_DIFEREM = ("BLOQUEAR", "MAIOR", "AVISAR")
_RE_REGEX = re.compile(r"[\\^$()\[\]|+*?{}]")


def _trim(t) -> str:
    return str(t if t is not None else "").strip()


def casa(padrao, texto) -> bool:
    """Vazio casa tudo; lista (JSON ou "A, B") compara por igualdade (sem caixa); senão regex sem caixa."""
    if padrao is None or padrao == "" or padrao == []:
        return True
    t = _trim(texto).upper()
    if isinstance(padrao, list):
        return any(t == _trim(x).upper() for x in padrao)
    p = _trim(padrao)
    if not p:
        return True
    if not _RE_REGEX.search(p):
        return t in [_trim(x).upper() for x in p.split(",") if _trim(x)]
    try:
        return re.search(p, _trim(texto), re.I | re.A) is not None
    except re.error:
        return False


def config_vazia() -> dict:
    return {"versao": 2, "cabos": [], "estruturas": [], "ordem_maior": list(ORDEM_MAIOR_PADRAO), "scripts": []}


def normalizar_config(cfg) -> dict:
    c = config_vazia()
    if isinstance(cfg, dict):
        c["cabos"] = [r for r in (cfg.get("cabos") or []) if isinstance(r, dict)]
        c["estruturas"] = [r for r in (cfg.get("estruturas") or []) if isinstance(r, dict)]
        om = cfg.get("ordem_maior")
        if isinstance(om, list) and om:
            c["ordem_maior"] = [_trim(x).upper() for x in om]
        c["scripts"] = [s for s in (cfg.get("scripts") or []) if isinstance(s, dict)]
    return c


def config_de_legado(regras_legadas) -> dict:
    """Migra `regras_vinculacao` (tipo_estrutura = 1º dígito, qtd_cabos, compatibilidade) para a Tabela de Estruturas. Tabela de Cabos fica vazia (aceita tudo)."""
    c = config_vazia()
    for r in regras_legadas or []:
        if not isinstance(r, dict) or not _trim(r.get("tipo_estrutura")):
            continue
        tipo = _trim(r["tipo_estrutura"])
        c["estruturas"].append({
            "padrao": f"^[^0-9]*{tipo}", "qtd_cabos": int(r.get("qtd_cabos") or 1),
            "lados": "IGUAIS" if r.get("compatibilidade") == "MESMO_TIPO_FASE_OPERACAO" else "LIVRE",
            "amarracao": True, "aceita_i": True, "cabos": "", "parametro": "cabo", "se_diferem": "MAIOR",
            "obs": r.get("descricao") or f"Migrada: estrutura tipo {tipo}", "ativa": True})
    return c


def validar_config(cfg) -> list:
    erros = []
    if not isinstance(cfg, dict):
        return ["A configuração precisa ser um objeto."]
    for tabela, chave in (("cabos", "Cabo"), ("estruturas", "Estrutura")):
        lista = cfg.get(tabela, [])
        if not isinstance(lista, list):
            erros.append(f"'{tabela}' precisa ser uma lista.")
            continue
        for i, r in enumerate(lista):
            p = f"{chave} #{i + 1}"
            if not isinstance(r, dict):
                erros.append(f"{p}: precisa ser um objeto.")
                continue
            for campo in (("padrao", "estruturas") if tabela == "cabos" else ("padrao", "cabos")):
                v = r.get(campo)
                if isinstance(v, str) and _RE_REGEX.search(v):
                    try:
                        re.compile(v)
                    except re.error as e:
                        erros.append(f"{p}: '{campo}' regex inválida ({e}).")
                elif v is not None and not isinstance(v, (str, list)):
                    erros.append(f"{p}: '{campo}' precisa ser texto ou lista.")
            if tabela == "cabos":
                pontas = r.get("pontas")
                if pontas is not None and (not isinstance(pontas, dict) or any(not isinstance(v, int) or isinstance(v, bool) or v < 0 for v in pontas.values())):
                    erros.append(f"{p}: 'pontas' precisa ser {{\"I\": n, \"M\": n}} com inteiros >= 0.")
                ops = r.get("operacoes")
                if ops is not None and not isinstance(ops, list):
                    erros.append(f"{p}: 'operacoes' precisa ser uma lista.")
                fs = r.get("fases")
                if fs is not None and fs != "" and any(not re.fullmatch(r"[1-9]", _trim(x)) for x in (fs if isinstance(fs, list) else str(fs).split(","))):
                    erros.append(f"{p}: 'fases' precisa ser números de 1 a 9 (ex.: \"3\" ou \"1,2\").")
                if _trim(r.get("nivel")) and _trim(r.get("nivel")).upper() not in ("MT", "BT"):
                    erros.append(f"{p}: 'nivel' precisa ser MT ou BT.")
                pt = r.get("padrao_texto")
                if isinstance(pt, str) and pt:
                    try:
                        re.compile(pt)
                    except re.error as e:
                        erros.append(f"{p}: 'padrao_texto' regex inválida ({e}).")
            else:
                for k in ("qtd_min", "qtd_max", "max_reais"):
                    v = r.get(k)
                    if v is not None and (not isinstance(v, int) or isinstance(v, bool) or v < 1):
                        erros.append(f"{p}: '{k}' precisa ser um inteiro >= 1.")
                if isinstance(r.get("qtd_min"), int) and isinstance(r.get("qtd_max"), int) and r["qtd_min"] > r["qtd_max"]:
                    erros.append(f"{p}: 'qtd_min' não pode ser maior que 'qtd_max'.")
                q = r.get("qtd_cabos")
                if q is not None and (not isinstance(q, int) or isinstance(q, bool) or q < 1):
                    erros.append(f"{p}: 'qtd_cabos' precisa ser um inteiro >= 1.")
                if r.get("lados") not in (None, "IGUAIS", "LIVRE"):
                    erros.append(f"{p}: 'lados' precisa ser IGUAIS ou LIVRE.")
                if r.get("se_diferem") not in (None,) + SE_DIFEREM:
                    erros.append(f"{p}: 'se_diferem' precisa ser um de {list(SE_DIFEREM)}.")
    om = cfg.get("ordem_maior")
    if om is not None and (not isinstance(om, list) or not all(isinstance(x, str) for x in om)):
        erros.append("'ordem_maior' precisa ser uma lista de textos (ex.: [\"C\", \"P\", \"A\"]).")
    return erros


# ── informações de um cabo e de uma estrutura ──────────────────────────────────────────────────────
def op_base(operacao) -> str:
    return "M" if _trim(operacao).upper() in ("M", "*M") else "I"


def _primeira(tabela, campo_padrao, valor):
    for i, r in enumerate(tabela):
        if r.get("ativa") is False:
            continue
        if casa(r.get(campo_padrao), valor):
            return i, r
    return None, None


# TASK-092: fases (1x/2x/3x), tensão (13,8kV) e nível (MT/BT) vêm do TEXTO ORIGINAL do cabo ("3xA35P(7,94)-13,8kV 45 m"); sem "Nx" = 1 fase; MT = tensão acima de 1 kV.
_RE_FASES = re.compile(r"(?<![0-9A-Za-z])([1-9])\s*[xX]\s*(?=[A-Za-z])")
_RE_TENSAO = re.compile(r"(\d+(?:[.,]\d+)?)\s*kV", re.I)


def derivar_cabo(texto) -> dict:
    t = _trim(texto)
    mf, mt = _RE_FASES.search(t), _RE_TENSAO.search(t)
    tensao = mt.group(1).replace(".", ",") + "kV" if mt else ""
    return {"fases": int(mf.group(1)) if mf else 1, "tensao": tensao, "nivel": "MT" if mt and float(mt.group(1).replace(",", ".")) > 1 else "BT"}


def extra_cabo(row) -> dict:
    """Dados que as regras do cabo podem usar: o que está na linha (editável) ou, na falta, o derivado do texto original."""
    row = row or {}
    d = derivar_cabo(row.get("textoOriginal"))
    try:
        f = int(_trim(row.get("fases"))) if _trim(row.get("fases")) != "" else d["fases"]
    except ValueError:
        f = d["fases"]
    return {"texto": _trim(row.get("textoOriginal")), "fases": f, "tensao": _trim(row.get("tensao")) or d["tensao"], "nivel": (_trim(row.get("nivel")).upper() or d["nivel"])}


def _regex_busca(p, texto) -> bool:
    """`padrao_texto` é SEMPRE regex (busca em qualquer ponto do texto original, sem caixa); inválida não casa."""
    try:
        return re.search(_trim(p), _trim(texto), re.I | re.A) is not None
    except re.error:
        return False


def _regra_cabo_casa(r, cod, extra) -> bool:
    if not casa(r.get("padrao"), cod):
        return False
    if not extra:
        return True
    fases = r.get("fases")
    if fases is not None and fases != "":
        lista = fases if isinstance(fases, list) else str(fases).split(",")
        try:
            if extra["fases"] not in [int(_trim(x)) for x in lista]:
                return False
        except ValueError:
            return False
    if _trim(r.get("nivel")) and _trim(r.get("nivel")).upper() != extra["nivel"]:
        return False
    if _trim(r.get("padrao_texto")) and not (extra["texto"] and _regex_busca(r.get("padrao_texto"), extra["texto"])):
        return False
    return True


def _primeira_cabo(tabela, cod, extra):
    for i, r in enumerate(tabela):
        if r.get("ativa") is False:
            continue
        if _regra_cabo_casa(r, cod, extra):
            return i, r
    return None, None


def info_cabo(codigo, operacao, cfg, extra=None) -> dict:
    """`codigo` = código do cabo principal (já sem 3x/tensão). Sem regra na tabela: aceita tudo, pontas 2 (I) / 1 (M)."""
    cod = _trim(codigo).upper()
    i, r = _primeira_cabo(cfg["cabos"], cod, extra) if cod else (None, None)
    base = op_base(operacao)
    pontas = 1 if base == "M" else 2
    if r is not None and isinstance(r.get("pontas"), dict) and isinstance(r["pontas"].get(base), int):
        pontas = r["pontas"][base]
    return {"codigo": cod, "regra": i, "valor_cabo": (_trim(r.get("valor_cabo")).upper() or None) if r else None, "pontas": pontas,
            "vincula": not (r is not None and (r.get("vincula") is False)), "estruturas": (r.get("estruturas") if r else None),
            "operacoes": (r.get("operacoes") if r else None), "op": _trim(operacao).upper(), "base": base,
            "fases": extra["fases"] if extra else None, "nivel": extra["nivel"] if extra else None}


def info_estrutura(codigo, cfg) -> dict:
    cod = _trim(codigo).upper()
    i, r = _primeira(cfg["estruturas"], "padrao", cod) if cod else (None, None)
    d = {"codigo": cod, "regra": i, "qtd": None, "qtd_min": None, "max_reais": None, "lados": "LIVRE", "amarracao": False, "aceita_i": True, "cabos": None, "parametro": "CABO", "se_diferem": "MAIOR"}
    if r is not None:
        # TASK-093: qtd_cabos fixo (min = max) OU faixa qtd_min..qtd_max; max_reais = quantos podem ser cabos REAIS (o excedente só entra como fictício)
        d["qtd"] = r["qtd_max"] if isinstance(r.get("qtd_max"), int) else (r.get("qtd_cabos") if isinstance(r.get("qtd_cabos"), int) else None)
        d["qtd_min"] = r["qtd_min"] if isinstance(r.get("qtd_min"), int) else d["qtd"]
        d["max_reais"] = r["max_reais"] if isinstance(r.get("max_reais"), int) else None
        d["lados"] = "IGUAIS" if r.get("lados") == "IGUAIS" else "LIVRE"
        d["amarracao"] = bool(r.get("amarracao", False))
        d["aceita_i"] = r.get("aceita_i") is not False
        d["cabos"] = r.get("cabos")
        d["parametro"] = (_trim(r.get("parametro")) or "cabo").upper() if r.get("parametro") != "" else ""
        d["se_diferem"] = r.get("se_diferem") if r.get("se_diferem") in SE_DIFEREM else "MAIOR"
    return d


def op_ficticio(e) -> str:
    """Operação assumida por um cabo fictício: M quando a estrutura só aceita cabo M (amarração), senão I."""
    return "M" if (e["regra"] is not None and e["amarracao"] and not e["aceita_i"]) else "I"


def pode_vincular(c, e):
    """(ok, motivo). O cabo aceita a estrutura E a estrutura aceita o cabo (e a operação)."""
    if not c["vincula"]:
        return False, f"o cabo {c['codigo']} não vincula (acompanha o cabo principal)"
    if isinstance(c["operacoes"], list) and c["operacoes"]:
        aceitas = [_trim(x).upper().replace("*", "") for x in c["operacoes"]]
        opn = c["op"].replace("*", "")
        if opn not in aceitas and not (opn == "R" and "I" in aceitas):
            return False, f"a regra do cabo {c['codigo']} não aceita a operação {c['op']}"
    if not casa(c["estruturas"], e["codigo"]):
        return False, f"o cabo {c['codigo']} não aceita a estrutura {e['codigo']}"
    if not casa(e["cabos"], c["codigo"]):
        return False, f"a estrutura {e['codigo']} não aceita o cabo {c['codigo']}"
    if c["base"] == "M" and e["regra"] is not None and not e["amarracao"]:
        return False, f"cabo M só se liga a estrutura de amarração ({e['codigo']} não é)"
    if c["base"] == "I" and e["regra"] is not None and not e["aceita_i"]:
        return False, f"a estrutura {e['codigo']} só aceita cabo M"
    return True, None


def _chave_lado(info_c, ficticio=False):
    return info_c["codigo"] if ficticio else f"{info_c['codigo']}|{info_c['op']}"


def pode_com_ocupantes(c, e, ocupantes, ficticio=False):
    """Como `pode_vincular`, mais capacidade e "lados iguais" contra quem já está ligado. `ocupantes` = [{"info": info_cabo, "ficticio": bool}]."""
    ok, motivo = pode_vincular(c, e)
    if not ok:
        return ok, motivo
    if e["qtd"] is not None and len(ocupantes) + 1 > e["qtd"]:
        return False, f"a estrutura {e['codigo']} já tem os {e['qtd']} cabo(s) que aceita"
    if not ficticio and e["max_reais"] is not None and len([o for o in ocupantes if not o.get("ficticio")]) + 1 > e["max_reais"]:
        return False, f"a estrutura {e['codigo']} aceita só {e['max_reais']} cabo(s) real(is); o excedente deve ser fictício"
    if e["lados"] == "IGUAIS" and ocupantes:
        ref = ocupantes[0]
        mesmo = ref["info"]["codigo"] == c["codigo"] and (ref.get("ficticio") or ref["info"]["op"] == c["op"])
        if not mesmo:
            return False, f"{e['codigo']} exige cabos iguais: {ref['info']['codigo']} × {c['codigo']}"
    return True, None


# ── análise completa (estados, vínculos inválidos, valor do parâmetro) ─────────────────────────────
def _codigo_cabo(row, prefixo_fn):
    return prefixo_fn(row.get("ativo")) or ""


def _eh_cabo_valido(row):
    return bool(row) and bool(_trim(row.get("ativo")))


def _eh_linha_estrutura(r):
    return bool(r) and r.get("entidade") in ("ESTRUTURA", "POSTE")


# TASK-094: uma linha de Outros pode ter VÁRIAS estruturas (`9400D, N1, N4` = poste + N1 + N4); cada uma é uma "unidade" vinculável.
# Identidade da unidade: o índice da linha para a 1ª estrutura (compatível com tudo que já foi salvo) e linha + posição/100 para as seguintes (12 → 12.01 → 12.02).
def id_unidade(ie, k):
    return round(ie + k / 100, 2) if k else ie


def linha_de(id_):
    return int(id_ + 1e-9)


def slot_de(id_):
    return round((id_ - linha_de(id_)) * 100)


def ficticios_de(row, k):
    """Cabos fictícios da unidade: `cabosFicticios` é da 1ª; as demais ficam em `cabosFicticiosSlots[k]`."""
    if not row:
        return []
    return list((row.get("cabosFicticiosSlots") or {}).get(str(k), [])) if k else list(row.get("cabosFicticios") or [])


def _codigos_da_linha(row, codigo_fn, codigos_fn):
    """Códigos das estruturas da linha (sem poste). Sem `codigos_fn` (ou lista vazia) vale o código único de `codigo_fn` (linha só com poste)."""
    lista = [c for c in (codigos_fn(row.get("ativo")) or []) if c] if codigos_fn else []
    return lista or [codigo_fn(row.get("ativo"), row)]


def estruturas_disponiveis(outros, codigo_fn=None, codigos_fn=None):
    r = []
    for i, row in enumerate(outros):
        if not _eh_linha_estrutura(row):
            continue
        if not codigos_fn:
            r.append(i)
            continue
        r.extend(id_unidade(i, k) for k, _ in enumerate(_codigos_da_linha(row, codigo_fn or (lambda a, row=None: ""), codigos_fn)))
    return r


def _expandir(cabos, outros, codigo_fn, codigos_fn):
    """Lista "virtual" de Outros com uma entrada por unidade (as demais linhas passam como estão) e os cabos com os vínculos traduzidos para ela."""
    virt, id_de, j_de = [], [], {}
    for ie, row in enumerate(outros):
        if not _eh_linha_estrutura(row):
            id_de.append(ie)
            virt.append(row)
            continue
        cods = _codigos_da_linha(row, codigo_fn, codigos_fn)
        for k, cod in enumerate(cods):
            id_ = id_unidade(ie, k)
            j_de[id_] = len(virt)
            id_de.append(id_)
            virt.append({**row, "cabosFicticios": ficticios_de(row, k), "_codigoUnidade": cod, "_linha": ie})
    cabos_v = [({**c, "vinculoEstruturas": [j_de.get(i, -1) for i in c["vinculoEstruturas"]]} if c and isinstance(c.get("vinculoEstruturas"), list) else c) for c in cabos]

    def codigo_v(ativo, row=None):
        return row["_codigoUnidade"] if row is not None and "_codigoUnidade" in row else codigo_fn(ativo, row)
    return virt, id_de, cabos_v, codigo_v


def analisar(cabos, outros, cfg, prefixo_fn, codigo_fn, codigos_fn=None) -> dict:
    """Estado de cada cabo e de cada estrutura (cada estrutura da linha é uma unidade; chave = id da unidade). `prefixo_fn(ativo)` = código do cabo principal;
    `codigo_fn(ativo)` = código da estrutura da linha; `codigos_fn(ativo)` (opcional) = lista com o código de CADA estrutura da linha."""
    virt, id_de, cabos_v, codigo_v = _expandir(cabos, outros, _com_linha(codigo_fn), codigos_fn)
    a = _analisar_base(cabos_v, virt, cfg, prefixo_fn, codigo_v)
    return {"estruturas": {id_de[j]: r for j, r in a["estruturas"].items()}, "cabos": a["cabos"],
            "links": [{**l, "estrutura": id_de[l["estrutura"]]} for l in a["links"]]}


def _com_linha(codigo_fn):
    """Aceita `codigo_fn(ativo)` (só o texto) ou `codigo_fn(ativo, row)`."""
    def f(ativo, row=None):
        try:
            return codigo_fn(ativo, row)
        except TypeError:
            return codigo_fn(ativo)
    return f


def _analisar_base(cabos, outros, cfg, prefixo_fn, codigo_fn) -> dict:
    cfg = normalizar_config(cfg)
    ests = estruturas_disponiveis(outros)
    info_c = {}
    for ic, c in enumerate(cabos):
        if _eh_cabo_valido(c):
            info_c[ic] = info_cabo(_codigo_cabo(c, prefixo_fn), c.get("operacao"), cfg, extra_cabo(c))
    info_e = {ie: info_estrutura(codigo_fn(outros[ie].get("ativo"), outros[ie]), cfg) for ie in ests}
    ocup = {ie: [] for ie in ests}                    # [(idx_cabo)]
    for ic, c in enumerate(cabos):
        if ic not in info_c:
            continue
        for ie in (c.get("vinculoEstruturas") or []):
            if ie in ocup:
                ocup[ie].append(ic)
    links, bad_links = [], set()
    res_e = {}
    for ie in ests:
        e = info_e[ie]
        fict = [_trim(x).upper() for x in (outros[ie].get("cabosFicticios") or []) if _trim(x)]
        tem = len(ocup[ie]) + len(fict)
        erros, avisos = [], []
        ocupantes = []
        for ic in ocup[ie]:
            ok, motivo = pode_com_ocupantes(info_c[ic], e, ocupantes)
            ocupantes.append({"info": info_c[ic], "ficticio": False})
            links.append({"cabo": ic, "estrutura": ie, "ok": ok, "motivo": motivo})
            if not ok:
                bad_links.add((ic, ie))
                erros.append(motivo)
        for cod in fict:
            ic_f = info_cabo(cod, op_ficticio(e), cfg)
            ok, motivo = pode_com_ocupantes(ic_f, e, ocupantes, True)
            ocupantes.append({"info": ic_f, "ficticio": True})
            if not ok:
                erros.append(motivo)
        valores = []
        for oc in ocupantes:
            v = oc["info"]["valor_cabo"]
            if v is not None:
                valores.append(v)
        valor = None
        distintos = sorted(set(valores))
        if len(distintos) == 1:
            valor = distintos[0]
        elif len(distintos) > 1:
            if e["se_diferem"] == "MAIOR":
                ordem = cfg["ordem_maior"]
                rank = lambda v: ordem.index(v) if v in ordem else len(ordem)
                valor = sorted(distintos, key=rank)[0]
                avisos.append(f"{e['codigo']}: cabos diferentes ({', '.join(distintos)}); usado o maior ({valor})")
            else:
                avisos.append(f"{e['codigo']}: cabos diferentes ({', '.join(distintos)}); parâmetro não preenchido")
        if erros:
            estado = "rd"
        elif tem == 0:
            estado = "k"
        elif e["qtd_min"] is not None and tem < e["qtd_min"]:
            estado = "pu"
        else:
            estado = "gr"
        res_e[ie] = {"estado": estado, "tem": tem, "precisa": e["qtd_min"], "maximo": e["qtd"], "valor_cabo": valor, "erros": erros, "avisos": avisos, "codigo": e["codigo"],
                     "regra": e["regra"], "ficticios": len(fict), "parametro": e["parametro"]}
    res_c = {}
    for ic, c in info_c.items():
        mine = [ie for ie in (cabos[ic].get("vinculoEstruturas") or []) if ie in ocup]
        invalido = any((ic, ie) in bad_links for ie in mine)
        precisa = c["pontas"] if c["vincula"] else 0
        if invalido:
            estado = "rd"
        elif not c["vincula"]:
            estado = "gr"
        elif not mine:
            estado = "k"
        elif len(mine) < precisa:
            estado = "pu"
        else:
            estado = "gr"
        res_c[ic] = {"estado": estado, "tem": len(mine), "precisa": precisa, "codigo": c["codigo"], "valor_cabo": c["valor_cabo"], "vincula": c["vincula"]}
    return {"estruturas": res_e, "cabos": res_c, "links": links}


def parametros_por_linha(cabos, outros, cfg, prefixo_fn, codigo_fn, codigos_fn=None) -> dict:
    """{indice_em_outros: {"CABO": "C"}} — o parâmetro que a estrutura recebe do(s) cabo(s) vinculado(s) (só onde há valor)."""
    a = analisar(cabos, outros, cfg, prefixo_fn, codigo_fn, codigos_fn)
    saida = {}
    for ie, r in sorted(a["estruturas"].items()):
        linha = linha_de(ie)
        if r["valor_cabo"] and r["parametro"] and linha not in saida:       # várias estruturas na linha: vale a 1ª que tem valor
            saida[linha] = {r["parametro"]: r["valor_cabo"]}
    return saida


def achados(cabos, outros, cfg, prefixo_fn, codigo_fn, codigos_fn=None) -> list:
    """Avisos de validação (formato de services/validacao_planilhas.py), severidade "aviso"."""
    cfg = normalizar_config(cfg)
    a = analisar(cabos, outros, cfg, prefixo_fn, codigo_fn, codigos_fn)
    existe_m = any(_eh_cabo_valido(c) and op_base(c.get("operacao")) == "M" for c in cabos)
    existe_retca = any(r and re.search(r"\bRETCA\b", r.get("ativo") or "", re.I | re.A) for r in outros)
    excecao_zero = existe_m or existe_retca
    saida = []

    def add(ie, msg, expl):
        saida.append({"linha_id": f"OUTROS-{linha_de(ie)}", "tabela": "outros", "regra_id": "VINCULO-ESTRUTURA", "severidade": "aviso", "mensagem": msg, "explicacao": expl})
    for ie, r in sorted(a["estruturas"].items()):
        linha_o = outros[linha_de(ie)]
        nome = f"{linha_o.get('ativo')} [{r['codigo']}]" if codigos_fn and len([c for c in (codigos_fn(linha_o.get("ativo")) or []) if c]) > 1 else linha_o.get("ativo")
        if r["regra"] is None:
            continue
        if r["tem"] == 0:
            if not excecao_zero:
                add(ie, f"Estrutura \"{nome}\" sem nenhum cabo vinculado.", f"{r['codigo']} exige {r['precisa'] or 1} cabo(s); nenhum encontrado, e não há cabo M nem RETCA no projeto para justificar.")
            continue
        for msg in r["erros"]:
            add(ie, f"Estrutura \"{nome}\": vínculo inválido.", msg)
        if not r["erros"] and r["precisa"] is not None and (r["tem"] < r["precisa"] or (r["maximo"] is not None and r["tem"] > r["maximo"])):
            faixa = f"{r['precisa']} a {r['maximo']}" if r["maximo"] is not None and r["maximo"] != r["precisa"] else f"{r['precisa']}"
            add(ie, f"Estrutura \"{nome}\" com {r['tem']} cabo(s) vinculado(s), esperado {faixa}.", f"{r['codigo']} exige {'exatamente ' if faixa == str(r['precisa']) else ''}{faixa} cabo(s) vinculado(s).")
        for msg in r["avisos"]:
            add(ie, f"Estrutura \"{nome}\": {msg}.", "Os cabos vinculados têm valores diferentes para o parâmetro da estrutura.")
    return saida


# ── script de vinculação automática (TASK-084, entrega 2) ───────────────────────────────────────────
MODOS_SCRIPT = ("proximidade", "sequencia")
CONFLITOS_SCRIPT = ("avisar", "ignorar", "parar")


def validar_script(script) -> list:
    """Erros (lista vazia = válido) de um script JSON de vinculação."""
    if not isinstance(script, dict):
        return ["O script precisa ser um objeto JSON."]
    erros = []
    if script.get("conflitos") not in (None,) + CONFLITOS_SCRIPT:
        erros.append(f"'conflitos' precisa ser um de {list(CONFLITOS_SCRIPT)}.")
    passos = script.get("passos")
    if not isinstance(passos, list) or not passos:
        return erros + ["'passos' precisa ser uma lista com pelo menos um passo."]
    for i, p in enumerate(passos):
        pre = f"Passo #{i + 1}"
        if not isinstance(p, dict):
            erros.append(f"{pre}: precisa ser um objeto.")
            continue
        if p.get("modo") not in MODOS_SCRIPT:
            erros.append(f"{pre}: 'modo' precisa ser um de {list(MODOS_SCRIPT)}.")
        raio = p.get("raio_m")
        if raio is not None and (not isinstance(raio, (int, float)) or isinstance(raio, bool) or raio <= 0):
            erros.append(f"{pre}: 'raio_m' precisa ser um número > 0.")
        pontas = p.get("pontas")
        if pontas is not None and (not isinstance(pontas, int) or isinstance(pontas, bool) or pontas < 1):
            erros.append(f"{pre}: 'pontas' precisa ser um inteiro >= 1.")
        cs = p.get("cabos")
        if cs is not None:
            if not isinstance(cs, dict):
                erros.append(f"{pre}: 'cabos' precisa ser um objeto ({{operacao, ativo}}).")
            else:
                if cs.get("operacao") is not None and not isinstance(cs["operacao"], list):
                    erros.append(f"{pre}: 'cabos.operacao' precisa ser uma lista.")
                if isinstance(cs.get("ativo"), str) and _RE_REGEX.search(cs["ativo"]):
                    try:
                        re.compile(cs["ativo"])
                    except re.error as e:
                        erros.append(f"{pre}: 'cabos.ativo' regex inválida ({e}).")
        es = p.get("estruturas")
        if es is not None and es != "aceitas":
            if not isinstance(es, dict):
                erros.append(f"{pre}: 'estruturas' precisa ser \"aceitas\" ou um objeto ({{ativo, amarracao}}).")
            elif isinstance(es.get("ativo"), str) and _RE_REGEX.search(es["ativo"]):
                try:
                    re.compile(es["ativo"])
                except re.error as e:
                    erros.append(f"{pre}: 'estruturas.ativo' regex inválida ({e}).")
    return erros


def _dist(a, b):
    dx, dy = a["_x"] - b["_x"], a["_y"] - b["_y"]
    return (dx * dx + dy * dy) ** 0.5


def _tem_xy(r):
    return r.get("_x") is not None and r.get("_y") is not None


def executar_script(script, cabos, outros, cfg, prefixo_fn, codigo_fn, aplicar=False, codigos_fn=None) -> dict:
    """Roda o script e devolve o que ele vincularia (`aplicar=False`) ou já grava `vinculoEstruturas`/`vinculoOrigem` nos cabos (`aplicar=True`).
    Nunca sobrescreve vínculo existente, a menos que o passo diga `"substituir": true`."""
    virt, id_de, cabos_v, codigo_v = _expandir(cabos, outros, _com_linha(codigo_fn), codigos_fn)
    r = _executar_script_base(script, cabos_v, virt, cfg, prefixo_fn, codigo_v)
    for v in r["vinculos"]:
        v["estruturas"] = [id_de[j] for j in v["estruturas"]]
    if aplicar:
        for v in r["vinculos"]:
            cabos[v["cabo"]]["vinculoEstruturas"] = list(v["estruturas"])
            cabos[v["cabo"]]["vinculoOrigem"] = "script"
    return r


def _executar_script_base(script, cabos, outros, cfg, prefixo_fn, codigo_fn) -> dict:
    cfg = normalizar_config(cfg)
    modo_conf = script.get("conflitos") if script.get("conflitos") in CONFLITOS_SCRIPT else "avisar"
    ests = estruturas_disponiveis(outros)
    info_e = {ie: info_estrutura(codigo_fn(outros[ie].get("ativo"), outros[ie]), cfg) for ie in ests}
    info_c = {ic: info_cabo(_codigo_cabo(c, prefixo_fn), c.get("operacao"), cfg, extra_cabo(c)) for ic, c in enumerate(cabos) if _eh_cabo_valido(c)}
    ligados = {ic: list(cabos[ic].get("vinculoEstruturas") or []) for ic in info_c}       # estado de trabalho
    fict = {ie: [_trim(x).upper() for x in (outros[ie].get("cabosFicticios") or []) if _trim(x)] for ie in ests}
    saida = {"vinculos": [], "conflitos": [], "sem_candidato": [], "passos": [], "parou": False}

    def ocupantes_de(ie, ignorando=None):
        lista = [{"info": info_c[ic], "ficticio": False} for ic in sorted(info_c) if ic != ignorando and ie in ligados[ic]]
        lista += [{"info": info_cabo(cod, op_ficticio(info_e[ie]), cfg), "ficticio": True} for cod in fict[ie]]
        return lista

    def pode(ic, ie):
        c, e = info_c[ic], info_e[ie]
        oc = ocupantes_de(ie, ignorando=ic)
        ok, motivo = pode_com_ocupantes(c, e, oc)
        if ok and e["qtd"] is None and len(oc) + 1 > c["pontas"]:
            return False, f"a estrutura {e['codigo']} já tem os {c['pontas']} cabo(s) esperados"
        return ok, motivo

    for ip, passo in enumerate(script.get("passos") or []):
        r = {"nome": passo.get("nome") or f"Passo {ip + 1}", "alvo": 0, "completos": 0, "parciais": 0, "sem_candidato": 0, "conflitos": 0, "ja_tinham": 0}
        fc = passo.get("cabos") or {}
        ops = [_trim(x).upper() for x in (fc.get("operacao") or [])]
        alvo = []
        for ic in sorted(info_c):
            c = info_c[ic]
            if not c["vincula"]:
                continue
            if ops and c["op"] not in ops and c["op"].replace("*", "") not in [o.replace("*", "") for o in ops]:
                continue
            if not casa(fc.get("ativo"), c["codigo"]):
                continue
            if ligados[ic] and not passo.get("substituir"):
                r["ja_tinham"] += 1
                continue
            alvo.append(ic)
        r["alvo"] = len(alvo)
        fe = passo.get("estruturas", "aceitas")
        def estrutura_ok(ie):
            if fe == "aceitas" or fe is None:
                return True
            if not casa(fe.get("ativo"), info_e[ie]["codigo"]):
                return False
            if fe.get("amarracao") is not None and bool(fe["amarracao"]) != info_e[ie]["amarracao"]:
                return False
            return True
        cand_global = [ie for ie in ests if estrutura_ok(ie)]
        raio = passo.get("raio_m")
        for pos, ic in enumerate(alvo):
            c = info_c[ic]
            need = passo.get("pontas") if isinstance(passo.get("pontas"), int) else c["pontas"]
            antes = list(ligados[ic])
            if passo.get("substituir"):
                ligados[ic] = []
            escolhidas, recusas = [], []
            if passo.get("modo") == "sequencia":
                candidatas = cand_global[pos:pos + need]
            else:
                if not _tem_xy(cabos[ic]):
                    r["sem_candidato"] += 1
                    saida["sem_candidato"].append({"passo": ip, "cabo": ic, "motivo": "cabo sem coordenada"})
                    ligados[ic] = antes
                    continue
                candidatas = sorted((ie for ie in cand_global if _tem_xy(outros[ie]) and (raio is None or _dist(cabos[ic], outros[ie]) <= raio)),
                                    key=lambda ie: (_dist(cabos[ic], outros[ie]), ie))
            for ie in candidatas:
                if len(escolhidas) >= need:
                    break
                ok, motivo = pode(ic, ie)
                if ok and outros[ie].get("_linha") is not None and any(outros[e2].get("_linha") == outros[ie].get("_linha") for e2 in escolhidas):
                    continue                                   # TASK-094: 2 estruturas do mesmo poste não valem como as 2 pontas
                if ok:
                    escolhidas.append(ie)
                    ligados[ic].append(ie)
                else:
                    recusas.append(motivo)
            completo = need > 0 and len(escolhidas) == need
            if completo or (escolhidas and passo.get("aceitar_parcial")):
                saida["vinculos"].append({"passo": ip, "cabo": ic, "estruturas": list(escolhidas), "completo": completo})
                r["completos" if completo else "parciais"] += 1
            else:
                ligados[ic] = antes
                r["sem_candidato"] += 1
                saida["sem_candidato"].append({"passo": ip, "cabo": ic, "motivo": "sem estrutura compatível" + (f" dentro de {raio} m" if raio else "")})
            if recusas and modo_conf != "ignorar":
                r["conflitos"] += 1
                saida["conflitos"].append({"passo": ip, "cabo": ic, "motivo": recusas[0]})
                if modo_conf == "parar":
                    saida["parou"] = True
                    break
        saida["passos"].append(r)
        if saida["parou"]:
            break
    return saida
