"""
services/regras_dominio.py — Camada 2 da validação: regras técnicas de domínio como DADOS (ADR-004, ADR-005).

Motor **v2** (TASK-016). As regras são editáveis pelo admin, por projeto, e avaliadas aqui. Função pura: sem I/O.

LINGUAGEM v2 (JSON)
-------------------
Regra de LINHA (cada linha da tabela Outros é um poste ou um conjunto de ativos soltos):

    {"id": "C2-CFU-SUPL", "versao": 2, "escopo": "linha", "operacoes": ["I", "*I"], "ativa": false,
     "severidade": "aviso", "mensagem": "...", "descricao": "...",
     "quando":   <condição>,      # gatilho (ausente = toda linha)
     "entao":    <condição>,      # o que DEVE valer quando o gatilho vale (ausente = "é proibido")
     "excecoes": <condição>}      # se valer, a regra não se aplica

  Há alerta quando `quando` vale, `excecoes` não vale e `entao` NÃO vale (ou não existe `entao`).

Regra de PLANILHA (agregado sobre Cabos e Outros, ex.: P50 total, total de postes):

    {"escopo": "planilha", "deve_ser": {"esq": <expr>, "cmp": ">=", "dir": <expr>, "rotulos": {"esq": "..", "dir": ".."}}}

CONDIÇÕES: {"tem": SEL, "qtd": CMP?} · {"soma": SEL, "qtd": CMP} · {"texto": "regex"} ·
           {"todos": [..]} · {"algum": [..]} · {"nenhum": [..]} · {"nao": COND}
SELETOR (SEL): "CFU" (código) · "TR[0-9]*" (curinga * ? [..]) · "@GRUPO" · {"regex": ".."} · {"e": [SEL, ..]} ·
               [SEL, ..] (qualquer um)
COMPARADOR (CMP): {"=": 1} {"!=": n} {">=": n} {"<=": n} {">": n} {"<": n} {"entre": [a, b]} (várias chaves = E)
VALOR de comparador: número ou {"base": n, "vezes": k, "mais": m, "se": COND}  → n, ou n*k+m quando `se` vale
EXPRESSÃO (planilha): número · {"soma_qtd": SEL} (Outros) · {"soma_metros": SEL} (Cabos, comprimento bruto) ·
                      {"soma": [expr, ..]} · {"vezes": [k, expr]}

Casamento de nomes sempre sem distinguir maiúsculas. `grupos` = {"NOME": SEL | [SEL, ..]} (globais no padrão,
redefiníveis por projeto — ver services/regras_camadas.py).

COMPATIBILIDADE: regras v1 (campo `tipo` + `parametros`, TASK-013) continuam sendo lidas: `converter_v1` as traduz para v2
de forma determinística e o resultado é idêntico (provado em tests/test_regras_v2.py contra o motor v1 congelado).
"""
import fnmatch
import re
from functools import lru_cache

from services.orcamento_calc import extrair_ativos_outros
from services.validacao_planilhas import OPERACOES_VALIDAS, _para_float, _tokens_cabo_tabela

VERSAO = 2
SEVERIDADES = {"erro", "aviso", "info"}
ESCOPOS_V2 = {"linha", "planilha"}
OPERACOES_PADRAO = ["I", "*I"]
COMPARADORES = ("=", "!=", ">=", "<=", ">", "<", "entre")
COMPARADORES_PLANILHA = ("=", "!=", ">=", "<=", ">", "<")
NOS_CONDICAO = ("tem", "soma", "texto", "todos", "algum", "nenhum", "nao")
CAMPOS_REGRA = {"id", "versao", "escopo", "operacoes", "ativa", "severidade", "mensagem", "descricao", "quando", "entao",
                "excecoes", "deve_ser", "origem", "oculta", "frase", "status_confirmacao"}
CAMPOS_META = {"origem", "oculta", "frase"}   # anotados pela camada de projetos/rotas; não fazem parte da regra
PROFUNDIDADE_MAX = 8
NOS_MAX = 200
TAM_REGEX_MAX = 500
_RE_NOME_GRUPO = re.compile(r"^[A-Z][A-Z0-9_]*$")


@lru_cache(maxsize=512)
def _re(padrao: str):
    return re.compile(padrao, re.IGNORECASE)


def _n(x):
    return f"{x:g}"


# ═══════════════════════════════════════════════════════════════════════════
# SELETORES E COMPARADORES
# ═══════════════════════════════════════════════════════════════════════════
def _lista(v):
    return v if isinstance(v, list) else [v]


def _casa(nome, sel, grupos, pilha=()):
    """O nome do ativo (já em maiúsculas) casa com o seletor?"""
    if isinstance(sel, list):
        return any(_casa(nome, s, grupos, pilha) for s in sel)
    if isinstance(sel, str):
        s = sel.strip()
        if s.startswith("@"):
            g = s[1:].upper()
            if g in pilha:
                return False  # ciclo (a validação já recusa; aqui só evita recursão infinita)
            return any(_casa(nome, m, grupos, pilha + (g,)) for m in _lista((grupos or {}).get(g, [])))
        s = s.upper()
        if any(c in s for c in "*?["):
            return fnmatch.fnmatchcase(nome, s)
        return nome == s
    if isinstance(sel, dict):
        if "regex" in sel:
            return bool(_re(sel["regex"]).search(nome))
        if "e" in sel:
            return all(_casa(nome, s, grupos, pilha) for s in sel["e"])
    return False


def _rotulo(sel):
    if isinstance(sel, list):
        return " ou ".join(_rotulo(s) for s in sel)
    if isinstance(sel, str):
        return sel.strip()
    if isinstance(sel, dict):
        if "regex" in sel:
            return f"/{sel['regex']}/"
        if "e" in sel:
            return " e ".join(_rotulo(s) for s in sel["e"])
    return "?"


def _valor(v, ctx):
    """Número, ou {"base","vezes","mais","se"} resolvido no contexto da linha."""
    if isinstance(v, dict):
        base = v.get("base", 0)
        if v.get("se") is not None and _cond(v["se"], ctx)[0]:
            return base * v.get("vezes", 1) + v.get("mais", 0)
        return base
    return v


def _cmp_ok(x, cmp_, ctx):
    for op, alvo in cmp_.items():
        if op == "entre":
            a, b = _valor(alvo[0], ctx), _valor(alvo[1], ctx)
            ok = a <= x <= b
        else:
            v = _valor(alvo, ctx)
            ok = {"=": x == v, "!=": x != v, ">=": x >= v, "<=": x <= v, ">": x > v, "<": x < v}[op]
        if not ok:
            return False
    return True


def _cmp_txt(cmp_, ctx):
    partes = []
    for op, alvo in cmp_.items():
        if op == "entre":
            partes.append(f"entre {_n(_valor(alvo[0], ctx))} e {_n(_valor(alvo[1], ctx))}")
        else:
            partes.append(f"{ {'>=': '≥', '<=': '≤', '!=': '≠'}.get(op, op) } {_n(_valor(alvo, ctx))}")
    return " e ".join(partes)


def _qtd_item(i):
    return f"{_n(i['qtd'])}-{i['ativo']}"


# ═══════════════════════════════════════════════════════════════════════════
# CONDIÇÕES (contexto = uma linha de Outros)
# ═══════════════════════════════════════════════════════════════════════════
class _Linha:
    def __init__(self, itens, texto, grupos):
        self.itens, self.texto, self.grupos = itens, texto, grupos


def _cond(c, ctx):
    """(vale?, fatos) — `fatos` descreve o que foi observado, para a explicação."""
    if "todos" in c:
        r = [_cond(x, ctx) for x in c["todos"]]
        return all(x[0] for x in r), [f for x in r for f in x[1]]
    if "algum" in c:
        r = [_cond(x, ctx) for x in c["algum"]]
        ok = any(x[0] for x in r)
        return ok, [f for x in r if (x[0] or not ok) for f in x[1]]
    if "nenhum" in c:
        r = [_cond(x, ctx) for x in c["nenhum"]]
        return not any(x[0] for x in r), [f for x in r for f in x[1]]
    if "nao" in c:
        ok, fatos = _cond(c["nao"], ctx)
        return not ok, fatos
    if "tem" in c:
        sel, cmp_ = c["tem"], c.get("qtd")
        casados = [i for i in ctx.itens if _casa(i["ativo"], sel, ctx.grupos)]
        certos = [i for i in casados if cmp_ is None or _cmp_ok(i["qtd"], cmp_, ctx)]
        if certos:
            return True, ["tem " + ", ".join(_qtd_item(i) for i in certos)]
        if casados:
            return False, [f"tem {', '.join(_qtd_item(i) for i in casados)}, mas a quantidade esperada é {_cmp_txt(cmp_, ctx)}"]
        return False, [f"não tem {_rotulo(sel)}" + (f" ({_cmp_txt(cmp_, ctx)})" if cmp_ else "")]
    if "soma" in c:
        sel, cmp_ = c["soma"], c["qtd"]
        s = sum(i["qtd"] for i in ctx.itens if _casa(i["ativo"], sel, ctx.grupos))
        ok = _cmp_ok(s, cmp_, ctx)
        return ok, [f"a soma de {_rotulo(sel)} é {_n(s)}" + ("" if ok else f" (esperado {_cmp_txt(cmp_, ctx)})")]
    if "texto" in c:
        ok = bool(_re(c["texto"]).search(ctx.texto))
        return ok, [f"o texto da linha {'casa' if ok else 'não casa'} com /{c['texto']}/"]
    return False, []


# ═══════════════════════════════════════════════════════════════════════════
# EXPRESSÕES (regras de planilha)
# ═══════════════════════════════════════════════════════════════════════════
class _Planilha:
    def __init__(self, itens_outros, cabos, grupos):
        self.itens_outros, self.cabos, self.grupos = itens_outros, cabos, grupos


def _expr(e, pl):
    if isinstance(e, (int, float)) and not isinstance(e, bool):
        return float(e)
    if "soma_qtd" in e:
        return sum(i["qtd"] for i in pl.itens_outros if _casa(i["ativo"], e["soma_qtd"], pl.grupos))
    if "soma_metros" in e:
        return sum(m for nome, m in pl.cabos if _casa(nome.upper(), e["soma_metros"], pl.grupos))
    if "soma" in e:
        return sum(_expr(x, pl) for x in e["soma"])
    if "vezes" in e:
        return e["vezes"][0] * _expr(e["vezes"][1], pl)
    if "n" in e:
        return float(e["n"])
    return 0.0


def _tabelas_expr(e):
    if isinstance(e, dict):
        if "soma_qtd" in e:
            return {"outros"}
        if "soma_metros" in e:
            return {"cabos"}
        if "soma" in e:
            return set().union(*[_tabelas_expr(x) for x in e["soma"]]) if e["soma"] else set()
        if "vezes" in e:
            return _tabelas_expr(e["vezes"][1])
    return set()


def _comprimento_bruto(cabo_ativo):
    tokens = _tokens_cabo_tabela(cabo_ativo)
    if len(tokens) < 2:
        return (tokens[0] if tokens else ""), 0.0
    valor = _para_float(tokens[-1])
    return tokens[0], (valor if valor is not None else 0.0)


# ═══════════════════════════════════════════════════════════════════════════
# COMPATIBILIDADE v1 → v2
# ═══════════════════════════════════════════════════════════════════════════
def _sel_regex(padrao):
    return {"regex": padrao}


def converter_v1(regra: dict) -> dict:
    """Traduz uma regra v1 (tipo + parametros) para v2, sem mudar o que ela detecta. Supõe regra v1 já validada."""
    p = regra.get("parametros") or {}
    tipo = regra.get("tipo")
    base = {"id": regra["id"], "versao": VERSAO, "ativa": regra.get("ativa", False), "severidade": regra["severidade"],
            "mensagem": regra["mensagem"], "operacoes": regra.get("operacoes") or list(OPERACOES_PADRAO)}
    if regra.get("descricao"):
        base["descricao"] = regra["descricao"]
    if regra.get("status_confirmacao"):
        base["status_confirmacao"] = regra["status_confirmacao"]

    if tipo == "minimo_total":
        contribuicoes = [{"vezes": [c["metros_por_unidade"], {"soma_qtd": _sel_regex(c["se_regex"])}]} for c in p["contribuicoes"]]
        return {**base, "escopo": "planilha", "deve_ser": {
            "esq": {"soma_metros": _sel_regex(p["ativo_regex"])}, "cmp": ">=", "dir": {"soma": contribuicoes}}}

    out = {**base, "escopo": "linha"}
    if tipo == "requer":
        minimo = p.get("qtd_min", 1)
        if p.get("dobra_se"):
            minimo = {"base": minimo, "vezes": p.get("multiplicador", 2), "se": {"algum": [
                {"soma": _sel_regex(c["regex"]), "qtd": {">=": c.get("qtd_min", 1)}} for c in p["dobra_se"]]}}
        out["quando"] = {"tem": _sel_regex(p["se_regex"])}
        out["entao"] = {"soma": _sel_regex(p["exige_regex"]), "qtd": {">=": minimo}}
        if p.get("exceto_se_regex"):
            out["excecoes"] = {"tem": _sel_regex(p["exceto_se_regex"])}
    elif tipo == "proibe":
        out["quando"] = {"todos": [{"tem": _sel_regex(p["se_regex"])}, {"tem": _sel_regex(p["com_regex"])}]}
        if p.get("exceto_se_regex"):
            out["excecoes"] = {"tem": _sel_regex(p["exceto_se_regex"])}
    elif tipo == "nao_isolado":
        minimo = {"base": p.get("min_outros", 1), "mais": 1,
                  "se": {"tem": {"e": [_sel_regex(p["se_regex"]), _sel_regex(p["acompanhado_por_regex"])]}}}
        out["quando"] = {"tem": _sel_regex(p["se_regex"])}
        out["entao"] = {"soma": _sel_regex(p["acompanhado_por_regex"]), "qtd": {">=": minimo}}
    elif tipo == "texto":
        out["quando"] = {"texto": p["texto_regex"]}
    return out


def normalizar_regra(regra: dict) -> dict:
    """Regra v2 como está; regra v1 (sem `versao` ou versao 1) convertida."""
    return regra if regra.get("versao") == VERSAO else converter_v1(regra)


# ═══════════════════════════════════════════════════════════════════════════
# AVALIAÇÃO
# ═══════════════════════════════════════════════════════════════════════════
def _ops(regra):
    return {str(o).strip().upper() for o in (regra.get("operacoes") or OPERACOES_PADRAO)}


def _tabela_de(regra):
    if regra.get("escopo") == "planilha":
        usadas = _tabelas_expr(regra["deve_ser"]["esq"]) | _tabelas_expr(regra["deve_ser"]["dir"])
        return "+".join(sorted(usadas, key=lambda t: t != "cabos")) if usadas else "cabos+outros"
    return "outros"


def _linha_regra(regra, ctx):
    """('alerta'|'ok'|'nao_aplica', explicação)"""
    q_ok, q_f = _cond(regra["quando"], ctx) if regra.get("quando") else (True, [])
    if not q_ok:
        return "nao_aplica", "; ".join(q_f) or "a condição da regra não vale nesta linha"
    if regra.get("excecoes"):
        e_ok, e_f = _cond(regra["excecoes"], ctx)
        if e_ok:
            return "nao_aplica", "exceção: " + "; ".join(e_f)
        sem_excecao = f" (sem exceção: {'; '.join(e_f)})"
    else:
        sem_excecao = ""
    if regra.get("entao") is None:
        return "alerta", "; ".join(q_f) + sem_excecao
    ok, f = _cond(regra["entao"], ctx)
    if ok:
        return "ok", "; ".join(q_f) + "; " + "; ".join(f)
    return "alerta", "; ".join(q_f) + ", mas " + "; ".join(f) + sem_excecao


def _planilha_regra(regra, cabos, outros, grupos):
    ops = _ops(regra)
    itens = []
    for linha in outros or []:
        if (linha.get("operacao") or "").strip().upper() in ops:
            itens += extrair_ativos_outros({"ativo": linha.get("ativo") or "", "operacao": ""})
    cabos_ok = []
    for linha in cabos or []:
        if (linha.get("operacao") or "").strip().upper() in ops and (linha.get("ativo") or "").strip():
            cabos_ok.append(_comprimento_bruto(linha["ativo"]))
    pl = _Planilha(itens, cabos_ok, grupos)
    d = regra["deve_ser"]
    esq, dir_ = _expr(d["esq"], pl), _expr(d["dir"], pl)
    ok = {"=": esq == dir_, "!=": esq != dir_, ">=": esq >= dir_, "<=": esq <= dir_, ">": esq > dir_, "<": esq < dir_}[d["cmp"]]
    rot = d.get("rotulos") or {}
    texto = f"{rot.get('esq', 'lado esquerdo')} = {_n(esq)}; {rot.get('dir', 'lado direito')} = {_n(dir_)}"
    return ("ok" if ok else "alerta"), texto + ("" if ok else f" (esperado: esquerdo {d['cmp']} direito)")


def _resultados(regras, cabos, outros, grupos):
    for regra in regras or []:
        r = normalizar_regra(regra)
        if not r.get("ativa") or r.get("oculta"):
            continue
        base = {"regra_id": r["id"], "severidade": r["severidade"], "mensagem": r["mensagem"], "tabela": _tabela_de(r)}
        if r.get("escopo") == "planilha":
            situacao, explicacao = _planilha_regra(r, cabos, outros, grupos)
            yield {**base, "linha_id": "GERAL", "situacao": situacao, "explicacao": explicacao}
            continue
        ops = _ops(r)
        for i, linha in enumerate(outros or []):
            if (linha.get("operacao") or "").strip().upper() not in ops:
                continue
            texto = (linha.get("ativo") or "").strip()
            if not texto:
                continue
            itens = extrair_ativos_outros({"ativo": texto, "operacao": ""})
            situacao, explicacao = _linha_regra(r, _Linha(itens, texto, grupos))
            yield {**base, "linha_id": linha.get("id") or f"OUTROS-{i}", "situacao": situacao, "explicacao": explicacao}


def avaliar(regras: list, cabos: list, outros: list, grupos: dict = None) -> list:
    """Achados da camada 2. Regras desligadas ou ocultas são ignoradas. `regras` deve ter passado por validar_regras."""
    return [{"linha_id": x["linha_id"], "tabela": x["tabela"], "regra_id": x["regra_id"], "severidade": x["severidade"],
             "mensagem": x["mensagem"], "explicacao": x["explicacao"]}
            for x in _resultados(regras, cabos, outros, grupos) if x["situacao"] == "alerta"]


def explicar(regras: list, cabos: list, outros: list, grupos: dict = None) -> list:
    """Uma entrada por (regra ligada, linha): situacao 'alerta' | 'ok' | 'nao_aplica' e o motivo em português."""
    return list(_resultados(regras, cabos, outros, grupos))


# ═══════════════════════════════════════════════════════════════════════════
# VALIDAÇÃO DO SCHEMA — nunca deixa salvar uma regra que o motor não saberia executar
# ═══════════════════════════════════════════════════════════════════════════
def _numero_ok(v):
    return isinstance(v, (int, float)) and not isinstance(v, bool)


def _regex_ok(valor, campo, erros):
    if not isinstance(valor, str) or not valor:
        erros.append(f"{campo}: regex obrigatória (texto).")
    elif len(valor) > TAM_REGEX_MAX:
        erros.append(f"{campo}: regex longa demais (máx. {TAM_REGEX_MAX} caracteres).")
    else:
        try:
            re.compile(valor)
        except re.error as e:
            erros.append(f"{campo}: regex inválida ({e}).")


def _ref_grupos(sel, grupos):
    """Nomes de grupos citados em um seletor (para validação e para 'usado por')."""
    if isinstance(sel, list):
        return [g for s in sel for g in _ref_grupos(s, grupos)]
    if isinstance(sel, str) and sel.strip().startswith("@"):
        return [sel.strip()[1:].upper()]
    if isinstance(sel, dict) and "e" in sel and isinstance(sel["e"], list):
        return [g for s in sel["e"] for g in _ref_grupos(s, grupos)]
    return []


def _validar_sel(sel, campo, erros, grupos, prof=0):
    if prof > PROFUNDIDADE_MAX:
        erros.append(f"{campo}: seletor aninhado demais.")
        return
    if isinstance(sel, list):
        if not sel:
            erros.append(f"{campo}: lista de seletores vazia.")
        for i, s in enumerate(sel):
            _validar_sel(s, f"{campo}[{i}]", erros, grupos, prof + 1)
    elif isinstance(sel, str):
        s = sel.strip()
        if not s or s == "@":
            erros.append(f"{campo}: seletor vazio.")
        elif s.startswith("@"):
            if s[1:].upper() not in (grupos or {}):
                erros.append(f"{campo}: grupo '{s}' não existe.")
    elif isinstance(sel, dict):
        if set(sel) == {"regex"}:
            _regex_ok(sel["regex"], f"{campo}.regex", erros)
        elif set(sel) == {"e"} and isinstance(sel["e"], list) and sel["e"]:
            for i, s in enumerate(sel["e"]):
                _validar_sel(s, f"{campo}.e[{i}]", erros, grupos, prof + 1)
        else:
            erros.append(f"{campo}: seletor inválido (use código, curinga, @GRUPO, {{regex}} ou {{e}}).")
    else:
        erros.append(f"{campo}: seletor inválido.")


def _validar_valor(v, campo, erros, grupos, cont):
    if _numero_ok(v):
        return
    if isinstance(v, dict) and set(v) <= {"base", "vezes", "mais", "se"} and _numero_ok(v.get("base")):
        for k in ("vezes", "mais"):
            if k in v and not _numero_ok(v[k]):
                erros.append(f"{campo}.{k}: precisa ser numérico.")
        if v.get("se") is not None:
            _validar_cond(v["se"], f"{campo}.se", erros, grupos, cont, 1)
        elif "vezes" in v or "mais" in v:
            erros.append(f"{campo}: 'vezes'/'mais' só têm efeito com 'se'.")
        return
    erros.append(f"{campo}: valor precisa ser um número ou {{base, vezes, mais, se}}.")


def _validar_cmp(cmp_, campo, erros, grupos, cont):
    if not isinstance(cmp_, dict) or not cmp_:
        erros.append(f"{campo}: comparador obrigatório, ex.: {{\">=\": 1}}.")
        return
    for op, alvo in cmp_.items():
        if op not in COMPARADORES:
            erros.append(f"{campo}: comparador '{op}' inválido (use {', '.join(COMPARADORES)}).")
        elif op == "entre":
            if not (isinstance(alvo, list) and len(alvo) == 2):
                erros.append(f"{campo}.entre: precisa ser [mínimo, máximo].")
            else:
                for i, x in enumerate(alvo):
                    _validar_valor(x, f"{campo}.entre[{i}]", erros, grupos, cont)
        else:
            _validar_valor(alvo, f"{campo}.{op}", erros, grupos, cont)


def _validar_cond(c, campo, erros, grupos, cont, prof=0):
    cont[0] += 1
    if cont[0] > NOS_MAX:
        if cont[0] == NOS_MAX + 1:
            erros.append(f"{campo}: regra grande demais (máx. {NOS_MAX} condições).")
        return
    if prof > PROFUNDIDADE_MAX:
        erros.append(f"{campo}: condições aninhadas demais (máx. {PROFUNDIDADE_MAX} níveis).")
        return
    if not isinstance(c, dict):
        erros.append(f"{campo}: condição precisa ser um objeto.")
        return
    chaves = [k for k in c if k in NOS_CONDICAO]
    if len(chaves) != 1 or (set(c) - {chaves[0]} - ({"qtd"} if chaves[0] in ("tem", "soma") else set())):
        erros.append(f"{campo}: condição inválida (use exatamente um de {', '.join(NOS_CONDICAO)}; 'qtd' só com tem/soma).")
        return
    k = chaves[0]
    if k in ("todos", "algum", "nenhum"):
        if not isinstance(c[k], list) or not c[k]:
            erros.append(f"{campo}.{k}: precisa ser uma lista com ao menos uma condição.")
        else:
            for i, x in enumerate(c[k]):
                _validar_cond(x, f"{campo}.{k}[{i}]", erros, grupos, cont, prof + 1)
    elif k == "nao":
        _validar_cond(c["nao"], f"{campo}.nao", erros, grupos, cont, prof + 1)
    elif k == "texto":
        _regex_ok(c["texto"], f"{campo}.texto", erros)
    else:  # tem / soma
        _validar_sel(c[k], f"{campo}.{k}", erros, grupos)
        if k == "soma" and "qtd" not in c:
            erros.append(f"{campo}: 'soma' exige 'qtd' (ex.: {{\">=\": 1}}).")
        if "qtd" in c:
            _validar_cmp(c["qtd"], f"{campo}.qtd", erros, grupos, cont)


def _validar_expr(e, campo, erros, grupos, prof=0):
    if prof > PROFUNDIDADE_MAX:
        erros.append(f"{campo}: expressão aninhada demais.")
        return
    if _numero_ok(e):
        return
    if not isinstance(e, dict) or len(e) != 1:
        erros.append(f"{campo}: expressão inválida (número, soma_qtd, soma_metros, soma ou vezes).")
        return
    k, v = next(iter(e.items()))
    if k in ("soma_qtd", "soma_metros"):
        _validar_sel(v, f"{campo}.{k}", erros, grupos)
    elif k == "soma":
        if not isinstance(v, list) or not v:
            erros.append(f"{campo}.soma: precisa ser uma lista de expressões.")
        else:
            for i, x in enumerate(v):
                _validar_expr(x, f"{campo}.soma[{i}]", erros, grupos, prof + 1)
    elif k == "vezes":
        if not (isinstance(v, list) and len(v) == 2 and _numero_ok(v[0])):
            erros.append(f"{campo}.vezes: precisa ser [número, expressão].")
        else:
            _validar_expr(v[1], f"{campo}.vezes[1]", erros, grupos, prof + 1)
    elif k == "n":
        if not _numero_ok(v):
            erros.append(f"{campo}.n: precisa ser numérico.")
    else:
        erros.append(f"{campo}: expressão '{k}' desconhecida.")


def validar_grupos(grupos) -> list:
    """Erros dos grupos: nomes, seletores válidos e sem ciclos."""
    if grupos is None:
        return []
    if not isinstance(grupos, dict):
        return ["'grupos' precisa ser um objeto {NOME: seletores}."]
    erros = []
    for nome, sels in grupos.items():
        if not isinstance(nome, str) or not _RE_NOME_GRUPO.match(nome):
            erros.append(f"Grupo '{nome}': nome deve usar MAIÚSCULAS, números e _ (ex.: ESTRUTURA_MT).")
            continue
        _validar_sel(_lista(sels), f"Grupo {nome}", erros, grupos)
    # ciclos
    def alcanca(inicio, atual, visto):
        for ref in _ref_grupos(_lista(grupos.get(atual, [])), grupos):
            if ref == inicio:
                return True
            if ref in grupos and ref not in visto and alcanca(inicio, ref, visto | {ref}):
                return True
        return False
    for nome in grupos:
        if isinstance(nome, str) and alcanca(nome.upper(), nome.upper(), set()):
            erros.append(f"Grupo {nome}: referência circular entre grupos.")
    return erros


def _validar_regra_v2(r, prefixo, erros, grupos):
    for k in r:
        if k not in CAMPOS_REGRA:
            erros.append(f"{prefixo}: campo desconhecido '{k}'.")
    if r.get("escopo") not in ESCOPOS_V2:
        erros.append(f"{prefixo}: 'escopo' precisa ser 'linha' ou 'planilha'.")
        return
    cont = [0]
    if r["escopo"] == "linha":
        if r.get("quando") is None and r.get("entao") is None:
            erros.append(f"{prefixo}: informe 'quando' e/ou 'entao' (senão a regra alertaria em toda linha).")
        if r.get("deve_ser") is not None:
            erros.append(f"{prefixo}: 'deve_ser' é só de regra de planilha.")
        for campo in ("quando", "entao", "excecoes"):
            if r.get(campo) is not None:
                _validar_cond(r[campo], f"{prefixo}.{campo}", erros, grupos, cont)
    else:
        for campo in ("quando", "entao", "excecoes"):
            if r.get(campo) is not None:
                erros.append(f"{prefixo}: '{campo}' é só de regra de linha.")
        d = r.get("deve_ser")
        if not isinstance(d, dict) or not {"esq", "cmp", "dir"} <= set(d) or set(d) - {"esq", "cmp", "dir", "rotulos"}:
            erros.append(f"{prefixo}: 'deve_ser' precisa ser {{esq, cmp, dir}} (e opcional 'rotulos').")
        else:
            if d["cmp"] not in COMPARADORES_PLANILHA:
                erros.append(f"{prefixo}.deve_ser.cmp: use um de {', '.join(COMPARADORES_PLANILHA)}.")
            _validar_expr(d["esq"], f"{prefixo}.deve_ser.esq", erros, grupos)
            _validar_expr(d["dir"], f"{prefixo}.deve_ser.dir", erros, grupos)
            if d.get("rotulos") is not None and not (isinstance(d["rotulos"], dict) and all(isinstance(x, str) for x in d["rotulos"].values())):
                erros.append(f"{prefixo}.deve_ser.rotulos: precisa ser {{esq, dir}} com textos.")


# — validação das regras v1 (TASK-013), mantida para regras antigas ainda não convertidas —
def _regex_v1(valor, campo, erros, obrigatorio=True):
    if valor in (None, ""):
        if obrigatorio:
            erros.append(f"{campo}: obrigatório.")
        return
    if not isinstance(valor, str):
        erros.append(f"{campo}: precisa ser texto.")
        return
    try:
        re.compile(valor)
    except re.error as e:
        erros.append(f"{campo}: regex inválida ({e}).")


def _numero_v1(valor, campo, erros, minimo=None):
    if isinstance(valor, bool) or not isinstance(valor, (int, float)):
        erros.append(f"{campo}: precisa ser numérico.")
    elif minimo is not None and valor < minimo:
        erros.append(f"{campo}: precisa ser >= {minimo}.")


def _validar_parametros_v1(tipo, p, prefixo, erros):
    if tipo == "requer":
        _regex_v1(p.get("se_regex"), f"{prefixo}.se_regex", erros)
        _regex_v1(p.get("exige_regex"), f"{prefixo}.exige_regex", erros)
        _regex_v1(p.get("exceto_se_regex"), f"{prefixo}.exceto_se_regex", erros, obrigatorio=False)
        if "qtd_min" in p:
            _numero_v1(p["qtd_min"], f"{prefixo}.qtd_min", erros, minimo=0)
        if "multiplicador" in p:
            _numero_v1(p["multiplicador"], f"{prefixo}.multiplicador", erros, minimo=1)
        dobra = p.get("dobra_se")
        if dobra is not None:
            if not isinstance(dobra, list):
                erros.append(f"{prefixo}.dobra_se: precisa ser uma lista.")
            else:
                for i, cond in enumerate(dobra):
                    if not isinstance(cond, dict):
                        erros.append(f"{prefixo}.dobra_se[{i}]: precisa ser um objeto {{regex, qtd_min}}.")
                        continue
                    _regex_v1(cond.get("regex"), f"{prefixo}.dobra_se[{i}].regex", erros)
                    _numero_v1(cond.get("qtd_min", 1), f"{prefixo}.dobra_se[{i}].qtd_min", erros, minimo=0)
    elif tipo == "proibe":
        _regex_v1(p.get("se_regex"), f"{prefixo}.se_regex", erros)
        _regex_v1(p.get("com_regex"), f"{prefixo}.com_regex", erros)
        _regex_v1(p.get("exceto_se_regex"), f"{prefixo}.exceto_se_regex", erros, obrigatorio=False)
    elif tipo == "nao_isolado":
        _regex_v1(p.get("se_regex"), f"{prefixo}.se_regex", erros)
        _regex_v1(p.get("acompanhado_por_regex"), f"{prefixo}.acompanhado_por_regex", erros)
        if "min_outros" in p:
            _numero_v1(p["min_outros"], f"{prefixo}.min_outros", erros, minimo=1)
    elif tipo == "texto":
        _regex_v1(p.get("texto_regex"), f"{prefixo}.texto_regex", erros)
    elif tipo == "minimo_total":
        _regex_v1(p.get("ativo_regex"), f"{prefixo}.ativo_regex", erros)
        contrib = p.get("contribuicoes")
        if not isinstance(contrib, list) or not contrib:
            erros.append(f"{prefixo}.contribuicoes: precisa ser uma lista com ao menos uma contribuição.")
        else:
            for i, c in enumerate(contrib):
                if not isinstance(c, dict):
                    erros.append(f"{prefixo}.contribuicoes[{i}]: precisa ser um objeto.")
                    continue
                _regex_v1(c.get("se_regex"), f"{prefixo}.contribuicoes[{i}].se_regex", erros)
                _numero_v1(c.get("metros_por_unidade"), f"{prefixo}.contribuicoes[{i}].metros_por_unidade", erros, minimo=0)


TIPOS_V1 = {"requer", "proibe", "nao_isolado", "texto", "minimo_total"}
ESCOPOS_V1 = {"outros", "cabos+outros"}


def _validar_regra_v1(r, prefixo, erros):
    tipo = r.get("tipo")
    if tipo not in TIPOS_V1:
        erros.append(f"{prefixo}: 'tipo' precisa ser um de {sorted(TIPOS_V1)} (recebido: {tipo!r}).")
    if r.get("escopo", "outros") not in ESCOPOS_V1:
        erros.append(f"{prefixo}: 'escopo' precisa ser um de {sorted(ESCOPOS_V1)}.")
    if tipo == "minimo_total" and r.get("escopo") != "cabos+outros":
        erros.append(f"{prefixo}: 'minimo_total' exige escopo 'cabos+outros'.")
    params = r.get("parametros")
    if not isinstance(params, dict):
        erros.append(f"{prefixo}: 'parametros' precisa ser um objeto.")
    elif tipo in TIPOS_V1:
        _validar_parametros_v1(tipo, params, f"{prefixo}.parametros", erros)


def validar_regras(regras, grupos=None) -> list:
    """Lista de erros (vazia = válido) de uma lista de regras v1 e/ou v2. `grupos` = grupos efetivos (para @GRUPO)."""
    if not isinstance(regras, list):
        return ["O payload de regras precisa ser uma lista."]
    erros, ids = [], set()
    for idx, r in enumerate(regras):
        prefixo = f"Regra #{idx + 1}"
        if not isinstance(r, dict):
            erros.append(f"{prefixo}: precisa ser um objeto.")
            continue
        rid = r.get("id")
        if not isinstance(rid, str) or not rid.strip():
            erros.append(f"{prefixo}: 'id' é obrigatório.")
        elif rid in ids:
            erros.append(f"{prefixo}: 'id' duplicado ({rid}).")
        else:
            ids.add(rid)
            prefixo = f"Regra {rid}"
        if r.get("severidade") not in SEVERIDADES:
            erros.append(f"{prefixo}: 'severidade' precisa ser um de {sorted(SEVERIDADES)}.")
        if not isinstance(r.get("mensagem"), str) or not r["mensagem"].strip():
            erros.append(f"{prefixo}: 'mensagem' é obrigatória.")
        if not isinstance(r.get("ativa"), bool):
            erros.append(f"{prefixo}: 'ativa' precisa ser true ou false.")
        ops = r.get("operacoes")
        if ops is not None and (not isinstance(ops, list) or not ops
                                or any(str(o).upper() not in OPERACOES_VALIDAS for o in ops)):
            erros.append(f"{prefixo}: 'operacoes' precisa ser uma lista com valores de {sorted(OPERACOES_VALIDAS)}.")
        versao = r.get("versao", 1)
        if versao == VERSAO:
            _validar_regra_v2(r, prefixo, erros, grupos)
        elif versao == 1:
            _validar_regra_v1(r, prefixo, erros)
        else:
            erros.append(f"{prefixo}: 'versao' precisa ser 1 ou 2.")
    return erros


def validar_conjunto(regras, grupos=None) -> list:
    """Grupos + regras (com os grupos disponíveis para @GRUPO)."""
    return validar_grupos(grupos) + validar_regras(regras, grupos or {})


def grupos_usados(regra: dict) -> set:
    """Nomes dos grupos citados por uma regra (v1 não cita grupos)."""
    r = normalizar_regra(regra) if regra.get("versao", 1) in (1, VERSAO) else regra
    achados = set()

    def visita(x):
        if isinstance(x, dict):
            for k, v in x.items():
                if k in ("tem", "soma", "soma_qtd", "soma_metros"):
                    achados.update(_ref_grupos(v, {}))
                elif k == "e":
                    achados.update(_ref_grupos(x, {}))
                visita(v)
        elif isinstance(x, list):
            for v in x:
                visita(v)
    visita({k: v for k, v in r.items() if k in ("quando", "entao", "excecoes", "deve_ser")})
    return achados


# ═══════════════════════════════════════════════════════════════════════════
# DESCRIÇÃO EM PORTUGUÊS (base do editor visual)
# ═══════════════════════════════════════════════════════════════════════════
def _frase_sel(sel):
    if isinstance(sel, dict) and "regex" in sel:
        return f"o padrão /{sel['regex']}/"
    if isinstance(sel, str) and sel.strip().startswith("@"):
        return f"um item do grupo {sel.strip()[1:].upper()}"
    if isinstance(sel, list):
        return " ou ".join(_frase_sel(s) for s in sel)
    return _rotulo(sel)


def _frase_valor(v):
    if isinstance(v, dict):
        texto = _n(v.get("base", 0))
        if v.get("se") is not None:
            extra = []
            if v.get("vezes", 1) != 1:
                extra.append(f"multiplicado por {_n(v['vezes'])}")
            if v.get("mais", 0):
                extra.append(f"mais {_n(v['mais'])}")
            texto += f" ({' e '.join(extra)} se a linha {_frase_situacao(v['se'])})"
        return texto
    return _n(v)


def _frase_cmp(cmp_):
    out = []
    for op, alvo in cmp_.items():
        if op == "entre":
            out.append(f"entre {_frase_valor(alvo[0])} e {_frase_valor(alvo[1])}")
        else:
            out.append(f"{ {'>=': '≥', '<=': '≤', '!=': '≠'}.get(op, op) } {_frase_valor(alvo)}")
    return " e ".join(out)


def _junta(partes, conector, nivel):
    t = f" {conector} ".join(partes)
    return f"({t})" if nivel and len(partes) > 1 else t


def _frase_situacao(c, nivel=0):
    """A condição como afirmação: 'tem CFU com quantidade = 1', 'a soma de SUPL é ≥ 1'."""
    if "todos" in c:
        return _junta([_frase_situacao(x, nivel + 1) for x in c["todos"]], "e", nivel)
    if "algum" in c:
        return _junta([_frase_situacao(x, nivel + 1) for x in c["algum"]], "ou", nivel)
    if "nenhum" in c:
        return "não tem nenhum destes: " + "; ".join(_frase_situacao(x, nivel + 1) for x in c["nenhum"])
    if "nao" in c:
        return "não é verdade que " + _frase_situacao(c["nao"], nivel + 1)
    if "tem" in c:
        return f"tem {_frase_sel(c['tem'])}" + (f" com quantidade {_frase_cmp(c['qtd'])}" if c.get("qtd") else "")
    if "soma" in c:
        return f"tem, no total, {_frase_sel(c['soma'])} {_frase_cmp(c['qtd'])}"
    if "texto" in c:
        return f"tem texto que casa com /{c['texto']}/"
    return "?"


def _frase_exigencia(c, nivel=0):
    """A mesma condição como exigência: 'deve ter SUPL', 'a soma de PR220 deve ser ≥ 2', 'não deve ter EF3H'."""
    if "todos" in c:
        return _junta([_frase_exigencia(x, nivel + 1) for x in c["todos"]], "e", nivel)
    if "algum" in c:
        return "pelo menos uma destas: " + "; ".join(_frase_exigencia(x, nivel + 1) for x in c["algum"])
    if "nenhum" in c:
        return "nenhuma destas: " + "; ".join(_frase_situacao(x, nivel + 1) for x in c["nenhum"])
    if "nao" in c:
        i = c["nao"]
        if "tem" in i:
            return f"não deve ter {_frase_sel(i['tem'])}" + (f" com quantidade {_frase_cmp(i['qtd'])}" if i.get("qtd") else "")
        return "não deve valer: " + _frase_situacao(i, nivel + 1)
    if "tem" in c:
        return f"deve ter {_frase_sel(c['tem'])}" + (f" com quantidade {_frase_cmp(c['qtd'])}" if c.get("qtd") else "")
    if "soma" in c:
        return f"a soma de {_frase_sel(c['soma'])} deve ser {_frase_cmp(c['qtd'])}"
    if "texto" in c:
        return f"o texto deve casar com /{c['texto']}/"
    return "?"


def _frase_expr(e):
    if _numero_ok(e):
        return _n(e)
    k, v = next(iter(e.items()))
    if k == "soma_qtd":
        return f"a soma das quantidades de {_frase_sel(v)} (Outros)"
    if k == "soma_metros":
        return f"a soma dos metros de {_frase_sel(v)} (Cabos)"
    if k == "soma":
        return "(" + " + ".join(_frase_expr(x) for x in v) + ")"
    if k == "vezes":
        return f"{_n(v[0])} × {_frase_expr(v[1])}"
    return _n(v)


def descrever(regra: dict, grupos: dict = None) -> str:
    """Frase em português da regra (v1 ou v2). Supõe regra válida."""
    r = normalizar_regra(regra)
    ops = ", ".join(r.get("operacoes") or OPERACOES_PADRAO)
    gravidade = f" Gravidade: {r.get('severidade', 'aviso')}."
    if r.get("escopo") == "planilha":
        d = r["deve_ser"]
        simbolo = {">=": "≥", "<=": "≤", "!=": "≠"}.get(d["cmp"], d["cmp"])
        return (f"Na planilha inteira (linhas com operação {ops}): {_frase_expr(d['esq'])} deve ser {simbolo} "
                f"{_frase_expr(d['dir'])}.{gravidade}")
    inicio = f"Em cada linha de Outros (operação {ops}): "
    if r.get("entao") is None:
        corpo = f"alerta quando a linha {_frase_situacao(r['quando'])}"
    elif r.get("quando"):
        corpo = f"se a linha {_frase_situacao(r['quando'])}, então {_frase_exigencia(r['entao'])}"
    else:
        corpo = _frase_exigencia(r["entao"])
    if r.get("excecoes"):
        corpo += f"; exceto se a linha {_frase_situacao(r['excecoes'])}"
    return f"{inicio}{corpo}.{gravidade}"
