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

from services.prazo import verificar_prazo
from services import anotacao_cabos as _ac
from services import convencao_postes as _cp
from services import dicionario_cabos as _dc
from services import perfil_exportacao as _pe

ORDEM_MAIOR_PADRAO = ["C", "P", "A"]
SE_DIFEREM = ("BLOQUEAR", "MAIOR", "AVISAR")
SE_FALTAR = ("AVISAR", "BLOQUEAR")                         # TASK-169
ORDENS = ("ORIGEM", "MAIOR_PRIMEIRO")
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
    return {"versao": 2, "cabos": [], "estruturas": [], "ordem_maior": list(ORDEM_MAIOR_PADRAO), "scripts": [], "dicionario_cabos": []}


def normalizar_config(cfg) -> dict:
    c = config_vazia()
    if isinstance(cfg, dict):
        c["cabos"] = [r for r in (cfg.get("cabos") or []) if isinstance(r, dict)]
        c["estruturas"] = [r for r in (cfg.get("estruturas") or []) if isinstance(r, dict)]
        c["dicionario_cabos"] = _dc.normalizar(cfg.get("dicionario_cabos"))                              # TASK-173
        if isinstance(cfg.get("postes"), dict):                                                          # TASK-174: convenção de postes (sem ela, a de Pernambuco)
            c["postes"] = cfg["postes"]
        if isinstance(cfg.get("exportacao"), dict):                                                      # TASK-176: perfil de exportação (sem ele, o de Neoenergia)
            c["exportacao"] = cfg["exportacao"]
        om = cfg.get("ordem_maior")
        if isinstance(om, list) and om:
            c["ordem_maior"] = [_trim(x).upper() for x in om]
        elif _dc.ordem_maior_padrao(c["dicionario_cabos"]):                                              # TASK-173: sem ordem própria, vale a do dicionário
            c["ordem_maior"] = _dc.ordem_maior_padrao(c["dicionario_cabos"])
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
    erros.extend(_dc.validar(cfg.get("dicionario_cabos")))                                              # TASK-173
    erros.extend(_cp.validar(cfg.get("postes")))                                                         # TASK-174
    erros.extend(_pe.validar(cfg.get("exportacao")))                                                     # TASK-176
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
                pr = r.get("parametros")
                if pr is not None and pr != "":
                    np_ = normalizar_parametros(pr)
                    brutos = len(pr) if isinstance(pr, list) else len(str(pr).split(","))
                    if not np_ or len(np_) != brutos or any(not re.fullmatch(r"[A-Z0-9_.]+", x["nome"]) or x["fonte"] not in FONTES_PARAMETRO for x in np_):
                        erros.append(f"{p}: 'parametros' precisa ser uma lista como \"cabo=ativo, cabo=ativo, tipocabo=valor_cabo\" (fonte: descricao, valor_cabo ou ativo).")
                if r.get("lados") not in (None, "IGUAIS", "LIVRE", "DERIVACAO"):
                    erros.append(f"{p}: 'lados' precisa ser IGUAIS, LIVRE ou DERIVACAO.")
                cv = r.get("cabos_vaos")                                                                  # TASK-170
                if cv is not None and cv != "":
                    if not isinstance(cv, list):
                        erros.append(f"{p}: 'cabos_vaos' precisa ser uma lista de regex (uma por vão).")
                    else:
                        for x in cv:
                            try:
                                re.compile(str(x))
                            except re.error as e:
                                erros.append(f"{p}: 'cabos_vaos' tem regex inválida ({e}).")
                df = r.get("define")
                if df is not None and df != "":
                    if not isinstance(df, list):
                        erros.append(f"{p}: 'define' precisa ser uma lista de {{linha_tem?, cabo?, param}}.")
                    else:
                        for k, it in enumerate(df):
                            q = f"{p}: define #{k + 1}"
                            if not isinstance(it, dict):
                                erros.append(f"{q}: precisa ser um objeto.")
                                continue
                            if it.get("linha_tem") in (None, "", []) and not _trim(it.get("cabo")):
                                erros.append(f"{q}: precisa de pelo menos uma condição (linha_tem ou cabo).")
                            if not isinstance(it.get("param"), str) or not re.fullmatch(r"\s*[A-Za-z0-9_.]+\s*=\s*[A-Za-z0-9_./-]+\s*", it.get("param") or ""):
                                erros.append(f"{q}: 'param' precisa ser nome=valor (ex.: tipo=RL).")
                            for campo in ("linha_tem", "cabo"):
                                v = it.get(campo)
                                if isinstance(v, str) and _RE_REGEX.search(v):
                                    try:
                                        re.compile(v)
                                    except re.error as e:
                                        erros.append(f"{q}: '{campo}' regex inválida ({e}).")
                if r.get("se_diferem") not in (None,) + SE_DIFEREM:
                    erros.append(f"{p}: 'se_diferem' precisa ser um de {list(SE_DIFEREM)}.")
                for k in ("ocupa", "distintos"):                                                          # TASK-169
                    if r.get(k) is not None and not isinstance(r[k], bool):
                        erros.append(f"{p}: '{k}' precisa ser true ou false.")
                if r.get("herda") not in (None, "", "POSTE"):
                    erros.append(f"{p}: 'herda' precisa ser \"POSTE\" (ou vazio).")
                if r.get("herda") == "POSTE" and r.get("ocupa") is not False:
                    erros.append(f"{p}: 'herda' só vale com 'ocupa': false.")
                if _trim(r.get("nivel")) and _trim(r.get("nivel")).upper() not in ("MT", "BT"):
                    erros.append(f"{p}: 'nivel' precisa ser MT ou BT.")
                if r.get("ordem") not in (None, "") + ORDENS:
                    erros.append(f"{p}: 'ordem' precisa ser um de {list(ORDENS)}.")
                if r.get("se_faltar") not in (None, "") + SE_FALTAR:
                    erros.append(f"{p}: 'se_faltar' precisa ser um de {list(SE_FALTAR)}.")
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


def derivar_cabo(texto, dic=None) -> dict:
    """`dic` = dicionário de cabos do projeto (TASK-173): sem tensão no texto, o nível e as fases padrão vêm da linha do cabo; sem dicionário, BT e 1 fase."""
    t = _trim(texto)
    mf, mt = _RE_FASES.search(t), _RE_TENSAO.search(t)
    tensao = mt.group(1).replace(".", ",") + "kV" if mt else ""
    fases, nivel = int(mf.group(1)) if mf else 1, "MT" if mt and float(mt.group(1).replace(",", ".")) > 1 else "BT"
    if dic and (not mt or not mf):
        itens = _ac.interpretar(t, dic)["itens"]
        p = ([i for i in itens if not i["neutro"]] or itens or [None])[0]
        if p:
            if not mt:
                nivel = _dc.nivel_padrao(dic, p["ativo"]) or nivel
            if not mf:
                fases = _dc.fases_padrao(dic, p["ativo"]) or fases
    return {"fases": fases, "tensao": tensao, "nivel": nivel}


def extra_cabo(row, dic=None) -> dict:
    """Dados que as regras do cabo podem usar: o que está na linha (editável) ou, na falta, o derivado do texto original."""
    row = row or {}
    d = derivar_cabo(row.get("textoOriginal"), dic)
    try:
        f = int(_trim(row.get("fases"))) if _trim(row.get("fases")) != "" else d["fases"]
    except ValueError:
        f = d["fases"]
    return {"texto": _trim(row.get("textoOriginal")), "descricao": descricao_cabo(row.get("textoOriginal") or row.get("ativo") or ""), "fases": f, "tensao": _trim(row.get("tensao")) or d["tensao"], "nivel": (_trim(row.get("nivel")).upper() or d["nivel"])}


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
    vc = (_trim(r.get("valor_cabo")).upper() or None) if r else None
    dic = cfg.get("dicionario_cabos")
    if not vc and dic:                                                                                     # TASK-173: sem valor na Tabela de Cabos, vale o do dicionário
        vc = _dc.valor_cabo(dic, cod)
    return {"codigo": cod, "regra": i, "valor_cabo": vc, "pontas": pontas, "secao": secao_cabo(cod, dic),
            "vincula": not (r is not None and (r.get("vincula") is False)), "estruturas": (r.get("estruturas") if r else None),
            "operacoes": (r.get("operacoes") if r else None), "op": _trim(operacao).upper(), "base": base,
            "fases": extra["fases"] if extra else None, "nivel": extra["nivel"] if extra else None,
            "descricao": extra["descricao"] if extra and extra.get("descricao") else cod}


def info_cabo_ficticio(codigo, estrutura, operacao, cfg) -> dict:
    """TASK-170: o cabo fictício de uma estrutura recebe o NÍVEL da primeira regra de cabo que casa com o código E aceita a estrutura (S04 numa LN3 = BT; numa N3 = MT);
    se essa regra não tem `nivel`, fica sem nível (como antes)."""
    cod = _trim(codigo).upper()
    info = info_cabo(cod, operacao, cfg)
    for r in cfg["cabos"]:
        if r.get("ativa") is False or not casa(r.get("padrao"), cod) or not casa(r.get("estruturas"), estrutura):
            continue
        n = _trim(r.get("nivel")).upper()
        info["nivel"] = n if n in ("MT", "BT") else None
        break
    return info


# TASK-096: parâmetros da estrutura com FONTE por cabo vinculado: "cabo=ativo, cabo=ativo, tipocabo=valor_cabo" (ativo = código do cabo sem 3x/tensão; valor_cabo = tipo da regra do cabo).
# O mesmo nome pode se repetir: a n-ésima ocorrência de um nome recebe o n-ésimo cabo (a ordem dos cabos não importa para o orçamento).
# TASK-100: o NOME do parâmetro define o que ele recebe — `cabo` = a DESCRIÇÃO COMPLETA do cabo (sem o comprimento, ex.: 3xS04-7,97kV → 3XS04-7.97KV);
# `tipocabo` = o VALOR DE CABO da Tabela de Cabos (C, P, A…); outro nome = o código do cabo. `nome=fonte` força a fonte (descricao | valor_cabo | ativo).
FONTES_PARAMETRO = ("ativo", "valor_cabo", "descricao")
_RE_COMPRIMENTO = re.compile(r"\s+\d+(?:[.,]\d+)?\s*m?\s*$", re.I)      # o "m" é opcional (RAMA25 12)


def _fonte_do_nome(nome) -> str:
    return "descricao" if re.fullmatch(r"CABO\d*", nome) else ("valor_cabo" if re.fullmatch(r"TIPOCABO\d*", nome) else "ativo")      # cabo, cabo2, cabo3… / tipocabo, tipocabo2…


def descricao_cabo(texto, codigo="") -> str:
    """Descrição completa do cabo, sem o comprimento, no formato aceito num parâmetro: maiúsculas, sem espaços, vírgula decimal → ponto, parênteses → "_"."""
    d = re.sub(r"\s+", "", _RE_COMPRIMENTO.sub("", _trim(texto))).upper().replace(",", ".")
    d = re.sub(r"[()\[\]]", "_", d)
    return d or _trim(codigo).upper()


def normalizar_parametros(v):
    if v is None or v == "":
        return None
    itens = v if isinstance(v, list) else str(v).split(",")
    lista = []
    for x in itens:
        if isinstance(x, dict):
            nome = _trim(x.get("nome")).upper()
            fonte = _trim(x.get("fonte")).lower() or _fonte_do_nome(nome)
        else:
            n, _, f = str(x).partition("=")
            nome = _trim(n).upper()
            fonte = _trim(f).lower() or _fonte_do_nome(nome)
        if nome:
            lista.append({"nome": nome, "fonte": fonte})
    return lista or None


def _normalizar_define(v) -> list:
    """TASK-170: [{"linha_tem", "cabo", "param": "nome=valor"}] → [{"linha_tem", "cabo", "param": (NOME, VALOR)}] (item sem condição ou sem param válido é ignorado)."""
    saida = []
    if not isinstance(v, list):
        return saida
    for it in v:
        if not isinstance(it, dict) or not isinstance(it.get("param"), str) or "=" not in it["param"]:
            continue
        nome, _, valor = it["param"].partition("=")
        nome, valor = _trim(nome).upper(), _trim(valor).upper()
        lt, cb = it.get("linha_tem"), it.get("cabo")
        lt = lt if (isinstance(lt, (str, list)) and lt != "" and lt != []) else None
        cb = cb if (isinstance(cb, str) and _trim(cb)) else None
        if not nome or not valor or (lt is None and cb is None):
            continue
        saida.append({"linha_tem": lt, "cabo": cb, "param": (nome, valor)})
    return saida


def info_estrutura(codigo, cfg) -> dict:
    cod = _trim(codigo).upper()
    i, r = _primeira(cfg["estruturas"], "padrao", cod) if cod else (None, None)
    d = {"codigo": cod, "regra": i, "qtd": None, "qtd_min": None, "max_reais": None, "parametros": None, "lados": "LIVRE", "amarracao": False, "aceita_i": True, "cabos": None, "parametro": "CABO", "se_diferem": "MAIOR",
         "ocupa": True, "herda": None, "nivel": None, "ordem": "ORIGEM", "distintos": False, "se_faltar": "AVISAR", "cabos_vaos": [], "define": []}
    if r is not None:
        # TASK-093: qtd_cabos fixo (min = max) OU faixa qtd_min..qtd_max; max_reais = quantos podem ser cabos REAIS (o excedente só entra como fictício)
        d["qtd"] = r["qtd_max"] if isinstance(r.get("qtd_max"), int) else (r.get("qtd_cabos") if isinstance(r.get("qtd_cabos"), int) else None)
        d["qtd_min"] = r["qtd_min"] if isinstance(r.get("qtd_min"), int) else d["qtd"]
        d["max_reais"] = r["max_reais"] if isinstance(r.get("max_reais"), int) else None
        d["lados"] = r["lados"] if r.get("lados") in ("IGUAIS", "DERIVACAO") else "LIVRE"          # TASK-170: DERIVACAO = 1 derivação + até 2 do tronco (iguais)
        d["amarracao"] = bool(r.get("amarracao", False))
        d["aceita_i"] = r.get("aceita_i") is not False
        d["cabos"] = r.get("cabos")
        d["parametro"] = (_trim(r.get("parametro")) or "cabo").upper() if r.get("parametro") != "" else ""
        d["parametros"] = normalizar_parametros(r.get("parametros"))
        if d["parametros"]:
            d["parametro"] = ""                                  # a lista substitui o parâmetro único
        d["se_diferem"] = r.get("se_diferem") if r.get("se_diferem") in SE_DIFEREM else "MAIOR"
        # TASK-169: equipamento (RL, CFU, para-raio…) que NÃO ocupa pontas e lê os cabos do poste; todos opcionais (ausentes = comportamento de antes)
        d["ocupa"] = r.get("ocupa") is not False
        d["herda"] = "POSTE" if _trim(r.get("herda")).upper() == "POSTE" else None
        d["nivel"] = _trim(r.get("nivel")).upper() if _trim(r.get("nivel")).upper() in ("MT", "BT") else None
        d["ordem"] = "MAIOR_PRIMEIRO" if _trim(r.get("ordem")).upper() == "MAIOR_PRIMEIRO" else "ORIGEM"
        d["distintos"] = r.get("distintos") is True
        d["se_faltar"] = "BLOQUEAR" if _trim(r.get("se_faltar")).upper() == "BLOQUEAR" else "AVISAR"
        # TASK-170: um cabo por vão (cada regex precisa de um cabo DIFERENTE) e parâmetros automáticos pela linha do poste / pelo 1º cabo
        vaos = r.get("cabos_vaos")
        if isinstance(vaos, list) and any(_trim(x) for x in vaos):
            d["cabos_vaos"] = [_trim(x) for x in vaos if _trim(x)]
            d["qtd"] = d["qtd_min"] = len(d["cabos_vaos"])
        d["define"] = _normalizar_define(r.get("define"))
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
    if not (any(casa(v, c["codigo"]) for v in e["cabos_vaos"]) if e["cabos_vaos"] else casa(e["cabos"], c["codigo"])):
        return False, f"a estrutura {e['codigo']} não aceita o cabo {c['codigo']}"
    if c["base"] == "M" and e["regra"] is not None and not e["amarracao"]:
        return False, f"cabo M só se liga a estrutura de amarração ({e['codigo']} não é)"
    if c["base"] == "I" and e["regra"] is not None and not e["aceita_i"]:
        return False, f"a estrutura {e['codigo']} só aceita cabo M"
    return True, None


def _atribuicao_vaos(vaos, codigos):
    """TASK-170: casamento um a um — cada cabo num vão DIFERENTE cuja regex ele satisfaz. → [índice do vão de cada cabo] ou None."""
    n = len(codigos)

    def tenta(i, usados):
        if i == n:
            return []
        for j, v in enumerate(vaos):
            if j in usados or not casa(v, codigos[i]):
                continue
            resto = tenta(i + 1, usados | {j})
            if resto is not None:
                return [j] + resto
        return None
    return tenta(0, frozenset())


def _par_do_tronco(codigos):
    """TASK-170: índices (i, j) do primeiro par de cabos iguais (o tronco da derivação) ou None."""
    for i in range(len(codigos)):
        for j in range(i + 1, len(codigos)):
            if codigos[i] == codigos[j]:
                return (i, j)
    return None


def _lista_derivacao(e, ocupantes):
    """TASK-170 (`lados: DERIVACAO`): → (ocupantes na ordem derivação, tronco…, avisos, incompleta). Tronco = o par de iguais (no máximo 2); derivação = o que sobra e é OBRIGATÓRIA.
    2 cabos iguais = só tronco (falta a derivação); 2 diferentes = ambíguo: o 1º ligado vira derivação e avisa."""
    cods = [o["info"]["codigo"] for o in ocupantes]
    if len(ocupantes) == 0:
        return [], [], False
    if len(ocupantes) == 1:
        return list(ocupantes), [], False
    if len(ocupantes) == 2:
        if cods[0] == cods[1]:
            return [], [f"{e['codigo']}: falta o cabo da derivação"], True
        return [ocupantes[0], ocupantes[1]], [f"{e['codigo']}: 2 cabos diferentes sem par ({cods[0]}, {cods[1]}); usado o 1º ligado como derivação"], False
    par = _par_do_tronco(cods)
    if par is None:
        return [], [], False                                          # tronco diferente: já é erro de vínculo (vermelho)
    resto = [i for i in range(len(ocupantes)) if i not in par]
    return [ocupantes[resto[0]]] + [ocupantes[i] for i in par], [], False


def _lista_vaos(e, ocupantes):
    """TASK-170 (`cabos_vaos`): os cabos na ORDEM DOS VÃOS (1º nome do parâmetro = 1º vão); vão sem cabo fica None."""
    atrib = _atribuicao_vaos(e["cabos_vaos"], [o["info"]["codigo"] for o in ocupantes])
    if atrib is None:
        return list(ocupantes)
    lista = [None] * len(e["cabos_vaos"])
    for o, j in zip(ocupantes, atrib):
        lista[j] = o
    while lista and lista[-1] is None:
        lista.pop()
    return lista


def _aplicar_define(e, ie, outros, ests, info_e, lista):
    """TASK-170 (`define`): parâmetros automáticos pela linha do poste (`linha_tem`) e/ou pelo 1º cabo (`cabo`). Os itens valem em ordem; o 1º que casa para cada nome vence.
    → ({NOME: VALOR}, {NOME: origem}, [{"param", "motivo"}] dos nomes que nenhum item definiu)."""
    saida, origem, faltou = {}, {}, []
    if not e["define"]:
        return saida, origem, faltou
    linha = outros[ie].get("_linha")
    outras = [info_e[j]["codigo"] for j in ests if j != ie and outros[j].get("_linha") == linha]
    primeiro = next((o["info"]["codigo"] for o in lista if o), None)
    for item in e["define"]:
        nome, valor = item["param"]
        if nome in saida:
            continue
        achou = None
        if item["linha_tem"] is not None:
            achou = next((c for c in outras if casa(item["linha_tem"], c)), None)
            if achou is None:
                continue
        if item["cabo"] is not None and (primeiro is None or not casa(item["cabo"], primeiro)):
            continue
        saida[nome] = valor
        origem[nome] = f"cabo {primeiro}" if item["cabo"] is not None else f"linha: {achou}"
    for item in e["define"]:
        nome = item["param"][0]
        if nome not in saida and nome not in [f["param"] for f in faltou]:
            lt = item["linha_tem"]
            quem = (lt if isinstance(lt, str) else ", ".join(lt)).split(",")[0].strip() if lt is not None else ""
            faltou.append({"param": nome, "motivo": f"a linha não tem {quem}" if quem else "nenhuma condição do define casou"})
    return saida, origem, faltou


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
    codigos = [o["info"]["codigo"] for o in ocupantes] + [c["codigo"]]
    if e["cabos_vaos"] and _atribuicao_vaos(e["cabos_vaos"], codigos) is None:                       # TASK-170: um cabo por vão
        return False, f"{e['codigo']}: já tem o cabo do vão em que o {c['codigo']} caberia (cada vão aceita um cabo)"
    if e["lados"] == "DERIVACAO" and len(codigos) >= 3 and _par_do_tronco(codigos) is None:        # TASK-170: com 3 cabos, 2 formam o tronco e precisam ser iguais
        return False, f"{e['codigo']}: os dois cabos do tronco precisam ser iguais ({', '.join(codigos)})"
    return True, None


# TASK-169: seção do cabo para ordenar "maior primeiro" (tronco em `cabo`, ramal em `cabo2`). Número do código (A185P → 185, S336 → 336); AWG com zero (S04, S02) e fração (S1/0) por tabela.
_AWG0 = {2: 33.6, 4: 21.2, 6: 13.3, 8: 8.4}
_AWG_X0 = {1: 53.5, 2: 67.4, 3: 85.0, 4: 107.2}


def secao_cabo(codigo, dic=None) -> float:
    c = _trim(codigo).upper()
    if dic:                                                                                               # TASK-173: dicionário primeiro; sem linha, a regra fixa
        s = _dc.secao(dic, c)
        if s is not None:
            return s
    m = re.search(r"(\d)/0", c)
    if m:
        return _AWG_X0.get(int(m.group(1)), float(m.group(1)))
    m = re.fullmatch(r"[A-Z]+0(\d)", c)
    if m:
        return _AWG0.get(int(m.group(1)), float(m.group(1)))
    m = re.search(r"(\d+)", c)
    return float(m.group(1)) if m else 0.0


def _ordenar_e_juntar(e, ocupantes) -> list:
    """TASK-169: `ordem: MAIOR_PRIMEIRO` (estável: empate = ordem de origem) e `distintos` (junta os cabos de mesma descrição) ANTES de distribuir nos parâmetros."""
    lista = list(ocupantes)
    if e["ordem"] == "MAIOR_PRIMEIRO":
        lista = sorted(lista, key=lambda o: -(o["info"]["secao"] if o["info"].get("secao") is not None else secao_cabo(o["info"]["codigo"])))
    if e["distintos"]:
        vistos, saida = set(), []
        for o in lista:
            k = o["info"]["descricao"]
            if k in vistos:
                continue
            vistos.add(k)
            saida.append(o)
        lista = saida
    return lista


def _op_aceita(c) -> bool:
    """Mesma regra de operação de `pode_vincular` (a regra do cabo lista as operações que aceita)."""
    if isinstance(c["operacoes"], list) and c["operacoes"]:
        aceitas = [_trim(x).upper().replace("*", "") for x in c["operacoes"]]
        opn = c["op"].replace("*", "")
        return opn in aceitas or (opn == "R" and "I" in aceitas)
    return True


def _nome_poste(row) -> str:
    return _trim(str((row or {}).get("ativo") or "").split(",")[0])


def _mapa_postes(outros, ests) -> list:
    """[(linha, (x, y))] dos postes (`_poste` da linha de Outros) — o mesmo da vinculação por poste (TASK-166)."""
    por_linha = {}
    for ie in ests:
        verificar_prazo()
        por_linha.setdefault(outros[ie].get("_linha"), []).append(ie)
    postes = [(linha, _par(outros[ies[0]].get("_poste"))) for linha, ies in sorted(por_linha.items(), key=lambda x: (x[0] is None, x[0]))]
    return [(linha, pos) for linha, pos in postes if pos]


# ── análise completa (estados, vínculos inválidos, valor do parâmetro) ─────────────────────────────
def _codigo_cabo(row, prefixo_fn):
    return prefixo_fn(row.get("ativo")) or ""


# TASK-095: só entram no vínculo linhas de operação I, M e as variantes *I e *M. R e *R (retirada/recolocação) ficam de fora — nem listadas, nem contadas.
def participa_vinculo(row) -> bool:
    return _trim((row or {}).get("operacao")).upper() not in ("R", "*R")


def estruturas_ficticias_de(row) -> list:
    """Estruturas fictícias do CABO (só o código): preenchem pontas do cabo, sem linha em Outros; entram na Totalizadora como `CABO_ESTRUTURA`."""
    return [_trim(x).upper() for x in ((row or {}).get("estruturasFicticias") or []) if _trim(x)]


def _eh_cabo_valido(row):
    return bool(row) and bool(_trim(row.get("ativo"))) and participa_vinculo(row)


def _eh_linha_estrutura(r):
    # TASK-125: também a linha ainda não classificada ('0' ou vazia) que tem ativo (mesma regra de static/vinculacao_regras.js)
    # Linha de Outros CONGELADA não é estrutura (fora do vínculo); o cabo congelado continua valendo como fonte de um vínculo manual (só não é ligado sozinho pelos algoritmos).
    return bool(r) and not r.get("congelada") and participa_vinculo(r) and (r.get("entidade") in ("ESTRUTURA", "POSTE") or (r.get("entidade") in (None, "", "0") and bool(_trim(r.get("ativo")))))


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


def _valor_e_parametros(e, cfg, ocupantes):
    """Valor consolidado do cabo (parâmetro único) e os parâmetros por nome da estrutura, a partir dos ocupantes (já ordenados/juntados). → (valor, avisos, parametros)."""
    avisos = []
    valores = []
    reais = [oc for oc in ocupantes if oc]                       # TASK-170: vão sem cabo = None (a posição do parâmetro é mantida)
    for oc in reais:
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
    parametros = None
    if e["parametros"]:                                          # TASK-096: um valor por cabo vinculado (reais e fictícios), na ordem dos ocupantes
        cont = {}
        vezes = {}
        for p in e["parametros"]:
            vezes[p["nome"]] = vezes.get(p["nome"], 0) + 1
        for p in e["parametros"]:
            m_suf = re.fullmatch(r"(?:CABO|TIPOCABO)(\d+)", p["nome"])
            k = int(m_suf.group(1)) - 1 if m_suf else cont.get(p["nome"], 0)           # cabo2 = o 2º cabo; cabo, cabo = 1º e 2º
            cont[p["nome"]] = cont.get(p["nome"], 0) + 1
            o = ocupantes[k] if 0 <= k < len(ocupantes) else None
            # `tipocabo` escrito UMA vez numa estrutura com 2 cabos = o tipo consolidado (se_diferem: MAIOR C>P>A…), como no parâmetro único
            v = (o["info"]["valor_cabo"] if p["fonte"] == "valor_cabo" else (o["info"]["descricao"] if p["fonte"] == "descricao" else o["info"]["codigo"])) if o else None
            if p["fonte"] == "valor_cabo" and vezes[p["nome"]] == 1 and len(reais) > 1:
                v = valor
            if v:
                parametros = parametros or {}
                parametros.setdefault(p["nome"], []).append(v)
    return valor, avisos, parametros


def _lista_de_parametros(e, ocupantes):
    """A ordem em que os cabos vão para os parâmetros: derivação+tronco (`DERIVACAO`), um por vão (`cabos_vaos`) ou ordem/distintos. → (lista, avisos_vinculo, incompleta)."""
    if e["lados"] == "DERIVACAO":
        return _lista_derivacao(e, ocupantes)
    if e["cabos_vaos"]:
        return _lista_vaos(e, ocupantes), [], False
    return _ordenar_e_juntar(e, ocupantes), [], False


def _mesclar_define(e, parametros, valor, saida):
    """TASK-170: junta os parâmetros do `define` aos do vínculo SEM sobrescrever o que o vínculo já deu (nem o digitado: `_com_param` não duplica)."""
    if not saida:
        return parametros
    novo = dict(parametros or {})
    if parametros is None and valor and e["parametro"]:
        novo[e["parametro"]] = valor                              # o parâmetro único (valor do cabo) continua valendo ao lado do define
    for nome, v in saida.items():
        if nome not in novo:
            novo[nome] = [v]
    return novo


def _analisar_base(cabos, outros, cfg, prefixo_fn, codigo_fn) -> dict:
    cfg = normalizar_config(cfg)
    ests = estruturas_disponiveis(outros)
    info_c = {}
    for ic, c in enumerate(cabos):
        if _eh_cabo_valido(c):
            info_c[ic] = info_cabo(_codigo_cabo(c, prefixo_fn), c.get("operacao"), cfg, extra_cabo(c, cfg.get("dicionario_cabos")))
    info_e = {ie: info_estrutura(codigo_fn(outros[ie].get("ativo"), outros[ie]), cfg) for ie in ests}
    ests_ocupa = [ie for ie in ests if info_e[ie]["ocupa"]]            # TASK-169: equipamento (ocupa: false) não consome ponta; é resolvido na 2ª passada
    ocup = {ie: [] for ie in ests_ocupa}                               # [(idx_cabo)]
    for ic, c in enumerate(cabos):
        if ic not in info_c:
            continue
        for ie in (c.get("vinculoEstruturas") or []):
            if ie in ocup:
                ocup[ie].append(ic)
    links, bad_links = [], set()
    res_e = {}
    for ie in ests_ocupa:
        verificar_prazo()
        e = info_e[ie]
        fict = [_trim(x).upper() for x in (outros[ie].get("cabosFicticios") or []) if _trim(x)]
        tem = len(ocup[ie]) + len(fict)
        erros = []
        ocupantes = []
        for ic in ocup[ie]:
            ok, motivo = pode_com_ocupantes(info_c[ic], e, ocupantes)
            ocupantes.append({"info": info_c[ic], "ficticio": False})
            links.append({"cabo": ic, "estrutura": ie, "ok": ok, "motivo": motivo})
            if not ok:
                bad_links.add((ic, ie))
                erros.append(motivo)
        for cod in fict:
            ic_f = info_cabo_ficticio(cod, e["codigo"], op_ficticio(e), cfg)
            ok, motivo = pode_com_ocupantes(ic_f, e, ocupantes, True)
            ocupantes.append({"info": ic_f, "ficticio": True})
            if not ok:
                erros.append(motivo)
        lista, avisos_vinculo, incompleta = _lista_de_parametros(e, ocupantes)
        valor, avisos, parametros = _valor_e_parametros(e, cfg, lista)
        d_saida, d_origem, d_faltou = _aplicar_define(e, ie, outros, ests, info_e, lista)
        parametros = _mesclar_define(e, parametros, valor, d_saida)
        if erros:
            estado = "rd"
        elif tem == 0:
            estado = "k"
        elif (e["qtd_min"] is not None and tem < e["qtd_min"]) or incompleta:
            estado = "pu"
        else:
            estado = "gr"
        res_e[ie] = {"estado": estado, "tem": tem, "precisa": e["qtd_min"], "maximo": e["qtd"], "valor_cabo": valor, "erros": erros, "avisos": avisos, "codigo": e["codigo"],
                     "regra": e["regra"], "ficticios": len(fict), "parametro": e["parametro"], "parametros": parametros, "ocupa": True,
                     "avisos_vinculo": avisos_vinculo, "define": {n: {"valor": v, "origem": d_origem[n]} for n, v in d_saida.items()}, "define_faltou": d_faltou}
    # TASK-169 — 2ª passada: equipamento (RL, CFU, para-raio…) lê os cabos do POSTE, sem ocupar pontas
    postes = _mapa_postes(outros, ests)
    for ie in ests:
        verificar_prazo()
        e = info_e[ie]
        if e["ocupa"]:
            continue
        linha = outros[ie].get("_linha")
        unidades = [j for j in ests_ocupa if outros[j].get("_linha") == linha]
        candidatos = []                                                    # (ic | None, ocupante): ordem de origem = estruturas da linha, depois os vínculos; cabos só da geometria; fictícios
        vistos = set()
        for j in unidades:
            for ic in ocup[j]:
                if ic not in vistos:
                    vistos.add(ic)
                    candidatos.append({"info": info_c[ic], "ficticio": False})
        pos = _par(outros[ie].get("_poste"))
        if pos:
            for ic in sorted(info_c):
                verificar_prazo()
                ext = _extremos(cabos[ic])
                if ic not in vistos and ext and any(_poste_mais_perto(pt, postes, RAIO_POSTE_PADRAO) == linha for pt in ext):
                    vistos.add(ic)
                    candidatos.append({"info": info_c[ic], "ficticio": False})
        for j in unidades:
            for x in (outros[j].get("cabosFicticios") or []):
                if _trim(x):
                    candidatos.append({"info": info_cabo_ficticio(_trim(x).upper(), info_e[j]["codigo"], op_ficticio(info_e[j]), cfg), "ficticio": True})
        lista = []
        for o in candidatos:
            c = o["info"]
            if not c["vincula"] or not _op_aceita(c) or not casa(e["cabos"], c["codigo"]):
                continue
            if e["nivel"] and c["nivel"] is not None and c["nivel"] != e["nivel"]:
                continue
            lista.append(o)
        lista, avisos_vinculo, incompleta = _lista_de_parametros(e, lista)
        erros, avisos_f, valor, parametros = [], [], None, None
        if e["herda"] == "POSTE":
            valor, avisos, parametros = _valor_e_parametros(e, cfg, lista)
            avisos_f += avisos
        d_saida, d_origem, d_faltou = _aplicar_define(e, ie, outros, ests, info_e, lista)
        parametros = _mesclar_define(e, parametros, valor, d_saida)
        if e["herda"] == "POSTE" and not lista:
            msg = f"{e['codigo']}: o poste {_nome_poste(outros[ie])} não tem cabo{' ' + e['nivel'] if e['nivel'] else ''} ligado"
            if e["se_faltar"] == "BLOQUEAR":
                erros.append(msg)
            else:
                avisos_f.append(msg)
        estado = "rd" if erros else ("gr" if (lista or e["herda"] != "POSTE") else "k")          # nunca "pu" (TASK-169)
        res_e[ie] = {"estado": estado, "tem": len([o for o in lista if o]), "precisa": None, "maximo": None, "valor_cabo": valor, "erros": erros, "avisos": avisos_f, "codigo": e["codigo"],
                     "regra": e["regra"], "ficticios": 0, "parametro": e["parametro"], "parametros": parametros, "ocupa": False,
                     "le_do_poste": {"poste": _nome_poste(outros[ie]), "cabos": [o["info"]["codigo"] for o in lista if o]},
                     "avisos_vinculo": avisos_vinculo, "define": {n: {"valor": v, "origem": d_origem[n]} for n, v in d_saida.items()}, "define_faltou": d_faltou}
    res_c = {}
    for ic, c in info_c.items():
        mine = [ie for ie in (cabos[ic].get("vinculoEstruturas") or []) if ie in ocup]
        fict_e = estruturas_ficticias_de(cabos[ic]) if c["vincula"] else []
        erros_f = [m for ok, m in (pode_vincular(c, info_estrutura(cod, cfg)) for cod in fict_e) if not ok]
        invalido = any((ic, ie) in bad_links for ie in mine) or bool(erros_f)
        precisa = c["pontas"] if c["vincula"] else 0
        tem = len(mine) + len(fict_e)
        if invalido:
            estado = "rd"
        elif not c["vincula"]:
            estado = "gr"
        elif not tem:
            estado = "k"
        elif tem < precisa:
            estado = "pu"
        else:
            estado = "gr"
        res_c[ic] = {"estado": estado, "tem": tem, "precisa": precisa, "codigo": c["codigo"], "valor_cabo": c["valor_cabo"], "vincula": c["vincula"],
                     "ficticias": len(fict_e), "erros": erros_f}
    return {"estruturas": res_e, "cabos": res_c, "links": links}


def parametros_do_cabo_na_estrutura(estrutura, cabo_row, cfg, prefixo_fn) -> dict:
    """TASK-171: {NOME: [valores]} que a estrutura `estrutura` receberia se ESTE cabo fosse o único ligado a ela (mesmas fontes da TASK-100: descrição, tipocabo, código).
    Sem regra de cabo/estrutura devolve o que a tabela der (ou {})."""
    cfg = normalizar_config(cfg)
    e = info_estrutura(estrutura, cfg)
    c = info_cabo(_codigo_cabo(cabo_row, prefixo_fn), (cabo_row or {}).get("operacao"), cfg, extra_cabo(cabo_row, cfg.get("dicionario_cabos")))
    valor, _avisos, params = _valor_e_parametros(e, cfg, [{"info": c, "ficticio": False}])
    if params is None and valor and e["parametro"]:
        params = {e["parametro"]: [valor]}
    return params or {}


def parametros_por_linha(cabos, outros, cfg, prefixo_fn, codigo_fn, codigos_fn=None) -> dict:
    """{indice_em_outros: {"CABO": "C"}} — o parâmetro que a estrutura recebe do(s) cabo(s) vinculado(s) (só onde há valor)."""
    a = analisar(cabos, outros, cfg, prefixo_fn, codigo_fn, codigos_fn)
    saida = {}
    for ie, r in sorted(a["estruturas"].items()):
        linha = linha_de(ie)
        if r["valor_cabo"] and r["parametro"] and linha not in saida:       # várias estruturas na linha: vale a 1ª que tem valor
            saida[linha] = {r["parametro"]: r["valor_cabo"]}
        elif r["parametros"] and linha not in saida:
            saida[linha] = {k: list(v) for k, v in r["parametros"].items()}
    return saida


def parametros_por_estrutura(cabos, outros, cfg, prefixo_fn, codigo_fn, codigos_fn=None) -> dict:
    """TASK-096: {linha: {CODIGO_DA_ESTRUTURA: {NOME: valor | [valores]}}} — cada estrutura da linha recebe SÓ os parâmetros dela (N1 e N4 do mesmo poste não se misturam)."""
    a = analisar(cabos, outros, cfg, prefixo_fn, codigo_fn, codigos_fn)
    saida = {}
    for ie, r in sorted(a["estruturas"].items()):
        params = r["parametros"] or ({r["parametro"]: r["valor_cabo"]} if r["valor_cabo"] and r["parametro"] else None)
        if not params:
            continue
        por_codigo = saida.setdefault(linha_de(ie), {})
        por_codigo.setdefault(r["codigo"], params)
    return saida


def achados(cabos, outros, cfg, prefixo_fn, codigo_fn, codigos_fn=None) -> list:
    """Avisos de validação (formato de services/validacao_planilhas.py), severidade "aviso"."""
    cfg = normalizar_config(cfg)
    a = analisar(cabos, outros, cfg, prefixo_fn, codigo_fn, codigos_fn)
    existe_m = any(_eh_cabo_valido(c) and not c.get("congelada") and op_base(c.get("operacao")) == "M" for c in cabos)
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
        if not r.get("ocupa", True):                                           # TASK-169: equipamento lê do poste — só avisa quando falta cabo (ou bloqueia)
            for msg in r["erros"]:
                add(ie, f"Estrutura \"{nome}\": {msg}.", "O equipamento lê os cabos do poste e a regra manda bloquear sem cabo (se_faltar = BLOQUEAR).")
            for msg in r["avisos"]:
                add(ie, f"Estrutura \"{nome}\": {msg}.", "O equipamento lê os cabos do poste; sem cabo o parâmetro fica no padrão da base.")
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
        for msg in r.get("avisos_vinculo", []):                                  # TASK-170: derivação incompleta / ambígua
            add(ie, f"Estrutura \"{nome}\": {msg}.", "A regra da estrutura (lados DERIVACAO) pede 1 cabo de derivação e até 2 do tronco, iguais entre si.")
    return saida


# ── vinculação POR POSTE (TASK-166) ─────────────────────────────────────────────────────────────────
# O vão (cabo) tem 2 pontas e cada ponta termina em UM poste (`_geo`: 1º e último ponto; o poste está em `_poste` na linha de Outros). Por isso o problema se separa por poste:
# cada estrutura do poste procura, entre os cabos que CHEGAM nele, os que cumprem as validações dela (cabos aceitos, operação, lados iguais, quantidade) e consome UMA ponta de cada.
# B1MT (2 cabos MT iguais): se o poste tem 2 cabos MT iguais chegando, a B1MT liga nos dois e cada cabo ganha a ponta daquele poste.
# Prioridade entre estruturas que disputam o mesmo cabo: a ordem da Tabela de Estruturas (primeira regra), depois a ordem da estrutura na linha do poste.
RAIO_POSTE_PADRAO = 5.0


def _par(p):
    try:
        if isinstance(p, (list, tuple)) and len(p) >= 2 and not isinstance(p[0], bool):
            return float(p[0]), float(p[1])
    except (TypeError, ValueError):
        pass
    return None


def postes_atribuidos(cabo) -> list:
    """Linhas de Outros (postes) a que o usuário atribuiu o cabo (aba Cabos ⇄ Postes); sem repetição, só inteiros >= 0."""
    v = (cabo or {}).get("vinculoPostes")
    r = []
    if isinstance(v, list):
        for x in v:
            if isinstance(x, int) and not isinstance(x, bool) and x >= 0 and x not in r:
                r.append(x)
    return r


def _extremos(cabo):
    g = cabo.get("_geo")
    if not isinstance(g, (list, tuple)) or len(g) < 2:
        return None
    a, b = _par(g[0]), _par(g[-1])
    return (a, b) if a and b else None


def _poste_mais_perto(pt, postes, raio):
    melhor, dmin = None, None
    for linha, pos in postes:
        verificar_prazo()
        d = ((pt[0] - pos[0]) ** 2 + (pt[1] - pos[1]) ** 2) ** 0.5
        if d <= raio and (dmin is None or d < dmin):
            melhor, dmin = linha, d
    return melhor


def vincular_por_poste(cabos, outros, cfg, prefixo_fn, codigo_fn, codigos_fn=None, raio=RAIO_POSTE_PADRAO, aplicar=False) -> dict:
    """Vincula cabos às estruturas dos postes pelo que CHEGA em cada poste. Vínculo marcado `vinculoOrigem == "manual"` é mantido (e ocupa a estrutura); os demais são refeitos.
    Cabo sem `_geo` não é tocado. `aplicar=True` grava `vinculoEstruturas` (+ `vinculoOrigem = "poste"`) nos cabos. Devolve {"vinculos", "achados", "resumo"}."""
    virt, id_de, cabos_v, codigo_v = _expandir(cabos, outros, _com_linha(codigo_fn), codigos_fn)
    r = _vincular_por_poste_base(cabos_v, virt, cfg, prefixo_fn, codigo_v, raio if isinstance(raio, (int, float)) and raio > 0 else RAIO_POSTE_PADRAO)
    for v in r["vinculos"]:
        v["estruturas"] = [id_de[j] for j in v["estruturas"]]
    for a in r["achados"]:
        if a.get("estrutura") is not None:
            a["estrutura"] = id_de[a["estrutura"]]
    if aplicar:
        for v in r["vinculos"]:
            cabos[v["cabo"]]["vinculoEstruturas"] = list(v["estruturas"])
            if v["estruturas"]:
                cabos[v["cabo"]]["vinculoOrigem"] = "poste"
            else:
                cabos[v["cabo"]].pop("vinculoOrigem", None)
    return r


def _vincular_por_poste_base(cabos, outros, cfg, prefixo_fn, codigo_fn, raio) -> dict:
    cfg = normalizar_config(cfg)
    ests = estruturas_disponiveis(outros)
    info_e = {ie: info_estrutura(codigo_fn(outros[ie].get("ativo"), outros[ie]), cfg) for ie in ests}
    info_c = {ic: info_cabo(_codigo_cabo(c, prefixo_fn), c.get("operacao"), cfg, extra_cabo(c, cfg.get("dicionario_cabos"))) for ic, c in enumerate(cabos) if _eh_cabo_valido(c)}
    por_linha = {}
    for ie in ests:
        verificar_prazo()
        por_linha.setdefault(outros[ie].get("_linha"), []).append(ie)
    todos_postes = [(linha, _par(outros[ies[0]].get("_poste"))) for linha, ies in sorted(por_linha.items(), key=lambda x: (x[0] is None, x[0]))]
    postes = [(linha, pos) for linha, pos in todos_postes if pos]          # todos_postes: também os atribuídos à mão (sem geometria)
    fict = {ie: [_trim(x).upper() for x in (outros[ie].get("cabosFicticios") or []) if _trim(x)] for ie in ests}
    manuais = {ic for ic in info_c if cabos[ic].get("vinculoOrigem") == "manual" and (cabos[ic].get("vinculoEstruturas") or [])}
    ocup = {ie: [] for ie in ests}                               # cabos ligados à estrutura (manuais primeiro; depois os que este algoritmo liga)
    for ic in sorted(manuais):
        for ie in cabos[ic]["vinculoEstruturas"]:
            if ie in ocup:
                ocup[ie].append(ic)
    ligados, restante, chegadas, soltas, sem_geo = {}, {}, {}, [], []
    saida = {"vinculos": [], "achados": [], "resumo": {}}
    for ic in sorted(info_c):
        c = info_c[ic]
        if not c["vincula"] or ic in manuais or cabos[ic].get("congelada"):          # cabo congelado: só entra no vínculo se já foi ligado à mão
            continue
        atrib = [l for l in postes_atribuidos(cabos[ic]) if l in por_linha]          # "Cabos ⇄ Postes": postes dados pelo usuário valem no lugar da geometria
        if atrib or cabos[ic].get("_mapaTopologia") is True:
            ligados[ic] = []
            restante[ic] = max(0, c["pontas"] - len(estruturas_ficticias_de(cabos[ic])))
            for linha in atrib:
                chegadas.setdefault(linha, []).append(ic)
            continue
        fim = _extremos(cabos[ic])
        if fim is None:
            sem_geo.append(ic)
            continue
        ligados[ic] = []
        restante[ic] = max(0, c["pontas"] - len(estruturas_ficticias_de(cabos[ic])))
        for k, pt in enumerate(fim):
            linha = _poste_mais_perto(pt, postes, raio)
            if linha is None:
                soltas.append((ic, k))
            else:
                chegadas.setdefault(linha, []).append(ic)

    def nome_poste(linha):
        return f"poste da linha {linha + 1} ({outros[por_linha[linha][0]].get('ativo')})"

    def add(tipo, msg, linha=None, cabo=None, estrutura=None):
        saida["achados"].append({"tipo": tipo, "poste": linha, "cabo": cabo, "estrutura": estrutura, "mensagem": msg})
    estr_incompletas = 0
    for linha, _pos in todos_postes:
        pend = list(chegadas.get(linha, []))
        unidades = sorted(por_linha[linha], key=lambda ie: ((info_e[ie]["regra"] if info_e[ie]["regra"] is not None else 10 ** 6), por_linha[linha].index(ie)))
        for ie in unidades:
            e = info_e[ie]
            if e["regra"] is None or e["qtd"] is None or not e["ocupa"]:
                continue                                          # sem regra/quantidade na Tabela de Estruturas (ou equipamento que não ocupa — TASK-169): não exige cabo
            ocupantes = [{"info": info_c[x], "ficticio": False} for x in ocup[ie]] + [{"info": info_cabo_ficticio(cod, e["codigo"], op_ficticio(e), cfg), "ficticio": True} for cod in fict[ie]]
            falta = e["qtd"] - len(ocupantes)
            if falta > 0:
                cands = []
                for ic in pend:
                    if restante.get(ic, 0) > 0 and pode_vincular(info_c[ic], e)[0] and ic not in cands:
                        cands.append(ic)
                if e["lados"] == "IGUAIS" and cands and not ocupantes:
                    grupos = {}
                    for ic in cands:
                        grupos.setdefault(_chave_lado(info_c[ic]), []).append(ic)
                    cands = max(grupos.values(), key=lambda g: min(len(g), falta))              # o grupo de cabos iguais que mais enche a estrutura (empate: o 1º)
                escolhidos = []
                for ic in cands:
                    if len(escolhidos) >= falta:
                        break
                    if pode_com_ocupantes(info_c[ic], e, ocupantes)[0]:
                        ocupantes.append({"info": info_c[ic], "ficticio": False})
                        escolhidos.append(ic)
                if len(cands) > len(escolhidos) and escolhidos:
                    add("escolha", f"{e['codigo']} no {nome_poste(linha)}: {len(cands)} cabos servem, ligados os {len(escolhidos)} primeiros (ordem da tabela de cabos).", linha, escolhidos[0], ie)
                for ic in escolhidos:
                    ocup[ie].append(ic)
                    ligados[ic].append(ie)
                    restante[ic] -= 1
                    pend.remove(ic)
            tem = len(ocup[ie]) + len(fict[ie])
            if e["qtd_min"] is not None and tem < e["qtd_min"]:
                estr_incompletas += 1
                add("incompleta", f"{e['codigo']} no {nome_poste(linha)}: precisa de {e['qtd_min']} cabo(s) e só chegaram {tem} que ela aceita.", linha, None, ie)
        for ic in pend:
            if restante.get(ic, 0) > 0:
                add("sobra", f"Cabo {info_c[ic]['codigo']} (linha {ic + 1}) chega ao {nome_poste(linha)} e nenhuma estrutura o usou.", linha, ic, None)
    for ic, k in soltas:
        add("ponta_solta", f"Cabo {info_c[ic]['codigo']} (linha {ic + 1}): a ponta {'A' if k == 0 else 'B'} não toca nenhum poste (raio {raio:g}).", None, ic, None)
    for ic in sem_geo:
        add("sem_geometria", f"Cabo {info_c[ic]['codigo']} (linha {ic + 1}) sem geometria (_geo): vínculo não alterado.", None, ic, None)
    completos = 0
    for ic in sorted(ligados):
        full = len(ligados[ic]) >= max(0, info_c[ic]["pontas"] - len(estruturas_ficticias_de(cabos[ic])))
        completos += 1 if full and ligados[ic] else 0
        saida["vinculos"].append({"cabo": ic, "estruturas": list(ligados[ic]), "completo": bool(full and ligados[ic])})
    saida["resumo"] = {"cabos": len(ligados), "completos": completos, "parciais": sum(1 for v in saida["vinculos"] if v["estruturas"] and not v["completo"]),
                       "sem_vinculo": sum(1 for v in saida["vinculos"] if not v["estruturas"]), "estruturas_incompletas": estr_incompletas,
                       "pontas_soltas": len(soltas), "manuais_mantidos": len(manuais), "sem_geometria": len(sem_geo)}
    return saida


# ── script de vinculação automática (TASK-084, entrega 2) ───────────────────────────────────────────
MODOS_SCRIPT = ("proximidade", "sequencia", "aleatorio")
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
        if p.get("semente") is not None and (not isinstance(p["semente"], int) or isinstance(p["semente"], bool)):
            erros.append(f"{pre}: 'semente' precisa ser um inteiro.")
        for k in ("completar", "puro", "aceitar_parcial", "substituir"):
            if p.get(k) is not None and not isinstance(p[k], bool):
                erros.append(f"{pre}: '{k}' precisa ser true ou false.")
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


# TASK-099: passo "aleatorio" (reserva do vínculo por coordenadas). Sorteio REPRODUZÍVEL: mesma semente + mesmos dados = mesmo resultado (mulberry32, igual ao JS).
def _mulberry32(a):
    estado = [a & 0xFFFFFFFF]

    def rnd():
        estado[0] = (estado[0] + 0x6D2B79F5) & 0xFFFFFFFF
        a_ = estado[0]
        t = ((a_ ^ (a_ >> 15)) * (1 | a_)) & 0xFFFFFFFF
        t = ((t + (((t ^ (t >> 7)) * (61 | t)) & 0xFFFFFFFF)) & 0xFFFFFFFF) ^ t
        return ((t ^ (t >> 14)) & 0xFFFFFFFF) / 4294967296
    return rnd


def _embaralhar(lista, rnd):
    l = list(lista)
    for i in range(len(l) - 1, 0, -1):
        j = int(rnd() * (i + 1))
        l[i], l[j] = l[j], l[i]
    return l


def _executar_script_base(script, cabos, outros, cfg, prefixo_fn, codigo_fn) -> dict:
    cfg = normalizar_config(cfg)
    modo_conf = script.get("conflitos") if script.get("conflitos") in CONFLITOS_SCRIPT else "avisar"
    ests = estruturas_disponiveis(outros)
    info_e = {ie: info_estrutura(codigo_fn(outros[ie].get("ativo"), outros[ie]), cfg) for ie in ests}
    info_c = {ic: info_cabo(_codigo_cabo(c, prefixo_fn), c.get("operacao"), cfg, extra_cabo(c, cfg.get("dicionario_cabos"))) for ic, c in enumerate(cabos) if _eh_cabo_valido(c)}
    ligados = {ic: list(cabos[ic].get("vinculoEstruturas") or []) for ic in info_c}       # estado de trabalho
    fict = {ie: [_trim(x).upper() for x in (outros[ie].get("cabosFicticios") or []) if _trim(x)] for ie in ests}
    saida = {"vinculos": [], "conflitos": [], "sem_candidato": [], "passos": [], "parou": False}

    def ocupantes_de(ie, ignorando=None):
        lista = [{"info": info_c[ic], "ficticio": False} for ic in sorted(info_c) if ic != ignorando and ie in ligados[ic]]
        lista += [{"info": info_cabo_ficticio(cod, info_e[ie]["codigo"], op_ficticio(info_e[ie]), cfg), "ficticio": True} for cod in fict[ie]]
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
            if (ligados[ic] or estruturas_ficticias_de(cabos[ic])) and not passo.get("substituir"):
                # `completar`: o cabo que ficou com menos pontas do que precisa (ex.: parcial de um passo anterior) volta a ser alvo e só ganha o que falta
                falta_ = (passo["pontas"] if isinstance(passo.get("pontas"), int) else c["pontas"]) - len(estruturas_ficticias_de(cabos[ic])) - len(ligados[ic])
                if not (passo.get("completar") and falta_ > 0):
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
        cand_global = [ie for ie in ests if estrutura_ok(ie) and info_e[ie]["ocupa"]]          # TASK-169: equipamento (ocupa: false) não recebe cabo
        raio = passo.get("raio_m")
        for pos, ic in enumerate(alvo):
            c = info_c[ic]
            need = max(0, (passo.get("pontas") if isinstance(passo.get("pontas"), int) else c["pontas"]) - (0 if passo.get("substituir") else len(estruturas_ficticias_de(cabos[ic]))))
            antes = list(ligados[ic])
            if passo.get("substituir"):
                ligados[ic] = []
            base = list(ligados[ic])
            falta = max(0, need - len(base))                  # só ganha o que falta (base vazia, exceto em `completar`)
            escolhidas, recusas = [], []
            if passo.get("modo") == "sequencia":
                candidatas = cand_global[pos:pos + need]
            elif passo.get("modo") == "aleatorio":
                # sorteio entre as estruturas aceitas; por padrão as com MENOS cabos primeiro (sorteio entre as vazias antes das ocupadas); `puro: true` = sorteio simples
                semente = passo["semente"] if isinstance(passo.get("semente"), int) and not isinstance(passo.get("semente"), bool) else 1
                rnd = _mulberry32(semente * 1000003 + ic * 7919 + ip * 104729)
                candidatas = _embaralhar([ie for ie in cand_global if ie not in base], rnd)
                if not passo.get("puro"):
                    candidatas = [x[0] for x in sorted(((ie, k, len(ocupantes_de(ie))) for k, ie in enumerate(candidatas)), key=lambda x: (x[2], x[1]))]
            else:
                if not _tem_xy(cabos[ic]):
                    r["sem_candidato"] += 1
                    saida["sem_candidato"].append({"passo": ip, "cabo": ic, "motivo": "cabo sem coordenada"})
                    ligados[ic] = antes
                    continue
                candidatas = sorted((ie for ie in cand_global if _tem_xy(outros[ie]) and (raio is None or _dist(cabos[ic], outros[ie]) <= raio)),
                                    key=lambda ie: (_dist(cabos[ic], outros[ie]), ie))
            for ie in candidatas:
                if len(escolhidas) >= falta:
                    break
                if ie in base:
                    continue
                ok, motivo = pode(ic, ie)
                if ok and outros[ie].get("_linha") is not None and any(outros[e2].get("_linha") == outros[ie].get("_linha") for e2 in base + escolhidas):
                    continue                                   # TASK-094: 2 estruturas do mesmo poste não valem como as 2 pontas
                if ok:
                    escolhidas.append(ie)
                    ligados[ic].append(ie)
                else:
                    recusas.append(motivo)
            completo = need > 0 and len(base) + len(escolhidas) == need
            if completo or (escolhidas and passo.get("aceitar_parcial")):
                saida["vinculos"].append({"passo": ip, "cabo": ic, "estruturas": base + escolhidas, "completo": completo})
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
