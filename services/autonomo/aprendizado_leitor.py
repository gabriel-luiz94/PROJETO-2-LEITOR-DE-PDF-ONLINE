"""
services/autonomo/aprendizado_leitor.py — Aprendizado de REGRAS DO LEITOR a partir das correções (TASK-059, etapa 2). Funções puras (o motor é injetado).

Conjunto de aprendizado = itens do desenho (texto, cor, camada) com o que o motor entregou (entidade, operação, ativo) e o que o USUÁRIO deixou:
  • Leitor (manual): o que o motor preencheu × o que está nas colunas depois das suas edições, no botão "Processar";
  • autônomo: o item reencontrado pela coordenada nas tabelas que o autônomo entregou × nas que você salvou.
Os itens que o usuário NÃO corrigiu também entram: servem de prova ("a regra proposta não estraga o que já estava certo").

Uma regra só é proposta quando, SIMULADA com o MESMO motor da tela (QuickJS) junto das regras atuais do projeto:
  (1) corrige pelo menos 80% das ocorrências do grupo (e ≥ MINIMO_PESO ocorrências, em ≥ MINIMO_FONTES arquivos/sessões), e
  (2) NÃO estraga nenhum campo (entidade, operação, ativo) que o motor já entregava como o usuário queria, em nenhum item observado (zero regressões).
Nada é aplicado sem aprovação. Os textos do desenho vão só para o armazenamento local (SQLite) do usuário.
"""
import json
import re

MINIMO_PESO = 3
MINIMO_FONTES = 2
COBERTURA_MINIMA = 0.8
MAX_ALTERNATIVAS = 6
MAX_GRUPOS = 25
MAX_PROPOSTAS = 20
LIMITE_ITENS = 6000
_JS_ESPECIAIS = re.compile(r"([\\^$.*+?()\[\]{}|/])")


def _esc(s: str) -> str:
    """Escapa só o que é especial em regex de JavaScript (o motor roda em JS)."""
    return _JS_ESPECIAIS.sub(r"\\\1", s)


def tri(e, o, a) -> tuple:
    """(entidade, operação, ativo) como aparece numa linha das tabelas (vazios viram '0', 'M' e '')."""
    return ((e or "0").strip() or "0", (o or "M").strip() or "M", (a or "").strip())


def _valor_json(o):
    return None if o in ("", None) else o


def montar_conjunto_autonomo(origens: list, agente: dict, usuario: dict) -> list:
    """Itens de uma execução do autônomo com o que o usuário deixou. `origens` = lista de {texto,cor,layer,x,y,e,o,a} (montagem);
    `agente`/`usuario` = {"cabos": [...], "outros": [...]} (linhas com `_x`/`_y`). Item cuja linha foi alterada por ajustes do autônomo (a linha
    final já não é o que o motor entregou) é ignorado: o que mudou ali não é culpa do leitor. Linha que sumiu no usuário = entidade '0'."""
    def indexar(tabelas):
        idx = {}
        for t in ("cabos", "outros"):
            dados = (tabelas or {}).get(t)
            dados = dados.get("data") if isinstance(dados, dict) else dados
            for r in dados or []:
                if isinstance(r, dict) and r.get("_x") is not None and r.get("_y") is not None:
                    idx.setdefault((round(float(r["_x"]), 4), round(float(r["_y"]), 4)), r)
        return idx
    ia, iu = indexar(agente), indexar(usuario)
    saida = []
    for o in origens or []:
        try:
            chave = (round(float(o["x"]), 4), round(float(o["y"]), 4))
        except (KeyError, TypeError, ValueError):
            continue
        ag = ia.get(chave)
        eng = tri(o.get("e"), o.get("o"), o.get("a"))
        if ag is None or tri(ag.get("entidade"), ag.get("operacao"), ag.get("ativo")) != eng:
            continue
        us = iu.get(chave)
        user = ("0", eng[1], eng[2]) if us is None else tri(us.get("entidade"), us.get("operacao"), us.get("ativo"))
        saida.append({"texto": o.get("texto") or "", "cor": o.get("cor") or "", "layer": o.get("layer") or "", "e0": eng[0], "o0": eng[1], "a0": eng[2],
                      "e1": user[0], "o1": user[1], "a1": user[2]})
    return saida


def agregar(linhas: list) -> list:
    """Junta linhas iguais (texto, cor, camada, antes e depois), somando `n` e contando as fontes distintas. Entrada: dicts com as chaves do conjunto
    + `fonte` e (opcional) `n`."""
    por = {}
    for l in linhas:
        k = (l["texto"], l["cor"], l["layer"], l["e0"], l["o0"], l["a0"], l["e1"], l["o1"], l["a1"])
        a = por.setdefault(k, {"texto": l["texto"], "cor": l["cor"], "layer": l["layer"], "user": tri(l["e1"], l["o1"], l["a1"]), "n": 0, "fontes": set()})
        a["n"] += int(l.get("n") or 1)
        a["fontes"].add(l.get("fonte"))
    return list(por.values())[:LIMITE_ITENS]


def _norm_layer(l):
    return (l or "").strip().upper()


def _alternativas(valores, max_alt=MAX_ALTERNATIVAS):
    v = sorted({x for x in valores if x})
    return v if 0 < len(v) <= max_alt else None


def _regex_exato(valores) -> str:
    return "^(?:" + "|".join(_esc(v) for v in valores) + ")$" if len(valores) > 1 else "^" + _esc(valores[0]) + "$"


def _ordem_min(regras, fase=None):
    os_ = [r.get("ordem", 0) for r in regras if fase is None or (r.get("fase", 3) == fase)]
    return (min(os_) if os_ else 1) - 1


def _ordem_max(regras, fase=None):
    os_ = [r.get("ordem", 0) for r in regras if fase is None or (r.get("fase", 3) == fase)]
    return (max(os_) if os_ else 0) + 10


def _candidatos(tipo, y, idxs, itens, base, proc, cls):
    """[(tabela_da_regra, regra, descrição)] do mais geral ao mais específico, para o grupo (`tipo`, valor desejado `y`)."""
    layers = {_norm_layer(itens[i]["layer"]) for i in idxs}
    layers_ok = _alternativas(layers, 3) if "" not in layers else None
    ativos = _alternativas([base[i][2] for i in idxs])
    textos = _alternativas([itens[i]["texto"].strip() for i in idxs])
    onde_layer = f" na camada {', '.join(layers_ok)}" if layers_ok else ""
    saida = []
    if tipo == "ent":
        def regra(**c):
            return {"ordem": _ordem_min(cls), "entidade": y, "parar": True, **c}
        if layers_ok and ativos:
            saida.append(("classificacao", regra(layer_em=layers_ok, ativo_regex=_regex_exato(ativos)), f"Ativo {', '.join(ativos)}{onde_layer} → entidade {y}"))
        if ativos:
            saida.append(("classificacao", regra(ativo_regex=_regex_exato(ativos)), f"Ativo {', '.join(ativos)} → entidade {y}"))
        if layers_ok:
            saida.append(("classificacao", regra(layer_em=layers_ok), f"Itens{onde_layer} → entidade {y}"))
        if textos:
            saida.append(("classificacao", regra(texto_regex=_regex_exato(textos)), f"Texto {', '.join(repr(t) for t in textos)} → entidade {y}"))
    elif tipo == "op":
        def tarde(**c):
            return {"fase": 3, "ordem": _ordem_max(proc, 3), "modo": "DEFINIR", "operacao": y, "parar": False, **c}

        def primeiro(**c):
            return {"fase": 3, "ordem": _ordem_min(proc, 3), "modo": "DEFINIR", "operacao": y, "parar": True, **c}
        for mk in (tarde, primeiro):
            if layers_ok:
                saida.append(("processamento", mk(layer_em=layers_ok), f"Itens{onde_layer} → operação {y}"))
            if ativos:
                saida.append(("processamento", {**mk(), "modo": "SUBSTITUIR", "ativo_regex": _regex_exato(ativos), "ativo_template": "$&",
                                                **({"layer_em": layers_ok} if layers_ok else {})}, f"Ativo {', '.join(ativos)}{onde_layer} → operação {y}"))
            if textos:
                saida.append(("processamento", mk(texto_regex=_regex_exato(textos)), f"Texto {', '.join(repr(t) for t in textos)} → operação {y}"))
    elif tipo == "ativo":
        if "{" in y or "}" in y or "$" in y:
            return []
        for ordem, parar in ((_ordem_max(proc, 3), False), (_ordem_min(proc, 3), True)):
            if textos:
                saida.append(("processamento", {"fase": 3, "ordem": ordem, "modo": "DEFINIR", "texto_regex": _regex_exato(textos), "ativo_template": y, "parar": parar},
                              f"Texto {', '.join(repr(t) for t in textos)} → ativo {y}"))
                if layers_ok:
                    saida.append(("processamento", {"fase": 3, "ordem": ordem, "modo": "DEFINIR", "texto_regex": _regex_exato(textos), "layer_em": layers_ok,
                                                    "ativo_template": y, "parar": parar}, f"Texto {', '.join(repr(t) for t in textos)}{onde_layer} → ativo {y}"))
    return saida


def minerar(itens: list, proc: list, cls: list, simular) -> list:
    """Propostas de regra do leitor. `itens` = saída de `agregar`; `proc`/`cls` = regras atuais do projeto; `simular(entrada, proc, cls)` devolve
    [{entidade, operacao, ativo}] (o motor real). Devolve propostas verificadas (ver o topo do arquivo), mais fortes primeiro."""
    if not itens or not (proc or cls):
        return []
    entrada = [{"pagina": 1, "texto": i["texto"], "cor": i["cor"] or None, "layer": i["layer"]} for i in itens]
    base = [tri(r.get("entidade"), r.get("operacao"), r.get("ativo")) for r in simular(entrada, proc, cls)]
    user = [i["user"] for i in itens]
    peso = [i["n"] for i in itens]
    grupos = {}
    for k in range(len(itens)):
        b, u = base[k], user[k]
        if b == u:
            continue
        if b[2] == u[2] and b[0] != u[0]:
            grupos.setdefault(("ent", u[0], ""), []).append(k)
        if b[1] != u[1]:
            grupos.setdefault(("op", u[1], ""), []).append(k)
        if b[2] != u[2]:
            grupos.setdefault(("ativo", u[2], itens[k]["texto"].strip()), []).append(k)
    testados = sum(peso[k] for k in range(len(itens)) if base[k] == user[k])
    campo = {"ent": 0, "op": 1, "ativo": 2}
    candidatos_grupo = []
    for (tipo, y, _), idxs in grupos.items():
        w = sum(peso[i] for i in idxs)
        fontes = set().union(*(itens[i]["fontes"] for i in idxs))
        if w >= MINIMO_PESO and len(fontes) >= MINIMO_FONTES:
            candidatos_grupo.append((w, tipo, y, idxs, fontes))
    # um grupo de "ativo" por texto pode ser fundido: itens com o mesmo ativo desejado e o mesmo tratamento formam UM grupo maior
    candidatos_grupo.sort(key=lambda g: -g[0])
    propostas, vistas = [], set()
    for w, tipo, y, idxs, fontes in candidatos_grupo[:MAX_GRUPOS]:
        for tabela, regra, descricao in _candidatos(tipo, y, idxs, itens, base, proc, cls):
            chave = tabela + "|" + json.dumps(regra, sort_keys=True, ensure_ascii=False)
            if chave in vistas:
                continue
            vistas.add(chave)
            novo = simular(entrada, proc, [regra] + cls) if tabela == "classificacao" else simular(entrada, proc + [regra], cls)
            depois = [tri(r.get("entidade"), r.get("operacao"), r.get("ativo")) for r in novo]
            corrigidos = sum(peso[i] for i in idxs if depois[i][campo[tipo]] == user[i][campo[tipo]])
            # regressão = qualquer CAMPO (entidade, operação, ativo) que estava certo e deixaria de estar, em qualquer item observado
            regressoes = sum(peso[i] for i in range(len(itens)) if any(base[i][f] == user[i][f] and depois[i][f] != user[i][f] for f in range(3)))
            if regressoes == 0 and corrigidos >= MINIMO_PESO and corrigidos >= COBERTURA_MINIMA * w:
                propostas.append({"chave": "leitor|" + chave, "tipo": "regra_leitor", "tabela": tabela, "regra": regra,
                                  "descricao": descricao, "corrigidos": corrigidos, "grupo": w, "testados": testados,
                                  "regressoes": 0, "fontes": len(fontes)})
                break          # a primeira candidata que passa (a mais geral) basta para este grupo
    propostas.sort(key=lambda p: (-p["corrigidos"], p["chave"]))
    return propostas[:MAX_PROPOSTAS]
