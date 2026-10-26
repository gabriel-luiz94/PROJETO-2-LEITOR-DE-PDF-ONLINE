"""
services/ativo_parametros.py — Ativos com parâmetros por categoria (TASK-079, 1ª entrega).

Um ativo da base de orçamento pode ter variantes declaradas no próprio nome:

    N1=N1(cabo=XLPE, poste=DT)    declaração (1ª linha do ativo): nomes, ordem e PADRÕES  ("N1=N1(cabo, poste)" = sem padrão)
    N1(cabo=XLPE)                 grupo: vale para cabo=XLPE e QUALQUER poste
    N1(cabo=COBRE, poste=CC)      grupo: só para essa combinação
    N1                            grupo comum: vale para qualquer variante

Em Outros/Totalizadora o pedido usa só os valores, na ordem cadastrada: ``3-N1(XLPE,DT)``; ``3-N1`` completa pelo padrão.
Nomes e valores: letras, números, "_" e "." (sem hífen), sem diferença de caixa (guardados em maiúsculas).

Este arquivo tem a MESMA lógica de static/ativo_parametros.js (paridade testada em tests/test_ativo_parametros.py).
"""
import re

_RE_ID = re.compile(r"^[A-Z0-9_.]+$")


def _norm(txt) -> str:
    return re.sub(r"\s+", "", str(txt if txt is not None else "")).upper()


def nome_base(texto) -> str:
    """"N1(XLPE,DT)" / "N1=N1(cabo=XLPE)" / "N1" → "N1" (maiúsculas, sem parâmetros)."""
    t = str(texto if texto is not None else "").strip()
    i = t.find("(")
    e = t.find("=")
    if e != -1 and (i == -1 or e < i):
        t = t[:e]
    elif i != -1:
        t = t[:i]
    return t.strip().upper()


def analisar_ativo(texto) -> dict:
    """Interpreta o texto de um ativo (base ou pedido).

    Retorna {"texto", "base", "tipo": "simples"|"grupo"|"declaracao", "params": [{"nome": str|None, "valor": str|None}], "erro": str|None}.
    "simples" = sem parênteses; "grupo" = com parênteses; "declaracao" = ``N1=N1(...)``. Em pedidos/grupos um item sem "=" é posicional
    ({"nome": None, "valor": X}); em declarações um item sem "=" é só o nome ({"nome": X, "valor": None}).
    """
    original = str(texto if texto is not None else "").strip()
    r = {"texto": original, "base": "", "tipo": "simples", "params": [], "erro": None}
    if not original:
        r["erro"] = "ativo vazio"
        return r
    i = original.find("(")
    e = original.find("=")
    declaracao = e != -1 and (i == -1 or e < i)
    if declaracao:
        esq, dir_ = original[:e].strip().upper(), original[e + 1:].strip()
        sub = analisar_ativo(dir_)
        r.update(base=esq, tipo="declaracao", params=sub["params"])
        if sub["erro"]:
            r["erro"] = sub["erro"]
        elif sub["tipo"] != "grupo":
            r["erro"] = "declaração precisa ter parâmetros: N1=N1(cabo=XLPE, poste=DT)"
        elif sub["base"] != esq:
            r["erro"] = f"declaração com nomes diferentes ({esq} = {sub['base']})"
        return r
    if i == -1:
        r["base"] = original.upper()
        return r
    if not original.endswith(")") or original.count("(") != 1 or original.count(")") != 1:
        r["base"] = original[:i].strip().upper()
        r["tipo"] = "grupo"
        r["erro"] = "parênteses mal formados"
        return r
    r["base"] = original[:i].strip().upper()
    r["tipo"] = "grupo"
    corpo = original[i + 1:-1]
    if not corpo.strip():
        r["erro"] = "parênteses vazios"
        return r
    vistos = set()
    for item in corpo.split(","):
        item = _norm(item)
        if not item:
            r["erro"] = "parâmetro vazio"
            return r
        if "=" in item:
            nome, _, valor = item.partition("=")
            par = {"nome": nome, "valor": valor}
        else:
            par = {"nome": None, "valor": item}
        for campo in ("nome", "valor"):
            v = par[campo]
            if v is not None and not _RE_ID.match(v):
                r["erro"] = f"'{v}' inválido (use letras, números, _ e .)"
                return r
        if par["nome"] == "":
            r["erro"] = "nome do parâmetro vazio"
            return r
        if par["valor"] == "" and par["nome"] is not None:
            r["erro"] = f"parâmetro {par['nome']} sem valor"
            return r
        if par["nome"] is not None:
            if par["nome"] in vistos:
                r["erro"] = f"parâmetro {par['nome']} repetido"
                return r
            vistos.add(par["nome"])
        r["params"].append(par)
    return r


def analisar_declaracao(texto) -> dict:
    """Como analisar_ativo, mas aceita itens só com o nome na parte direita ("N1=N1(cabo, poste)")."""
    original = str(texto if texto is not None else "").strip()
    e = original.find("=")
    i = original.find("(")
    if e == -1 or (i != -1 and e > i):
        return analisar_ativo(original)
    esq, dir_ = original[:e].strip().upper(), original[e + 1:].strip()
    r = {"texto": original, "base": esq, "tipo": "declaracao", "params": [], "erro": None}
    j = dir_.find("(")
    if j == -1 or not dir_.endswith(")") or dir_.count("(") != 1 or dir_.count(")") != 1:
        r["erro"] = "declaração precisa ter parâmetros: N1=N1(cabo=XLPE, poste=DT)"
        return r
    if dir_[:j].strip().upper() != esq:
        r["erro"] = f"declaração com nomes diferentes ({esq} = {dir_[:j].strip().upper()})"
        return r
    corpo = dir_[j + 1:-1]
    if not corpo.strip():
        r["erro"] = "parênteses vazios"
        return r
    vistos = set()
    for item in corpo.split(","):
        item = _norm(item)
        if not item:
            r["erro"] = "parâmetro vazio"
            return r
        nome, ig, valor = item.partition("=")
        if not _RE_ID.match(nome) or (ig and not _RE_ID.match(valor)):
            r["erro"] = f"'{item}' inválido (use letras, números, _ e .)"
            return r
        if nome in vistos:
            r["erro"] = f"parâmetro {nome} repetido"
            return r
        vistos.add(nome)
        r["params"].append({"nome": nome, "valor": valor if ig else None})
    return r


def analisar_celula_base(texto) -> dict:
    """Interpreta a célula ATIVO da base (aceita declaração só com nomes)."""
    t = str(texto if texto is not None else "")
    e = t.find("=")
    i = t.find("(")
    if e != -1 and (i == -1 or e < i):
        return analisar_declaracao(t)
    return analisar_ativo(t)


def montar_assinaturas(celulas) -> dict:
    """A partir das células ATIVO da base (na ordem da tabela, de UM escopo de projeto) devolve, por nome-base,
    {"nomes": [...], "padroes": {nome: valor}, "valores": {nome: [...]}, "declarada": bool}. Só entram ativos com parâmetros."""
    saida = {}
    for cel in celulas:
        a = analisar_celula_base(cel)
        if a["erro"] or not a["params"]:
            continue
        s = saida.setdefault(a["base"], {"nomes": [], "padroes": {}, "valores": {}, "declarada": False})
        if a["tipo"] == "declaracao":
            if s["declarada"]:
                continue                                   # 2ª declaração: ignorada aqui (a importação a recusa)
            s["declarada"] = True
            nomes = [p["nome"] for p in a["params"]]
            s["nomes"] = nomes + [n for n in s["nomes"] if n not in nomes]
            for p in a["params"]:
                if p["valor"] is not None:
                    s["padroes"][p["nome"]] = p["valor"]
                    s["valores"].setdefault(p["nome"], [])
                    if p["valor"] not in s["valores"][p["nome"]]:
                        s["valores"][p["nome"]].append(p["valor"])
        else:
            for p in a["params"]:
                if p["nome"] is None:
                    continue
                if p["nome"] not in s["nomes"]:
                    s["nomes"].append(p["nome"])
                s["valores"].setdefault(p["nome"], [])
                if p["valor"] not in s["valores"][p["nome"]]:
                    s["valores"][p["nome"]].append(p["valor"])
    return saida


def resolver_pedido(texto, assinaturas) -> dict:
    """Normaliza um pedido ("N1", "N1(ALUM)", "N1(poste=CC,cabo=COBRE)").

    Retorna {"base", "texto" (forma normalizada), "params" {NOME: VALOR} (informados + padrões), "faltando" [nomes],
    "erros" [msgs], "com_assinatura" bool}. Completo → "N1(XLPE,DT)"; incompleto → "N1(CABO=ALUM)" (só os informados, com nome)."""
    a = analisar_ativo(texto)
    r = {"base": a["base"], "texto": a["base"], "params": {}, "faltando": [], "erros": [], "com_assinatura": False}
    if a["erro"]:
        r["erros"].append(a["erro"])
        r["texto"] = a["texto"].upper()
        return r
    if a["tipo"] == "declaracao":
        r["erros"].append("declaração não pode ser usada como pedido")
        return r
    sig = (assinaturas or {}).get(a["base"])
    if sig is None:
        if a["params"]:
            r["erros"].append(f"{a['base']} não tem variantes cadastradas")
            r["texto"] = a["texto"].upper().replace(" ", "")
        return r
    r["com_assinatura"] = True
    nomes = sig["nomes"]
    dados = {}
    pos = 0
    for p in a["params"]:
        if p["nome"] is None:
            while pos < len(nomes) and nomes[pos] in dados:
                pos += 1
            if pos >= len(nomes):
                r["erros"].append(f"parâmetros demais para {a['base']} (esperado: {', '.join(nomes)})")
                continue
            dados[nomes[pos]] = p["valor"]
            pos += 1
        else:
            if p["nome"] not in nomes:
                r["erros"].append(f"parâmetro {p['nome']} desconhecido em {a['base']} (existem: {', '.join(nomes)})")
            elif p["nome"] in dados:
                r["erros"].append(f"parâmetro {p['nome']} repetido")
            else:
                dados[p["nome"]] = p["valor"]
    for n in nomes:
        if n not in dados and n in sig["padroes"]:
            dados[n] = sig["padroes"][n]
    for n in nomes:
        if n in dados:
            existentes = sig["valores"].get(n, [])
            if existentes and dados[n] not in existentes:
                r["erros"].append(f"{n}={dados[n]} não cadastrado em {a['base']} (existem: {', '.join(existentes)})")
        else:
            r["faltando"].append(n)
    r["params"] = {n: dados[n] for n in nomes if n in dados}
    if not r["faltando"]:
        r["texto"] = f"{a['base']}({','.join(dados[n] for n in nomes)})"
    elif dados:
        r["texto"] = f"{a['base']}({','.join(n + '=' + dados[n] for n in nomes if n in dados)})"
    return r


def linha_vale(params_linha, params_pedido) -> bool:
    """O grupo da base (params_linha) vale para o pedido? Grupo comum ({}) sempre; cada parâmetro do grupo precisa existir e ser igual."""
    for nome, valor in (params_linha or {}).items():
        if (params_pedido or {}).get(nome) != valor:
            return False
    return True


def dividir_tokens(texto) -> list:
    """Divide o texto de uma linha de Outros por espaços, sem quebrar o que está entre parênteses ("2-N1(XLPE, DT)" é um token só)."""
    saida, atual, prof = [], [], 0
    for ch in str(texto if texto is not None else ""):
        if ch == "(":
            prof += 1
        elif ch == ")" and prof > 0:
            prof -= 1
        if ch.isspace() and prof == 0:
            if atual:
                saida.append("".join(atual))
                atual = []
        else:
            atual.append(ch)
    if atual:
        saida.append("".join(atual))
    return saida


def _escopo(linha) -> str:
    return (linha.get("projeto") or "").strip().upper()


def _comp(linha) -> str:
    return (linha.get("componente") or "").strip().upper()


def validar_base(linhas) -> dict:
    """Valida as linhas da base (dicts com ativo, componente, codigo, projeto, lv…, na ordem da tabela).
    Retorna {"erros": [...], "avisos": [...]}. Escopo: por (projeto, nome-base). Linhas sem parênteses nem "=" passam direto."""
    erros, avisos = [], []
    analises = []                                                     # (n, linha, analise) de TODA linha com ativo
    for n, linha in enumerate(linhas, start=1):
        cel = (linha.get("ativo") or "").strip()
        if not cel:
            continue
        a = analisar_celula_base(cel)
        if a["erro"]:
            erros.append(f"Linha {n}: ativo '{cel}': {a['erro']}.")
            continue
        if a["tipo"] == "grupo" and any(p["nome"] is None for p in a["params"]):
            erros.append(f"Linha {n}: ativo '{cel}': na base os parâmetros levam nome (cabo=XLPE).")
            continue
        analises.append((n, linha, a))

    parametrizadas = {}                                               # (escopo, base) → [(n, linha, analise)] das linhas com parâmetros
    for n, linha, a in analises:
        if a["tipo"] != "simples":
            parametrizadas.setdefault((_escopo(linha), a["base"]), []).append((n, linha, a))

    # declaração: uma só, na primeira linha do ativo, cobrindo todos os nomes usados; padrões com grupo cadastrado
    for (esc, base), itens in parametrizadas.items():
        todas = [(n, a) for n, l, a in analises if _escopo(l) == esc and a["base"] == base]
        decls = [(n, a) for n, l, a in itens if a["tipo"] == "declaracao"]
        if len(decls) > 1:
            erros.append(f"Linha {decls[1][0]}: {base} tem mais de uma declaração ({base}=…); use só uma, na primeira linha do ativo.")
        if decls:
            n_decl, decl = decls[0]
            if n_decl != min(n for n, _ in todas):
                erros.append(f"Linha {n_decl}: a declaração de {base} precisa ser a primeira linha do ativo (linha {min(n for n, _ in todas)}).")
            nomes = [p["nome"] for p in decl["params"]]
            for n, l, a in itens:
                if a["tipo"] == "grupo":
                    for p in a["params"]:
                        if p["nome"] not in nomes:
                            erros.append(f"Linha {n}: parâmetro {p['nome']} não está na declaração de {base} ({', '.join(nomes)}).")
            usados = {}
            for n, l, a in itens:
                if a["tipo"] == "grupo":
                    for p in a["params"]:
                        usados.setdefault(p["nome"], set()).add(p["valor"])
            for p in decl["params"]:
                if p["valor"] is not None and p["valor"] not in usados.get(p["nome"], set()):
                    erros.append(f"Linha {n_decl}: padrão {p['nome']}={p['valor']} de {base} não tem nenhum grupo cadastrado.")

    # grupo de cada COMPONENTE (por escopo); componente usado por grupos diferentes é erro (a expansão traria os dois)
    bases_param = {k for k in parametrizadas}
    grupos_do_comp = {}                                               # (escopo, comp) → {chave_do_grupo: n}
    for n, linha, a in analises:
        comp, esc = _comp(linha), _escopo(linha)
        if not comp or (esc, a["base"]) not in bases_param:
            continue
        if a["tipo"] == "declaracao" and not (linha.get("codigo") or "").strip():
            continue                                                  # linha só de declaração
        params = tuple(sorted((p["nome"], p["valor"]) for p in a["params"])) if a["tipo"] == "grupo" else ()
        grupos_do_comp.setdefault((esc, comp), {}).setdefault((a["base"], params), n)
    for (esc, comp), grupos in grupos_do_comp.items():
        if len(grupos) > 1:
            rotulos = ", ".join(sorted(b + ("(" + ",".join(k + "=" + v for k, v in ps) + ")" if ps else "") for (b, ps) in grupos))
            erros.append(f"Linha {max(grupos.values())}: o COMPONENTE {comp} é usado por grupos diferentes ({rotulos}); cada variante precisa do seu COMPONENTE.")

    # sobreposição: mesmo código em grupos do mesmo ativo que podem valer juntos (exceto LV x LM)
    for (esc, base) in bases_param:
        params_do_comp = {comp: dict(next(iter(g))[1]) for (e, comp), g in grupos_do_comp.items() if e == esc and next(iter(g))[0] == base}
        usos = {}
        for n, linha in enumerate(linhas, start=1):
            comp, cod = _comp(linha), (linha.get("codigo") or "").strip().upper()
            if cod and comp in params_do_comp and _escopo(linha) == esc:
                usos.setdefault(cod, []).append((comp, (linha.get("lv") or "A").strip().upper() or "A", n))
        for cod, lista in usos.items():
            for x in range(len(lista)):
                for y in range(x + 1, len(lista)):
                    (c1, lv1, n1), (c2, lv2, n2) = lista[x], lista[y]
                    if c1 == c2 or {lv1, lv2} == {"LV", "LM"}:
                        continue
                    p1, p2 = params_do_comp[c1], params_do_comp[c2]
                    if all(p1[k] == p2[k] for k in set(p1) & set(p2)):
                        avisos.append(f"Código {cod} aparece nos grupos {c1} (linha {n1}) e {c2} (linha {n2}), que podem valer juntos: a quantidade soma duas vezes.")
    return {"erros": erros, "avisos": avisos}


def assinaturas_do_projeto(linhas, projeto) -> dict:
    """Assinaturas das linhas da base válidas para o projeto (genéricas + as do projeto), na ordem da tabela (espelha o JS)."""
    p = str(projeto or "").strip().upper()
    validas = []
    for l in linhas or []:
        lp = str(l.get("projeto") or "").strip().upper()
        if not lp or not p or p in [x.strip() for x in lp.split("/")]:
            validas.append(l)
    return montar_assinaturas([l.get("ativo") or "" for l in validas])


def resumo_erros(erros, limite=8) -> str:
    """Junta as primeiras mensagens de erro para uma resposta HTTP (as demais viram "… e mais N")."""
    msg = " | ".join(erros[:limite])
    return msg + (f" | … e mais {len(erros) - limite}" if len(erros) > limite else "")
