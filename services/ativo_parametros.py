"""
services/ativo_parametros.py — Ativos com parâmetros por categoria (TASK-079 e TASK-080).

Um ativo da base de orçamento pode ter variantes declaradas no próprio nome:

    N1=N1(cabo=XLPE, poste=DT)    declaração (1ª linha do ativo): nomes, ordem e PADRÕES  ("N1=N1(cabo, poste)" = sem padrão)
    N1(cabo=XLPE)                 grupo: vale para cabo=XLPE e QUALQUER poste
    N1(cabo=COBRE, poste=CC)      grupo: só para essa combinação
    N1                            grupo comum: vale para qualquer variante

TASK-080 — condições nos grupos da base (o pedido e a declaração continuam com valores concretos):

    N1(poste=D%)                  curinga "%": começa com D   (%T termina em T, %CC% contém CC)
    N1(cabo=ALUM|COBRE)           lista: um ou outro
    N1(cabo!=XLPE)                exceto (vale para a lista inteira: cabo!=ALUM|COBRE = nenhum dos dois)
    N1(poste=D%, cabo!=XLPE)      vários parâmetros na mesma célula = E (OU entre parâmetros = dois grupos)

Em Outros/Totalizadora o pedido usa só os valores, na ordem cadastrada: ``3-N1(XLPE,DT)``; ``3-N1`` completa pelo padrão.
Nomes e valores: letras, números, "_" e "." (sem hífen), sem diferença de caixa (guardados em maiúsculas). ``%``, ``|`` e ``!=`` só nos
grupos da base. COMPONENTE em branco nos grupos é preenchido automaticamente (preencher_componentes).

Este arquivo tem a MESMA lógica de static/ativo_parametros.js (paridade testada em tests/test_ativo_parametros.py).
"""
import re

_RE_ID = re.compile(r"^[A-Z0-9_.]+$")
_RE_ALT_BASE = re.compile(r"^[A-Z0-9_.%]+$")


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


def casar_padrao(padrao, valor) -> bool:
    """Valor casa o padrão (igualdade, ou com "%" = qualquer sequência, inclusive vazia)? Sem diferença de caixa."""
    padrao, valor = str(padrao).upper(), str(valor).upper()
    if "%" not in padrao:
        return padrao == valor
    return re.fullmatch(".*".join(re.escape(p) for p in padrao.split("%")), valor, re.S) is not None


def restricao_vale(restricao, valor) -> bool:
    """restricao = {"op": "="|"!=", "alts": [...]} ; texto simples = igualdade."""
    if isinstance(restricao, str):
        restricao = {"op": "=", "alts": [restricao]}
    acertou = any(casar_padrao(a, valor) for a in restricao["alts"])
    return acertou if restricao["op"] == "=" else not acertou


def _item(item, modo):
    """Interpreta UM item de parâmetro já sem espaços/em maiúsculas. modo: "pedido" | "grupo" | "declaracao".
    Retorna ({"nome", "op", "valor", "alts"}, erro|None)."""
    op = "="
    if "!=" in item:
        nome, _, valor = item.partition("!=")
        op = "!="
    elif "=" in item:
        nome, _, valor = item.partition("=")
    elif modo == "declaracao":
        nome, valor = item, None
    else:
        nome, valor = None, item
    p = {"nome": nome, "op": op, "valor": valor, "alts": []}
    if modo != "grupo" and (op == "!=" or (valor and ("|" in valor or "%" in valor))):
        return p, f"'{item}': %, | e != só valem nos grupos da base"
    if nome is not None and not _RE_ID.match(nome):
        return p, f"'{nome}' inválido (use letras, números, _ e .)"
    if valor is None:
        return p, None
    alts = valor.split("|") if modo == "grupo" else [valor]
    regex = _RE_ALT_BASE if modo == "grupo" else _RE_ID
    for a in alts:
        if not regex.match(a):
            return p, f"'{a}' inválido (use letras, números, _ e .{' ; no grupo da base também % e |' if modo == 'grupo' else ''})"
    p["alts"] = alts
    return p, None


def analisar_ativo(texto, modo_base=False) -> dict:
    """Interpreta o texto de um ativo (pedido, ou célula da base com modo_base=True).

    Retorna {"texto", "base", "tipo": "simples"|"grupo"|"declaracao", "params": [{"nome", "op", "valor", "alts"}], "erro": str|None}.
    "simples" = sem parênteses; "grupo" = com parênteses; "declaracao" = ``N1=N1(...)``. Um item sem "=" é posicional
    (nome None). No pedido só valem valores concretos; operadores (%, |, !=) só nos grupos da base."""
    original = str(texto if texto is not None else "").strip()
    r = {"texto": original, "base": "", "tipo": "simples", "params": [], "erro": None}
    if not original:
        r["erro"] = "ativo vazio"
        return r
    i = original.find("(")
    e = original.find("=")
    if e != -1 and (i == -1 or e < i):
        esq, dir_ = original[:e].strip().upper(), original[e + 1:].strip()
        sub = analisar_ativo(dir_, False)
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
    r["base"] = original[:i].strip().upper()
    r["tipo"] = "grupo"
    if not original.endswith(")") or original.count("(") != 1 or original.count(")") != 1:
        r["erro"] = "parênteses mal formados"
        return r
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
        p, erro = _item(item, "grupo" if modo_base else "pedido")
        if erro:
            r["erro"] = erro
            return r
        if p["nome"] == "":
            r["erro"] = "nome do parâmetro vazio"
            return r
        if p["nome"] is not None and p["valor"] == "":
            r["erro"] = f"parâmetro {p['nome']} sem valor"
            return r
        if p["nome"] is not None:
            if p["nome"] in vistos:
                r["erro"] = f"parâmetro {p['nome']} repetido"
                return r
            vistos.add(p["nome"])
        r["params"].append(p)
    return r


def analisar_declaracao(texto) -> dict:
    """Como analisar_ativo, mas aceita itens só com o nome na parte direita ("N1=N1(cabo, poste)"). Padrões só concretos."""
    original = str(texto if texto is not None else "").strip()
    e = original.find("=")
    i = original.find("(")
    if e == -1 or (i != -1 and e > i):
        return analisar_ativo(original, True)
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
        p, erro = _item(item, "declaracao")
        if erro:
            r["erro"] = erro
            return r
        if p["nome"] in vistos:
            r["erro"] = f"parâmetro {p['nome']} repetido"
            return r
        vistos.add(p["nome"])
        r["params"].append(p)
    return r


def analisar_celula_base(texto) -> dict:
    """Interpreta a célula ATIVO da base (declaração com nomes, grupos com operadores)."""
    t = str(texto if texto is not None else "")
    e = t.find("=")
    i = t.find("(")
    if e != -1 and (i == -1 or e < i):
        return analisar_declaracao(t)
    return analisar_ativo(t, True)


def restricoes_do_grupo(analise) -> dict:
    """{NOME: {"op", "alts"}} dos parâmetros nomeados de um grupo (declaração/simples → {})."""
    if analise["tipo"] != "grupo":
        return {}
    return {p["nome"]: {"op": p["op"], "alts": list(p["alts"])} for p in analise["params"] if p["nome"]}


def montar_assinaturas(celulas) -> dict:
    """A partir das células ATIVO da base (na ordem da tabela, de UM escopo de projeto) devolve, por nome-base,
    {"nomes", "padroes", "valores" (exatos), "mascaras" (padrões com %), "fechados" (parâmetros que só aceitam o cadastrado), "declarada"}.
    Só entram ativos com parâmetros."""
    saida = {}
    for cel in celulas:
        a = analisar_celula_base(cel)
        if a["erro"] or not a["params"]:
            continue
        s = saida.setdefault(a["base"], {"nomes": [], "padroes": {}, "valores": {}, "mascaras": {}, "fechados": [], "declarada": False})
        if a["tipo"] == "declaracao":
            if s["declarada"]:
                continue                                   # 2ª declaração: ignorada aqui (a importação a recusa)
            s["declarada"] = True
            nomes = [p["nome"] for p in a["params"]]
            s["nomes"] = nomes + [n for n in s["nomes"] if n not in nomes]
            for p in a["params"]:
                if p["valor"] is not None:
                    s["padroes"][p["nome"]] = p["valor"]
                    lista = s["valores"].setdefault(p["nome"], [])
                    if p["valor"] not in lista:
                        lista.append(p["valor"])
                    if p["nome"] not in s["fechados"]:
                        s["fechados"].append(p["nome"])
        else:
            for p in a["params"]:
                if p["nome"] is None:
                    continue
                if p["nome"] not in s["nomes"]:
                    s["nomes"].append(p["nome"])
                for alt in p["alts"]:
                    destino = s["mascaras"] if "%" in alt else s["valores"]
                    lista = destino.setdefault(p["nome"], [])
                    if alt not in lista:
                        lista.append(alt)
                if p["op"] == "=" and p["nome"] not in s["fechados"]:
                    s["fechados"].append(p["nome"])
    return saida


def resolver_pedido(texto, assinaturas) -> dict:
    """Normaliza um pedido ("N1", "N1(ALUM)", "N1(poste=CC,cabo=COBRE)").

    Retorna {"base", "texto" (forma normalizada), "params" {NOME: VALOR} (informados + padrões), "faltando" [nomes],
    "erros" [msgs], "com_assinatura" bool}. Completo → "N1(XLPE,DT)"; incompleto → "N1(CABO=ALUM)" (só os informados, com nome).
    O valor precisa ser igual a um valor cadastrado ou casar um padrão (%) do parâmetro; parâmetro só com "!=" aceita qualquer valor."""
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
            if n in sig["fechados"]:
                exatos, mascaras = sig["valores"].get(n, []), sig["mascaras"].get(n, [])
                if dados[n] not in exatos and not any(casar_padrao(m, dados[n]) for m in mascaras):
                    r["erros"].append(f"{n}={dados[n]} não cadastrado em {a['base']} (existem: {', '.join(exatos + mascaras)})")
        else:
            r["faltando"].append(n)
    r["params"] = {n: dados[n] for n in nomes if n in dados}
    if not r["faltando"]:
        r["texto"] = f"{a['base']}({','.join(dados[n] for n in nomes)})"
    elif dados:
        r["texto"] = f"{a['base']}({','.join(n + '=' + dados[n] for n in nomes if n in dados)})"
    return r


def linha_vale(params_linha, params_pedido) -> bool:
    """O grupo da base vale para o pedido? Grupo comum ({}) sempre; cada parâmetro do grupo precisa existir no pedido e satisfazer a
    restrição (texto simples = igualdade; {"op","alts"} = lista/curinga/exceto)."""
    for nome, restricao in (params_linha or {}).items():
        if nome not in (params_pedido or {}) or not restricao_vale(restricao, params_pedido[nome]):
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


# ═══════════════════════════════════════════════════════════════════════════
# Importação: preenchimento do COMPONENTE e validação do cadastro (só no servidor)
# ═══════════════════════════════════════════════════════════════════════════
def _escopo(linha) -> str:
    return (linha.get("projeto") or "").strip().upper()


def _comp(linha) -> str:
    return (linha.get("componente") or "").strip().upper()


def _nome_do_grupo(a) -> str:
    """Nome do COMPONENTE gerado: N1 (grupo comum) ou N1.CABO=XLPE / N1.CABO!=XLPE,POSTE=D% (na ordem em que foi escrito)."""
    if a["tipo"] != "grupo":
        return a["base"]
    return a["base"] + "." + ",".join(f"{p['nome']}{p['op']}{p['valor']}" for p in a["params"] if p["nome"])


def preencher_componentes(linhas):
    """TASK-080: preenche o COMPONENTE em branco dos ativos COM parâmetros. Devolve (linhas_novas, quantidade_preenchida).

    O início de um grupo é a linha com ATIVO (declaração com código, grupo ``N1(cabo=XLPE)`` ou o ``N1`` comum de um ativo
    parametrizado); as linhas seguintes sem ATIVO e sem COMPONENTE herdam o componente do grupo até a próxima linha com ATIVO.
    Componente já digitado nunca muda. Ativos sem parâmetros não são tocados."""
    novas = [dict(l) for l in linhas]
    analises = {}
    parametrizados = set()
    for i, l in enumerate(novas):
        cel = (l.get("ativo") or "").strip()
        if cel:
            a = analisar_celula_base(cel)
            analises[i] = a
            if not a["erro"] and a["tipo"] != "simples":
                parametrizados.add((_escopo(l), a["base"]))
    atual, n = {}, 0
    for i, l in enumerate(novas):
        esc = _escopo(l)
        a = analises.get(i)
        codigo = (l.get("codigo") or "").strip()
        if a is not None:
            if a["erro"] or (esc, a["base"]) not in parametrizados or (a["tipo"] == "declaracao" and not codigo and not _comp(l)):
                atual[esc] = None
                continue
            if not _comp(l) and (codigo or a["tipo"] != "declaracao"):
                l["componente"] = _nome_do_grupo(a)
                n += 1
            atual[esc] = (l.get("componente") or "").strip() or None
        elif atual.get(esc) and not _comp(l) and codigo:
            l["componente"] = atual[esc]
            n += 1
    return novas, n


def _alts_coincidem(a, b) -> bool:
    """Conservador: dois padrões podem casar o mesmo valor?"""
    if "%" not in a and "%" not in b:
        return a.upper() == b.upper()
    if "%" not in a:
        return casar_padrao(b, a)
    if "%" not in b:
        return casar_padrao(a, b)
    # dois padrões: só podem coincidir se os começos e os fins fixos forem compatíveis (D% x C% nunca; D% x %T sim)
    a, b = a.upper(), b.upper()
    ini_a, ini_b, fim_a, fim_b = a.split("%")[0], b.split("%")[0], a.split("%")[-1], b.split("%")[-1]
    return (ini_a.startswith(ini_b) or ini_b.startswith(ini_a)) and (fim_a.endswith(fim_b) or fim_b.endswith(fim_a))


def _compativeis(r1, r2) -> bool:
    """Existe valor que satisfaça as duas restrições (conservador: na dúvida, True)?"""
    if r1["op"] == "=" and r2["op"] == "=":
        return any(_alts_coincidem(x, y) for x in r1["alts"] for y in r2["alts"])
    if r1["op"] == "!=" and r2["op"] == "!=":
        return True
    pos, neg = (r1, r2) if r1["op"] == "=" else (r2, r1)
    return any("%" in p or not any(casar_padrao(n, p) for n in neg["alts"]) for p in pos["alts"])


def _chave_grupo(a):
    return (a["base"], tuple(sorted((p["nome"], p["op"], p["valor"]) for p in a["params"] if p["nome"])) if a["tipo"] == "grupo" else ())


def _rotulo_grupo(chave) -> str:
    base, ps = chave
    return base + ("(" + ",".join(f"{n}{op}{v}" for n, op, v in ps) + ")" if ps else "")


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

    # declaração: uma só, na primeira linha do ativo, cobrindo todos os nomes usados; padrões com grupo que os aceite
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
            for p in decl["params"]:
                if p["valor"] is None:
                    continue
                aceitam = [r[p["nome"]] for _, _, a in itens if a["tipo"] == "grupo" for r in [restricoes_do_grupo(a)] if p["nome"] in r]
                if not any(restricao_vale(r, p["valor"]) for r in aceitam):
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
        grupos_do_comp.setdefault((esc, comp), {}).setdefault(_chave_grupo(a), n)
    for (esc, comp), grupos in grupos_do_comp.items():
        if len(grupos) > 1:
            rotulos = ", ".join(sorted(_rotulo_grupo(g) for g in grupos))
            erros.append(f"Linha {max(grupos.values())}: o COMPONENTE {comp} é usado por grupos diferentes ({rotulos}); cada variante precisa do seu COMPONENTE.")

    # sobreposição: mesmo código em grupos do mesmo ativo que podem valer juntos (exceto LV x LM)
    for (esc, base) in bases_param:
        restr_do_comp = {}
        for (e, comp), g in grupos_do_comp.items():
            chave = next(iter(g))
            if e == esc and chave[0] == base:
                restr_do_comp[comp] = {nome: {"op": op, "alts": v.split("|")} for nome, op, v in chave[1]}
        usos = {}
        for n, linha in enumerate(linhas, start=1):
            comp, cod = _comp(linha), (linha.get("codigo") or "").strip().upper()
            if cod and comp in restr_do_comp and _escopo(linha) == esc:
                usos.setdefault(cod, []).append((comp, (linha.get("lv") or "A").strip().upper() or "A", n))
        for cod, lista in usos.items():
            for x in range(len(lista)):
                for y in range(x + 1, len(lista)):
                    (c1, lv1, n1), (c2, lv2, n2) = lista[x], lista[y]
                    if c1 == c2 or {lv1, lv2} == {"LV", "LM"}:
                        continue
                    p1, p2 = restr_do_comp[c1], restr_do_comp[c2]
                    if all(_compativeis(p1[k], p2[k]) for k in set(p1) & set(p2)):
                        avisos.append(f"Código {cod} aparece nos grupos {c1} (linha {n1}) e {c2} (linha {n2}), que podem valer juntos: a quantidade soma duas vezes.")
    return {"erros": erros, "avisos": avisos}
