"""
services/orcamento_calc.py — Lógica de cálculo do orçamento.
"""
import re


from services.orcamento_lv import operacao_viva, passa_lv
from services.anotacao_outros import contexto_da_base, interpretar_linha
from services.ativo_parametros import (
    analisar_celula_base, linha_vale, montar_assinaturas, nome_base, resolver_pedido, restricoes_do_grupo,
)


def extrair_ativo_cabo(item: dict):
    """Interpreta uma linha de CABOS ("ATIVO FASE COMPRIMENTO"). Devolve o dict de ativo ou None.

    Tolerante por desenho: fase/comprimento ausentes ou não numéricos ficam em 1.0 sem aviso.
    A validação (services/validacao_planilhas.py) reutiliza esta função para não divergir do cálculo.
    """
    ativo_raw = item.get("ativo", "").strip().upper()
    operacao  = item.get("operacao", "").strip().upper()
    if not ativo_raw:
        return None

    txt = ativo_raw

    parts = txt.split()
    if not parts:
        return None

    nome_ativo = parts[0]

    fases = 1.0
    comprimento = 1.0

    if len(parts) >= 3:
        try:
            fases = float(parts[1].replace(",", "."))
        except Exception:
            pass
        try:
            comprimento = float(parts[2].replace(",", "."))
        except Exception:
            pass
    elif len(parts) == 2:
        # Fallback caso tenham omitido fase ou comprimento
        try:
            comprimento = float(parts[1].replace(",", "."))
        except Exception:
            pass

    # Prioriza qtdAtivos calculado pelo frontend; fallback para lógica local
    qtd_ativos_frontend = item.get("qtdAtivos")
    if qtd_ativos_frontend is not None:
        try:
            multiplicador = float(qtd_ativos_frontend)
        except (ValueError, TypeError):
            multiplicador = fases
    elif nome_ativo.startswith("CAZ") or nome_ativo.startswith("M"):
        multiplicador = 1.0
    else:
        multiplicador = fases

    # Calcula a quantidade final processada (comprimento * multiplicador).
    # A margem de perdas (ex: +5%) agora é gerida pela Tabela de Regras no frontend.
    qtd_final = comprimento * multiplicador

    return {"ativo": nome_ativo, "qtd": qtd_final, "operacao": operacao, "origem": "CABOS"}


def tokenizar_outros(ativo_raw: str) -> list:
    """Quebra o texto de uma linha de OUTROS em tokens alternados [qtd, ativo, qtd, ativo, ...].

    TASK-079: o que está entre parênteses ("N1(XLPE,DT)", "N1(poste=CC, cabo=COBRE)") é protegido — nenhum "-", "*" ou espaço
    de dentro dele é tratado como separador/sinal.
    """
    txt = (ativo_raw or "").strip()
    if not txt:
        return []
    protegidos = []

    def _guarda(m):
        protegidos.append(m.group(0))
        return f"\x00{len(protegidos) - 1}\x00"

    txt = re.sub(r"\([^()]*\)", _guarda, txt)
    # Se inicia com DT ou CV sem quantidade, adiciona "1-"
    if re.match(r'^(DT|CV)', txt, re.IGNORECASE) and not re.match(r'^\d', txt):
        txt = "1-" + txt
    # "-" → " " (separa quantidade do ativo)
    txt = txt.replace("-", " ")
    # "*" → "-" (representa negativo / linha viva)
    txt = txt.replace("*", "-")
    return [re.sub(r"\x00(\d+)\x00", lambda m: protegidos[int(m.group(1))], tok) for tok in txt.split()]


def _classico_outros(texto: str) -> list:
    """Formato ANTIGO de uma linha de OUTROS ("[qtd]-[ativo] [qtd]-[ativo] ..."): [{"ativo", "qtd"}].

    Tolerante por desenho: quantidade não numérica vira 1.0 e um token sobrando no fim é descartado.
    """
    # Posições pares = quantidades, posições ímpares = ativos
    tokens = tokenizar_outros(texto)
    resultado = []
    i = 0
    while i + 1 < len(tokens):
        try:
            qtd = float(tokens[i].replace(",", "."))
        except ValueError:
            qtd = 1.0
        resultado.append({"ativo": tokens[i + 1].upper(), "qtd": qtd})
        i += 2
    return resultado


def extrair_ativos_outros(item: dict, ctx: dict = None) -> list:
    """Interpreta uma linha de OUTROS e devolve a lista de ativos {ativo, qtd, operacao, origem}.

    TASK-081: aceita o formato ANTIGO (``2-N3 1-T3``) e o COMPACTO (``11600D, 2 N3, T3``) — nunca misturados —, resolve apelidos e notação
    compacta de poste, preenche o parâmetro ``poste`` das estruturas e aplica a regra do M/*M (poste ignorado, demais itens I/*I).
    ``ctx`` (opcional) = contexto_da_base(...). Linha com erro de formato devolve só o que se entende sem ambiguidade (nada).
    """
    operacao = item.get("operacao", "").strip().upper()
    r = interpretar_linha(item.get("ativo", ""), operacao, _classico_outros, ctx)
    return [{"ativo": i["ativo"], "qtd": i["qtd"], "operacao": i["operacao"], "origem": "OUTROS"} for i in r["itens"]]


def _linha_vale_no_projeto(row, req_proj) -> bool:
    row_proj = (row.get("projeto") or "").strip().upper()
    if not row_proj:
        return True
    return not req_proj or req_proj in [p.strip() for p in row_proj.split("/")]


def _preparar_parametros(orcamento_rows, req_proj):
    """TASK-079: ativos com parâmetros na base. Devolve (linhas, assinaturas, grupos).

    Cada GRUPO (``N1(cabo=XLPE)``) vira uma chave própria ``N1#CABO=XLPE`` no índice por ativo; a linha de declaração
    (``N1=N1(…)``) vira o grupo comum ``N1`` (se tiver código/componente) ou some (se for só declaração). Base sem parâmetros
    passa intacta. ``grupos`` = {base: {chave: {NOME: VALOR}}} só com as linhas válidas para o projeto do pedido.
    """
    validas = [r for r in orcamento_rows if _linha_vale_no_projeto(r, req_proj)]
    assinaturas = montar_assinaturas([r.get("ativo") or "" for r in validas])
    if not assinaturas and not any("(" in (r.get("ativo") or "") or "," in (r.get("ativo") or "") for r in orcamento_rows):
        return orcamento_rows, assinaturas, {}

    def _chave(a):
        return a["base"] + "#" + ",".join(f"{p['nome']}{p['op']}{p['valor']}" for p in sorted(a["params"], key=lambda p: p["nome"]))

    linhas, grupos = [], {}
    for r in orcamento_rows:
        cel = (r.get("ativo") or "").strip()
        a = analisar_celula_base(cel) if cel and ("(" in cel or "=" in cel or "," in cel) else None
        if a is None or a["erro"]:
            linhas.append(r)
            continue
        novo = dict(r)
        if a["tipo"] == "simples":                         # TASK-081: "DT11/300, 11300D" → o nome oficial é o 1º
            novo["ativo"] = a["base"]
            linhas.append(novo)
            continue
        if a["tipo"] == "declaracao":
            if not ((r.get("codigo") or "").strip() or (r.get("componente") or "").strip()):
                continue                                   # linha só de declaração: nunca entra no cálculo
            novo["ativo"] = a["base"]                      # a declaração é a 1ª linha do grupo comum
        else:
            novo["ativo"] = _chave(a)
            if r in validas:
                grupos.setdefault(a["base"], {})[novo["ativo"]] = restricoes_do_grupo(a)
        linhas.append(novo)
    return linhas, assinaturas, grupos


def processar_calculo(req_cabos: list, req_outros: list, req_projeto: str, orcamento_rows: list) -> dict:
    req_proj = (req_projeto or "").strip().upper()
    ctx_outros = contexto_da_base(orcamento_rows, req_proj)          # TASK-081: apelidos, postes e assinaturas (antes de normalizar a base)
    orcamento_rows, assinaturas, grupos_param = _preparar_parametros(orcamento_rows, req_proj)
    avisos = []

    # ── 1. Índice ATIVO → linhas do orçamento (O(1) lookup com Fallback de Projeto) ──
    
    raw_by_ativo = {}
    for r in orcamento_rows:
        key = (r.get("ativo") or "").strip().upper()
        if key:
            raw_by_ativo.setdefault(key, []).append(dict(r))

    orcamento_dict = {}
    for key, rows in raw_by_ativo.items():
        exact_matches = []
        generic_matches = []
        fallback_matches = []

        for r in rows:
            row_proj = (r.get("projeto") or "").strip().upper()
            if not row_proj:
                generic_matches.append(r)
            else:
                valid_projs = [p.strip() for p in row_proj.split("/")]
                if req_proj and req_proj in valid_projs:
                    exact_matches.append(r)
                else:
                    fallback_matches.append(r)

        if exact_matches:
            orcamento_dict[key] = exact_matches
        elif generic_matches:
            orcamento_dict[key] = generic_matches
        elif fallback_matches:
            orcamento_dict[key] = fallback_matches
        else:
            orcamento_dict[key] = rows

    # ── 2. Extrair tabela interna [ativo, qtd, operacao] ────────────────────
    ativos_qtd = []

    # 2a. CABOS: Regra rígida de formato "ATIVO FASE COMPRIMENTO"
    for item in req_cabos:
        av = extrair_ativo_cabo(item)
        if av:
            ativos_qtd.append(av)

    # 2b. OUTROS: texto separado por espaços no modelo "[qtd]-[ativo] [qtd]-[ativo] ..."
    for item in req_outros:
        ativos_qtd.extend(extrair_ativos_outros(item, ctx_outros))

    # ── 3. Cross-reference: Ativo → 1º match → Componente → todos os Códigos ─
    
    # Filtrar orcamento_rows (todas as linhas) pelo projeto
    valid_rows = []
    for r in orcamento_rows:
        row_proj = (r.get("projeto") or "").strip().upper()
        if not row_proj:
            valid_rows.append(r)
        else:
            valid_projs = [p.strip() for p in row_proj.split("/")]
            if not req_proj or req_proj in valid_projs:
                valid_rows.append(r)

    # Índice auxiliar: COMPONENTE → lista de todas as linhas daquele componente
    componente_dict = {}
    for row in valid_rows:
        comp = row.get("componente", "").strip().upper()
        if comp:
            componente_dict.setdefault(comp, []).append(row)

    # Mapa garantido: codigo → row representativa (prioriza as que têm MDO)
    codigo_lookup = {}
    for row in valid_rows:
        codigo = row.get("codigo")
        if not codigo:
            continue
        if codigo not in codigo_lookup:
            codigo_lookup[codigo] = row
        elif (row.get("mdo") or "").strip() and not (codigo_lookup[codigo].get("mdo") or "").strip():
            codigo_lookup[codigo] = row

    componentes_qtd = []
    nao_encontrados = []

    # Índice auxiliar: DESC_ATIVO (upper) → lista de rows (para busca parcial)
    desc_ativo_index = {}
    for row in valid_rows:
        desc = (row.get("desc_ativo") or "").strip().upper()
        if desc:
            desc_ativo_index.setdefault(desc, []).append(row)

    def _buscar(nome_ativo, aceita):
        """Cascata de busca (ativo → componente → desc exata → desc parcial → código). Retorna ("linhas", [rows]) | ("codigo", row) | None."""
        linhas = [r for r in orcamento_dict.get(nome_ativo, []) if aceita(r)]
        if not linhas and nome_ativo in componente_dict:                      # Passo 2: COMPONENTE exato
            linhas = [r for r in componente_dict[nome_ativo] if aceita(r)]
        if not linhas and nome_ativo in desc_ativo_index:                     # Passo 3: DESC_ATIVO exato
            linhas = [r for r in desc_ativo_index[nome_ativo] if aceita(r)]
        if not linhas:                                                        # Passo 4: busca parcial em DESC_ATIVO
            matches = [rows for desc, rows in desc_ativo_index.items() if nome_ativo in desc]
            if matches:
                linhas = [r for r in matches[0] if aceita(r)]
        if linhas:
            return ("linhas", linhas)
        if nome_ativo in codigo_lookup and aceita(codigo_lookup[nome_ativo]):  # Passo 5: CÓDIGO direto
            return ("codigo", codigo_lookup[nome_ativo])
        return None

    def _processar_um(nome_ativo, av, rotulo):
        """Resolve UMA chave do índice (cascata completa). Retorna (achou, dica_de_linha_viva_morta|None)."""
        origem_req = av.get("origem", "").upper()
        operacao_req = av["operacao"]

        def _matches_origem(row):
            row_origem = (row.get("origem") or "").strip().upper()
            return not row_origem or row_origem == origem_req

        def _aceita(row):                       # origem + linha viva/morta (TASK-078)
            return _matches_origem(row) and passa_lv(row.get("lv"), operacao_req)

        achado = _buscar(nome_ativo, _aceita)
        if achado is None:
            if _buscar(nome_ativo, _matches_origem) is not None:
                # existe na base, mas só com o outro tipo de linha: dica para quem corrige o orçamento
                return False, f"só existe para linha {'morta' if operacao_viva(operacao_req) else 'viva'}"
            return False, None

        if achado[0] == "codigo":
            row_clone = dict(achado[1])
            row_clone["fator_i"] = 1.0
            row_clone["fator_r"] = 1.0
            componentes_qtd.append({"qtd": av["qtd"], "operacao": av["operacao"], "row": row_clone})
            return True, None

        linhas_ativo = achado[1]
        # Pega o COMPONENTE do primeiro match do ativo
        componente = linhas_ativo[0].get("componente", "").strip().upper()
        if not componente:
            return True, None

        # Agora busca TODOS os códigos desse componente filtrados pela origem e pelo tipo de linha
        # Para cada código, garante que a row escolhida tem mdo preenchido (se existir)
        linhas_componente = [r for r in componente_dict.get(componente, []) if _aceita(r)]

        # Agrupa linhas por codigo, priorizando as que têm mdo não-vazio
        melhor_row_por_codigo = {}
        for row in linhas_componente:
            codigo = row["codigo"]
            mdo_row = (row.get("mdo") or "").strip()
            if codigo not in melhor_row_por_codigo:
                melhor_row_por_codigo[codigo] = row
            elif mdo_row and not (melhor_row_por_codigo[codigo].get("mdo") or "").strip():
                # Substitui por essa que tem mdo preenchido
                melhor_row_por_codigo[codigo] = row

        for codigo, row in melhor_row_por_codigo.items():
            componentes_qtd.append({
                "qtd":      av["qtd"],
                "operacao": av["operacao"],
                "row":      row,
            })
        return True, None

    for av in ativos_qtd:
        nome_ativo = av["ativo"]
        base = nome_base(nome_ativo)

        if base in assinaturas or "(" in nome_ativo:
            # TASK-079: ativo com parâmetros → grupo comum + grupos que valem para os parâmetros pedidos
            pedido = resolver_pedido(nome_ativo, assinaturas)
            if pedido["erros"]:
                nao_encontrados.append(f"{nome_ativo} ({'; '.join(pedido['erros'])})")
                continue
            if pedido["faltando"]:
                sig = assinaturas[base]
                avisos.append(
                    f"{pedido['texto']}: faltou informar " + "; ".join(f"{n} ({', '.join(sig['valores'].get(n, []))})" for n in pedido["faltando"])
                    + f" — calculado só o que não depende disso. Ex.: {base}({','.join(sig['padroes'].get(n) or sig['valores'].get(n, ['?'])[0] for n in sig['nomes'])})."
                )
            chaves = [base] if orcamento_dict.get(base) else []
            chaves += [k for k, params in grupos_param.get(base, {}).items() if linha_vale(params, pedido["params"])]
            achou, dica = False, None
            for chave in chaves:
                ok, d = _processar_um(chave, av, pedido["texto"])
                achou = achou or ok
                dica = dica or d
            if not achou:
                nao_encontrados.append(f"{nome_ativo} ({dica or 'variante não cadastrada'})")
            continue

        achou, dica = _processar_um(nome_ativo, av, nome_ativo)
        if not achou:
            nao_encontrados.append(f"{nome_ativo} ({dica})" if dica else nome_ativo)

    # ── 4. Soma I e Soma R por CÓDIGO (equivalente ao SOMASE do Excel) ───────
    codigos = {}
    for c in componentes_qtd:
        row     = c["row"]
        codigo  = row["codigo"]
        # MDO: da linha, ou fallback pelo código — nunca fica vazio (usa "-" se não achar em lugar nenhum)
        fallback_mdo = codigo_lookup.get(codigo, {}).get("mdo", "")
        mdo     = (row.get("mdo") or "").strip() or (fallback_mdo or "").strip() or "-"
        key     = (codigo, mdo)
        is_inst = c["operacao"] in ("I", "*I")
        fator_i = float(row.get("fator_i") or 0)
        fator_r = float(row.get("fator_r") or 0)
        if key not in codigos:
            codigos[key] = {
                "mdo":       mdo,
                "codigo":    codigo,
                "desc_codigo": row.get("desc_codigo", ""),
                "filtro":    row.get("filtro", ""),
                "soma_i":    0.0,
                "soma_r":    0.0,
            }
        if is_inst:
            codigos[key]["soma_i"] += c["qtd"] * fator_i
        else:
            codigos[key]["soma_r"] += c["qtd"] * fator_r

    # ── 5. Montar resultado com uma linha por operação presente ──────────────
    resultado = []
    for key in codigos.keys():
        v = codigos[key]
        if v["soma_i"] > 0:
            resultado.append({"operacao": "I", "mdo": v["mdo"], "codigo": v["codigo"],
                               "desc_codigo": v["desc_codigo"], "filtro": v["filtro"], "total": round(v["soma_i"], 2)})
        if v["soma_r"] > 0:
            resultado.append({"operacao": "R", "mdo": v["mdo"], "codigo": v["codigo"],
                               "desc_codigo": v["desc_codigo"], "filtro": v["filtro"], "total": round(v["soma_r"], 2)})

    def get_mdo_rank(mdo_val):
        m = (mdo_val or "").strip().upper()
        # Vazio ou '-' = MAO-DE-OBRA (padrão), rank 1
        if m in ("", "-", "MAO-DE-OBRA", "MÃO-DE-OBRA", "MAO DE OBRA"):
            return 1
        elif m == "MATERIAL":
            return 2
        return 3

    # Ordenar: MDO explícito (MAO-DE-OBRA = 1 < MATERIAL = 2), Operação A-Z (I < R), Descrição A-Z
    resultado.sort(key=lambda x: (get_mdo_rank(x["mdo"]), x["operacao"], x["desc_codigo"].upper()))

    saida = {"resultado": resultado, "nao_encontrados": list(set(nao_encontrados))}
    if avisos:
        saida["avisos"] = list(dict.fromkeys(avisos))
    return saida
