"""
services/orcamento_calc.py — Lógica de cálculo do orçamento.
"""
import re


def processar_calculo(req_cabos: list, req_outros: list, req_projeto: str, orcamento_rows: list) -> dict:
    # ── 1. Índice ATIVO → linhas do orçamento (O(1) lookup com Fallback de Projeto) ──
    req_proj = (req_projeto or "").strip().upper()
    
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
        ativo_raw = item.get("ativo", "").strip().upper()
        operacao  = item.get("operacao", "").strip().upper()
        if not ativo_raw:
            continue
            
        txt = ativo_raw
        txt = txt.replace("CAA ", "CAA")
        txt = txt.replace("CA ", "CA")
        txt = txt.replace("CU ", "CU")
        txt = txt.replace("CAZ ", "CAZ")
        txt = txt.replace("P ", "P")
        
        # O 'm' final pode ser removido
        txt = re.sub(r'\s+m\s*$', '', txt, flags=re.IGNORECASE)
        
        parts = txt.split()
        if not parts:
            continue
            
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
            
        # Adiciona a margem de 5% sobre a quantidade processada (comprimento * multiplicador)
        qtd_final = (comprimento * multiplicador) * 1.05
        
        ativos_qtd.append({"ativo": nome_ativo, "qtd": qtd_final, "operacao": operacao})

    # 2b. OUTROS: texto separado por espaços no modelo "[qtd]-[ativo] [qtd]-[ativo] ..."
    for item in req_outros:
        ativo_raw = item.get("ativo", "").strip()
        operacao  = item.get("operacao", "").strip().upper()
        if not ativo_raw:
            continue
        txt = ativo_raw.strip()
        # Se inicia com DT ou CV sem quantidade, adiciona "1-"
        if re.match(r'^(DT|CV)', txt, re.IGNORECASE) and not re.match(r'^\d', txt):
            txt = "1-" + txt
        # "-" → " " (separa quantidade do ativo)
        txt = txt.replace("-", " ")
        # "*" → "-" (representa negativo / linha viva)
        txt = txt.replace("*", "-")
        # Posições pares = quantidades, posições ímpares = ativos
        tokens = txt.split()
        i = 0
        while i + 1 < len(tokens):
            try:
                qtd = float(tokens[i].replace(",", "."))
            except ValueError:
                qtd = 1.0
            nome_str = tokens[i + 1].upper()
            ativos_qtd.append({"ativo": nome_str, "qtd": qtd, "operacao": operacao})
            i += 2

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

    for av in ativos_qtd:
        nome_ativo = av["ativo"]
        linhas_ativo = orcamento_dict.get(nome_ativo, [])

        if not linhas_ativo:
            # ── Passo 2: busca por COMPONENTE exato ─────────────────────────
            if nome_ativo in componente_dict:
                linhas_ativo = componente_dict[nome_ativo]

        if not linhas_ativo:
            # ── Passo 3: busca por DESC_ATIVO exato ─────────────────────────
            if nome_ativo in desc_ativo_index:
                linhas_ativo = desc_ativo_index[nome_ativo]

        if not linhas_ativo:
            # ── Passo 4: busca parcial em DESC_ATIVO (nome_ativo contido na desc) ─
            matches = [rows for desc, rows in desc_ativo_index.items() if nome_ativo in desc]
            if matches:
                linhas_ativo = matches[0]  # Usa o primeiro match parcial

        if not linhas_ativo:
            # ── Passo 5: busca como CÓDIGO direto ───────────────────────────
            if nome_ativo in codigo_lookup:
                row_clone = dict(codigo_lookup[nome_ativo])
                row_clone["fator_i"] = 1.0
                row_clone["fator_r"] = 1.0
                componentes_qtd.append({
                    "qtd":      av["qtd"],
                    "operacao": av["operacao"],
                    "row":      row_clone,
                })
            else:
                nao_encontrados.append(nome_ativo)
            continue

        # Pega o COMPONENTE do primeiro match do ativo
        componente = linhas_ativo[0].get("componente", "").strip().upper()
        if not componente:
            continue

        # Agora busca TODOS os códigos desse componente
        # Para cada código, garante que a row escolhida tem mdo preenchido (se existir)
        linhas_componente = componente_dict.get(componente, [])
        
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

    # Ordenar: Descrição A-Z, depois Operação A-Z, depois MDO A-Z
    resultado.sort(key=lambda x: (x["desc_codigo"].upper(), x["operacao"], x["mdo"].upper()))

    return {"resultado": resultado, "nao_encontrados": list(set(nao_encontrados))}
