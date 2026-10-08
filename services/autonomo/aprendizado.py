"""
services/autonomo/aprendizado.py — Aprendizado supervisionado a partir das correções do usuário (TASK-059). Funções PURAS: sem banco, sem rede.

Fluxo: o autônomo entrega tabelas (Cabos/Outros); o usuário abre a obra, corrige e salva. `comparar` guarda só as DIFERENÇAS entre as duas
(eventos como "o item U4 virou U3", "a linha X foi excluída") mais contadores de quantas vezes cada coisa apareceu SEM ser corrigida.
`detectar_padroes` junta os eventos de várias obras e, quando um padrão se repete o bastante e quase sempre (consistência), propõe um
AJUSTE (receita do motor `ajustes_planilhas`). Nada é aplicado sozinho: toda proposta passa pela aprovação do administrador.

Privacidade: os eventos levam só códigos de ativo e operações — nunca coordenadas, textos do desenho ou nomes de arquivo.
"""
import difflib
import hashlib
import re
from collections import Counter

from services.ajustes_planilhas import validar_receitas
from services.orcamento_calc import extrair_ativos_outros

MINIMO_OCORRENCIAS = 3        # vezes que o padrão precisa aparecer
MINIMO_EXECUCOES = 2          # em quantas obras diferentes
CONSISTENCIA_MINIMA = 0.8     # ocorrências ÷ vezes em que a situação existia
IA_MINIMO_OBRAS = 10          # obras com correção registrada antes de oferecer a análise por IA
LIMITE_LINHAS_COMPARAR = 3000
LIMITE_EVENTOS = 1500
LIMITE_BASE = 3000
TIPOS_COM_RECEITA = ("item_trocado", "item_removido", "item_adicionado", "linha_excluida", "texto_trocado")
ROTULOS = {
    "item_trocado": "Troca de ativo", "item_removido": "Ativo removido", "item_adicionado": "Ativo adicionado", "linha_excluida": "Linha excluída",
    "texto_trocado": "Troca de texto no cabo", "operacao_alterada": "Operação alterada", "qtd_alterada": "Quantidade alterada",
    "linha_adicionada": "Linha adicionada", "linha_editada": "Linha editada",
}
_RE_NUM = re.compile(r"^[\d.,]+$")


def _norm(texto) -> str:
    return re.sub(r"\s+", " ", str(texto or "").strip().upper())


def _linhas(valor) -> list:
    dados = valor.get("data") if isinstance(valor, dict) else valor
    return [l for l in (dados if isinstance(dados, list) else []) if isinstance(l, dict)]


def _itens(ativo) -> list:
    """[(código, qtd)] de uma linha de Outros (mesma leitura do cálculo)."""
    return [(i["ativo"], i["qtd"]) for i in extrair_ativos_outros({"ativo": ativo or "", "operacao": ""})]


def _fmt_qtd(q) -> str:
    return str(int(q)) if float(q).is_integer() else str(round(q, 4))


def _contar_base(tab, linhas, base):
    for l in linhas:
        a = _norm(l.get("ativo"))
        if not a:
            continue
        base["linha"][f"{tab}|{a}"] += 1
        if tab == "outros":
            for cod in {c for c, _ in _itens(a)}:
                base["item"][cod] += 1
        else:
            for tok in {t for t in a.split() if not _RE_NUM.match(t)}:
                base["token"][tok] += 1


def comparar(agente: dict, usuario: dict) -> dict:
    """Diferenças entre as tabelas do autônomo (`agente`) e as salvas pelo usuário. Cada tabela: lista de linhas ou {data: [...]}.

    Devolve {"eventos": [...], "base": {"linha": {}, "item": {}, "token": {}}, "linhas": {"agente": n, "usuario": n}, "ignoradas": [...]}.
    """
    base = {"linha": Counter(), "item": Counter(), "token": Counter()}
    eventos, ignoradas, n_ag, n_us = [], [], 0, 0
    for tab in ("cabos", "outros"):
        A, B = _linhas((agente or {}).get(tab)), _linhas((usuario or {}).get(tab))
        n_ag += len(A); n_us += len(B)
        _contar_base(tab, A, base)
        if len(A) > LIMITE_LINHAS_COMPARAR or len(B) > LIMITE_LINHAS_COMPARAR:
            ignoradas.append(f"{tab}: tabela grande demais para comparar")
            continue
        ka = [(_norm(l.get("operacao")), _norm(l.get("ativo"))) for l in A]
        kb = [(_norm(l.get("operacao")), _norm(l.get("ativo"))) for l in B]
        excluidas, adicionadas = [], []
        for tag, i1, i2, j1, j2 in difflib.SequenceMatcher(None, ka, kb, autojunk=False).get_opcodes():
            if tag == "equal":
                continue
            pares = _parear(ka, kb, range(i1, i2), range(j1, j2)) if tag == "replace" else []
            for i, j in pares:
                eventos += _evento_par(tab, A[i], B[j])
            excluidas += [i for i in range(i1, i2) if i not in {p[0] for p in pares}]
            adicionadas += [j for j in range(j1, j2) if j not in {p[1] for p in pares}]
        # linha só mudou de lugar (mesma operação e ativo excluída aqui e inserida ali) não é correção
        mov = Counter(ka[i] for i in excluidas) & Counter(kb[j] for j in adicionadas)
        for i in excluidas:
            if mov[ka[i]] > 0:
                mov[ka[i]] -= 1
                continue
            eventos.append({"tipo": "linha_excluida", "tabela": tab, "ativo": ka[i][1], "operacao": ka[i][0]})
        mov = Counter(ka[i] for i in excluidas) & Counter(kb[j] for j in adicionadas)
        for j in adicionadas:
            if mov[kb[j]] > 0:
                mov[kb[j]] -= 1
                continue
            eventos.append({"tipo": "linha_adicionada", "tabela": tab, "ativo": kb[j][1], "operacao": kb[j][0]})
    return {"eventos": eventos[:LIMITE_EVENTOS], "base": {k: dict(v.most_common(LIMITE_BASE)) for k, v in base.items()},
            "linhas": {"agente": n_ag, "usuario": n_us}, "ignoradas": ignoradas}


LIMIAR_PAREAMENTO = 0.5


def _parear(ka, kb, ra, rb) -> list:
    """Em um trecho que mudou, casa cada linha do autônomo com a linha do usuário MAIS PARECIDA (texto do ativo; parecença >= 0,5), em vez de
    casar por posição: assim 'linha excluída' e 'linha editada' não se confundem. Blocos enormes caem no casamento por posição."""
    ra, rb = list(ra), list(rb)
    if len(ra) * len(rb) > 40000:
        return list(zip(ra, rb))
    cand = []
    for i in ra:
        for j in rb:
            r = 1.0 if ka[i][1] == kb[j][1] else difflib.SequenceMatcher(None, ka[i][1], kb[j][1], autojunk=False).ratio()
            if r >= LIMIAR_PAREAMENTO:
                cand.append((-r, abs((i - ra[0]) - (j - rb[0])), i, j))
    cand.sort()
    usados_a, usados_b, pares = set(), set(), []
    for _, _, i, j in cand:
        if i not in usados_a and j not in usados_b:
            usados_a.add(i); usados_b.add(j); pares.append((i, j))
    return sorted(pares)


def _evento_par(tab, a: dict, b: dict) -> list:
    """Eventos entre duas linhas que se correspondem (a do autônomo, b do usuário)."""
    ev = []
    oa, ob = _norm(a.get("operacao")), _norm(b.get("operacao"))
    aa, ab = _norm(a.get("ativo")), _norm(b.get("ativo"))
    if oa != ob:
        ev.append({"tipo": "operacao_alterada", "tabela": tab, "ativo": aa, "de": oa, "para": ob})
    if aa == ab:
        return ev
    if tab == "outros":
        ia, ib = _itens(aa), _itens(ab)
        ca, cb = Counter(c for c, _ in ia), Counter(c for c, _ in ib)
        removidos, adicionados = sorted((ca - cb).elements()), sorted((cb - ca).elements())
        comuns = sorted(set(ca) & set(cb))
        if len(removidos) == 1 and len(adicionados) == 1:
            ev.append({"tipo": "item_trocado", "tabela": tab, "de": removidos[0], "para": adicionados[0]})
        else:
            for c in removidos:
                ev.append({"tipo": "item_removido", "tabela": tab, "ativo": c})
            qb = {c: q for c, q in ib}
            for c in adicionados:
                for ctx in comuns[:3]:
                    ev.append({"tipo": "item_adicionado", "tabela": tab, "ativo": c, "qtd": _fmt_qtd(qb[c]), "contexto": ctx})
        qa, qb = {c: q for c, q in ia}, {c: q for c, q in ib}
        for c in comuns:
            if qa[c] != qb[c]:
                ev.append({"tipo": "qtd_alterada", "tabela": tab, "ativo": c, "de": _fmt_qtd(qa[c]), "para": _fmt_qtd(qb[c])})
        if not (removidos or adicionados) and not any(e["tipo"] == "qtd_alterada" for e in ev):
            ev.append({"tipo": "linha_editada", "tabela": tab, "ativo": aa})
        return ev
    ta, tb = aa.split(), ab.split()
    ops = [o for o in difflib.SequenceMatcher(None, ta, tb, autojunk=False).get_opcodes() if o[0] != "equal"]
    if len(ops) == 1 and ops[0][0] == "replace" and ops[0][2] - ops[0][1] == 1 and ops[0][4] - ops[0][3] == 1 \
            and not _RE_NUM.match(ta[ops[0][1]]) and not _RE_NUM.match(tb[ops[0][3]]):
        ev.append({"tipo": "texto_trocado", "tabela": tab, "de": ta[ops[0][1]], "para": tb[ops[0][3]]})
    else:
        ev.append({"tipo": "linha_editada", "tabela": tab, "ativo": aa})
    return ev


# ── padrões e propostas ─────────────────────────────────────────────────────
def _chave_evento(e: dict) -> tuple:
    t = e["tipo"]
    if t == "item_trocado" or t == "texto_trocado":
        return (t, e["tabela"], e["de"], e["para"])
    if t == "item_removido":
        return (t, e["tabela"], e["ativo"])
    if t == "item_adicionado":
        return (t, e["tabela"], e["ativo"], e["qtd"], e["contexto"])
    if t in ("linha_excluida", "linha_adicionada", "linha_editada"):
        return (t, e["tabela"], e["ativo"])
    if t == "operacao_alterada":
        return (t, e["tabela"], e["ativo"], e["de"], e["para"])
    return (t, e["tabela"], e["ativo"], e["de"], e["para"])        # qtd_alterada


def _base_do_evento(e: dict, base: dict) -> int:
    t = e["tipo"]
    if t == "item_trocado":
        return base.get("item", {}).get(e["de"], 0)
    if t == "item_removido":
        return base.get("item", {}).get(e["ativo"], 0)
    if t == "item_adicionado":
        return base.get("item", {}).get(e["contexto"], 0)
    if t == "texto_trocado":
        return base.get("token", {}).get(e["de"], 0)
    if t in ("qtd_alterada",):
        return base.get("item", {}).get(e["ativo"], 0)
    return base.get("linha", {}).get(f"{e['tabela']}|{e['ativo']}", 0)


def descrever(e: dict) -> str:
    t, tab = e["tipo"], {"cabos": "Cabos", "outros": "Outros"}[e["tabela"]]
    if t == "item_trocado":
        return f"Em {tab}, você trocou o ativo {e['de']} por {e['para']}."
    if t == "texto_trocado":
        return f"Em {tab}, você trocou '{e['de']}' por '{e['para']}' no texto do ativo."
    if t == "item_removido":
        return f"Em {tab}, você removeu o ativo {e['ativo']} das linhas."
    if t == "item_adicionado":
        return f"Em {tab}, você adicionou {e['qtd']}-{e['ativo']} nas linhas que têm {e['contexto']}."
    if t == "linha_excluida":
        return f"Em {tab}, você excluiu a linha '{e['ativo']}'."
    if t == "linha_adicionada":
        return f"Em {tab}, você adicionou a linha '{e['ativo']}' (operação {e['operacao']})."
    if t == "operacao_alterada":
        return f"Em {tab}, você mudou a operação de '{e['ativo']}' de {e['de'] or '(vazia)'} para {e['para'] or '(vazia)'}."
    if t == "qtd_alterada":
        return f"Em {tab}, você mudou a quantidade de {e['ativo']} de {e['de']} para {e['para']}."
    return f"Em {tab}, você editou a linha '{e['ativo']}'."


def acao_do_evento(e: dict):
    """Ação do motor de ajustes que reproduz a correção, ou None quando o tipo é só informativo."""
    t, tab = e["tipo"], e["tabela"]
    if t == "item_trocado":
        return {"acao": "substituir", "tabela": tab, "de": e["de"], "para": e["para"], "modo": "item"}
    if t == "texto_trocado":
        return {"acao": "substituir", "tabela": tab, "de": e["de"], "para": e["para"], "modo": "texto", "palavra_inteira": True}
    if t == "item_removido":
        return {"acao": "remover_ativo", "tabela": tab, "ativo": e["ativo"]}
    if t == "item_adicionado":
        return {"acao": "adicionar_ativo", "tabela": tab, "ativo": e["ativo"], "qtd": float(e["qtd"]) if not float(e["qtd"]).is_integer() else int(float(e["qtd"])),
                "se_ja_existe": "ignorar", "quando": {"tem": e["contexto"]}}
    if t == "linha_excluida":
        return {"acao": "excluir_linhas", "tabela": tab, "onde": {"texto": "^" + re.escape(e["ativo"]) + "$"}}
    return None


def id_da_chave(chave: tuple) -> str:
    return hashlib.sha1("|".join(chave).encode("utf-8")).hexdigest()[:10]


def receita_da_acao(acao: dict, rotulo: str, descricao: str, sufixo: str) -> dict:
    return {"id": f"apr_{sufixo}", "nome": f"Aprendido: {rotulo}"[:120], "ativa": True, "descricao": descricao[:500], "acoes": [acao]}


def detectar_padroes(registros: list, grupos=None, minimo=MINIMO_OCORRENCIAS, minimo_execucoes=MINIMO_EXECUCOES, consistencia_minima=CONSISTENCIA_MINIMA) -> list:
    """`registros` = [{"execucao_id", "eventos", "base"}] (uma por obra corrigida). Devolve os padrões que passam nos limiares, mais fortes primeiro.

    Cada padrão: {chave, id, tipo, tabela, descricao, ocorrencias, execucoes, base, consistencia, receita|None}. `receita` só existe quando o motor de ajustes
    aceita a ação (validada aqui); sem ela o padrão é informativo ("considere criar uma regra de classificação")."""
    agregado = {}
    for reg in registros:
        for e in reg.get("eventos") or []:
            try:
                k = _chave_evento(e)
            except KeyError:
                continue
            a = agregado.setdefault(k, {"evento": e, "n": 0, "execs": set(), "base": 0})
            a["n"] += 1
            a["execs"].add(reg.get("execucao_id"))
    for a in agregado.values():        # a base soma TODAS as obras (inclusive as em que o usuário não corrigiu aquilo)
        a["base"] = sum(_base_do_evento(a["evento"], reg.get("base") or {}) for reg in registros)
    padroes = []
    for k, a in agregado.items():
        if a["n"] < minimo or len(a["execs"]) < minimo_execucoes:
            continue
        consistencia = a["n"] / max(a["base"], a["n"])
        if consistencia < consistencia_minima:
            continue
        e = a["evento"]
        acao = acao_do_evento(e)
        receita = None
        if acao is not None:
            sufixo = id_da_chave(k)
            texto = f"{descrever(e)} Visto {a['n']} vez(es) em {len(a['execs'])} obra(s), em {round(consistencia * 100)}% dos casos."
            r = receita_da_acao(acao, descrever(e).rstrip("."), texto, sufixo)
            if not validar_receitas([r], grupos or {}):
                receita = r
        padroes.append({"chave": "|".join(k), "id": id_da_chave(k), "tipo": e["tipo"], "rotulo": ROTULOS[e["tipo"]], "tabela": e["tabela"],
                        "descricao": descrever(e), "ocorrencias": a["n"], "execucoes": len(a["execs"]), "base": a["base"],
                        "consistencia": round(consistencia, 3), "receita": receita})
    padroes.sort(key=lambda p: (-p["ocorrencias"], -p["consistencia"], p["chave"]))
    return padroes


def resumo_para_ia(padroes: list, registros: list, limite=40) -> str:
    """Texto compacto (só códigos de ativo e contagens) para pedir sugestões à IA."""
    linhas = [f"Foram analisadas {len(registros)} obras corrigidas pelo usuário depois do processamento automático.",
              "Correções observadas (tipo | descrição | vezes | obras | consistência):"]
    for p in padroes[:limite]:
        linhas.append(f"- {p['tipo']} | {p['descricao']} | {p['ocorrencias']} | {p['execucoes']} | {round(p['consistencia'] * 100)}%")
    return "\n".join(linhas)
