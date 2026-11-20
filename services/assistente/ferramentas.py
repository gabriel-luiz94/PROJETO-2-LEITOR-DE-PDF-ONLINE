"""
services/assistente/ferramentas.py (TASK-191) — As ferramentas do assistente de IA: o que ele pode CONSULTAR, CALCULAR e PROPOR no programa.

Regra de ouro: nada aqui altera a tabela do usuário. A tabela mora no navegador; o servidor recebe uma cópia a cada pedido (`Contexto.cabos/outros`).
Consultas e cálculos respondem na hora; as mudanças viram um PLANO (`Contexto.plano`) que o navegador mostra e só aplica depois do clique em «Aplicar»
(com desfazer). Cada ferramenta devolve um dict JSON-seguro; erro vira {"erro": "..."} para o modelo corrigir a chamada em vez de derrubar a conversa.
"""
import json
import math
import re
from dataclasses import dataclass, field
from typing import Any, Callable, Optional

from services.assistente import catalogo, skills

LIMITE_LINHAS = 300            # linhas devolvidas por consulta (ler_tabela / totalizar)
LIMITE_PLANO = 200             # alterações num plano
LIMITE_TEXTO = 300             # tamanho de um ativo digitado
OPERACOES = ("I", "R", "*I", "*R", "M", "*M")
_RE_COMPRIMENTO_FIM = re.compile(r"\d+(?:[.,]\d+)?\s*M?\s*$", re.IGNORECASE)
_RE_TERMINA_EM_M = re.compile(r"\d+[.,]?\d*\s*M\s*$", re.IGNORECASE)


@dataclass
class Contexto:
    """Tudo que as ferramentas enxergam de UM pedido. Os carregadores são injetáveis (testes); em produção leem o banco só quando uma ferramenta pede."""
    projeto_codigo: str = ""
    projeto_nome: str = ""
    user: dict = field(default_factory=dict)
    cabos: list = field(default_factory=list)          # [{id, operacao, ativo, entidade, ...}] como está na tela (sem as linhas vazias)
    outros: list = field(default_factory=list)
    selecao_mapa: Optional[dict] = None
    tabela_incompleta: bool = False
    obra: Optional[dict] = None                        # parâmetros OBRA.* da obra aberta (município, UTD…)
    plano: list = field(default_factory=list)          # alterações PROPOSTAS (nunca aplicadas aqui)
    resumo_plano: str = ""
    acoes_ui: list = field(default_factory=list)       # ordenar/filtrar a tela (sem confirmação: não altera dados)
    carregar_pipeline: Optional[Callable[[], Any]] = None
    carregar_obras: Optional[Callable[[], list]] = None
    buscar_obra: Optional[Callable[[str], Optional[dict]]] = None
    _pipe: Any = None
    _indice: Optional[dict] = None

    @property
    def slug(self) -> str:
        return skills.slug(self.projeto_nome)

    def pipe(self):
        """O contexto do pipeline (regras de conversão + base de orçamento do usuário) — carregado uma vez."""
        if self._pipe is None:
            if self.carregar_pipeline is None:
                from services.autonomo.pipeline import carregar_contexto
                self._pipe = carregar_contexto(self.projeto_codigo, (self.user or {}).get("user_id"))
            else:
                self._pipe = self.carregar_pipeline()
        return self._pipe

    def base(self) -> list:
        return self.pipe().base_orcamento or []

    def indice(self) -> dict:
        if self._indice is None:
            self._indice = catalogo.indexar(self.base(), self.projeto_nome, self.projeto_codigo)
        return self._indice


# ── utilidades ────────────────────────────────────────────────────────────────────────────────────────
def _json_seguro(v):
    """nan/inf → None (o Totalizador devolve NaN como o JS); sets → listas."""
    if isinstance(v, float):
        return None if (math.isnan(v) or math.isinf(v)) else v
    if isinstance(v, dict):
        return {str(k): _json_seguro(x) for k, x in v.items()}
    if isinstance(v, (list, tuple, set)):
        return [_json_seguro(x) for x in v]
    return v


def _texto(v) -> str:
    return str(v).strip() if v is not None else ""


def _inteiro(v):
    try:
        if isinstance(v, bool):
            return None
        f = float(v)
        return int(f) if f == int(f) else None
    except (TypeError, ValueError):
        return None


def _linhas_da_tela(ctx: Contexto, tabela: str) -> list:
    return [r for r in (ctx.cabos if tabela == "cabos" else ctx.outros) if isinstance(r, dict) and not r.get("congelada")]


def _tabela(v, padrao=None) -> Optional[str]:
    t = _texto(v).lower()
    return t if t in ("cabos", "outros") else padrao


def _operacao(v, padrao="I") -> Optional[str]:
    o = _texto(v).upper().replace(" ", "")
    return (o if o in OPERACOES else None) if o else padrao


def aplicar_plano(cabos: list, outros: list, plano: list):
    """Simula sobre uma cópia e mantém os vínculos posicionais compatíveis com a tabela resultante."""
    from services.vinculacao_regras import id_unidade, linha_de, slot_de

    tabelas = {
        nome: [dict(r, id=r.get("id", i)) for i, r in enumerate(linhas) if isinstance(r, dict)]
        for nome, linhas in (("cabos", cabos), ("outros", outros))
    }
    por_id = {nome: {r["id"]: r for r in linhas} for nome, linhas in tabelas.items()}
    removidas = {"cabos": set(), "outros": set()}
    for it in plano:
        acao = it.get("acao")
        if acao not in ("adicionar", "editar", "remover"):
            continue
        tabela = it["tabela"]
        if acao == "adicionar":
            tabelas[tabela].append({"id": None, "entidade": "CABO" if tabela == "cabos" else "0", "operacao": it["operacao"], "ativo": it["ativo"]})
        else:
            linha = por_id[tabela][it["id"]]
            if acao == "remover":
                removidas[tabela].add(it["id"])
            else:
                linha.update(operacao=it["operacao"], ativo=it["ativo"])
    cabos = [r for r in tabelas["cabos"] if r["id"] not in removidas["cabos"]]
    outros = [r for r in tabelas["outros"] if r["id"] not in removidas["outros"]]
    mapa = {r["id"]: i for i, r in enumerate(outros) if r["id"] is not None}
    for cabo in cabos:
        if isinstance(cabo.get("vinculoPostes"), list):
            cabo["vinculoPostes"] = [mapa[i] for i in cabo["vinculoPostes"] if i in mapa]
        if isinstance(cabo.get("vinculoEstruturas"), list):
            cabo["vinculoEstruturas"] = [
                id_unidade(mapa[linha_de(i)], slot_de(i)) for i in cabo["vinculoEstruturas"]
                if isinstance(i, (int, float)) and math.isfinite(i) and linha_de(i) in mapa
            ]
    return cabos, outros


def _totalizar_bruto(ctx: Contexto, cabos: list, outros: list) -> list:
    """A mesma Totalizadora da tela (regras de conversão do projeto + vínculos cabo↔estrutura que vieram nas linhas)."""
    from services.autonomo.totalizadora import montar_totalizadora
    from services.autonomo.pipeline import _cfg_vinculo
    from services.autonomo.vinculacao import itens_vinculo_para_totalizadora

    p = ctx.pipe()
    itens_vinculo, cfg = itens_vinculo_para_totalizadora(cabos, outros), _cfg_vinculo(p)
    return montar_totalizadora(cabos, outros, p.regras_conversao, p.base_orcamento, p.projeto_nome or ctx.projeto_nome,
                               itens_vinculo=itens_vinculo, cfg_vinculo=cfg, obra=ctx.obra)


def _tabelas_para_calculo(ctx: Contexto, com_plano: bool):
    if ctx.tabela_incompleta:
        raise ValueError("A cópia das tabelas está incompleta. Calcule a obra completa pelo Resumo.")
    return aplicar_plano(ctx.cabos, ctx.outros, ctx.plano if com_plano else [])


# ── ferramentas ───────────────────────────────────────────────────────────────────────────────────────
def listar_skills(ctx: Contexto, a: dict) -> dict:
    return {"skills": skills.listar(ctx.slug, ctx.projeto_codigo)}


def carregar_skill(ctx: Contexto, a: dict) -> dict:
    nome = _texto(a.get("nome"))
    try:
        return {"nome": nome, "conteudo": skills.carregar(ctx.slug, nome, ctx.projeto_codigo)}
    except skills.ErroSkill as e:
        return {"erro": str(e), "disponiveis": [s["nome"] for s in skills.listar(ctx.slug, ctx.projeto_codigo)]}


def buscar_ativo(ctx: Contexto, a: dict) -> dict:
    texto = _texto(a.get("texto"))
    if not texto:
        return {"erro": "Informe o texto a buscar (código ou parte da descrição)."}
    limite = min(max(_inteiro(a.get("limite")) or 15, 1), 40)
    achados = catalogo.buscar(ctx.indice(), texto, limite)
    return {"projeto": ctx.projeto_nome, "total": len(achados),
            "ativos": [{"ativo": e["nome"], "descricao": e["descricao"], "componentes": e["componentes"], "formas": e["formas"]} for e in achados],
            "aviso": "" if achados else "Nenhum ativo da base casa com isso. Não invente: pergunte ao usuário ou tente outras palavras."}


def detalhar_ativo(ctx: Contexto, a: dict) -> dict:
    nome = _texto(a.get("ativo"))
    if not nome:
        return {"erro": "Informe o código do ativo."}
    comps = catalogo.componentes_de(ctx.base(), nome, ctx.projeto_nome, ctx.projeto_codigo)
    if not comps:
        parecidos = [e["nome"] for e in catalogo.buscar(ctx.indice(), nome, 5)]
        return {"erro": f"Ativo '{nome}' não está na base do projeto.", "parecidos": parecidos}
    return {"ativo": nome.upper(), "componentes": _json_seguro(comps)}


def ler_tabela(ctx: Contexto, a: dict) -> dict:
    tabela = _texto(a.get("tabela")).lower() or "ambas"
    alvo = ("cabos", "outros") if tabela == "ambas" else ((tabela,) if tabela in ("cabos", "outros") else None)
    if alvo is None:
        return {"erro": "tabela deve ser 'cabos', 'outros' ou 'ambas'."}
    limite = min(max(_inteiro(a.get("limite")) or LIMITE_LINHAS, 1), LIMITE_LINHAS)
    inicio = max(_inteiro(a.get("inicio")) or 0, 0)
    texto = _texto(a.get("texto")).casefold()
    saida, restante = {}, limite
    for t in alvo:
        linhas = [r for r in _linhas_da_tela(ctx, t) if texto in _texto(r.get("ativo")).casefold()]
        # Ao consultar ambas, mantenha uma cota para cada tabela e um cursor por resultado.
        cota = max(1, restante // 2) if t == "cabos" and len(alvo) == 2 else restante
        pagina = linhas[inicio:inicio + cota]
        mostradas = [{"id": r.get("id"), "operacao": r.get("operacao") or "", "ativo": r.get("ativo") or "",
                      **({"bloqueada": True} if r.get("bloqueada") else {})} for r in pagina]
        restante -= len(mostradas)
        proximo = inicio + len(mostradas)
        saida[t] = {"total": len(linhas), "linhas": mostradas, "cortado": proximo < len(linhas),
                    "proximo_inicio": proximo if proximo < len(linhas) else None}
        if ctx.tabela_incompleta:
            saida[t]["copia_incompleta"] = True
    return saida


def consultar_vinculos(ctx: Contexto, a: dict) -> dict:
    """Explica a seleção do mapa com a configuração e o analisador usados pelo projeto."""
    from services.autonomo.montagem import prefixo_cabo
    from services.autonomo.pipeline import _cfg_vinculo
    from services.autonomo.vinculacao import codigo_estrutura, codigos_estrutura
    from services.vinculacao_regras import analisar, id_unidade, linha_de, slot_de

    selecao = ctx.selecao_mapa
    if not selecao and a.get("cabo_id") is None and a.get("poste_id") is None:
        return {"erro": "Abra o assistente pelo poste ou vão no mapa, ou informe o ID de uma linha."}
    cabos, outros = _tabelas_para_calculo(ctx, False)
    ids_cabos = {a["cabo_id"]} if a.get("cabo_id") is not None else set((selecao or {}).get("cabos", []))
    ids_outros = {a["poste_id"]} if a.get("poste_id") is not None else set((selecao or {}).get("outros", []))
    if a.get("cabo_id") is not None:
        ids_outros = set()
    if a.get("poste_id") is not None:
        ids_cabos = set()
    cabos_alvo = {i for i, r in enumerate(cabos) if r["id"] in ids_cabos}
    outros_alvo = {i for i, r in enumerate(outros) if r["id"] in ids_outros}
    outros_alvo.update(linha_de(e) for i in cabos_alvo for e in cabos[i].get("vinculoEstruturas", []))
    analise = analisar(cabos, outros, _cfg_vinculo(ctx.pipe()), prefixo_cabo, codigo_estrutura, codigos_estrutura)
    links = [l for l in analise["links"] if l["cabo"] in cabos_alvo or linha_de(l["estrutura"]) in outros_alvo]
    cabos_alvo.update(l["cabo"] for l in links)
    if not cabos_alvo and not outros_alvo:
        return {"erro": "Não encontrei as linhas da seleção na obra atual."}

    def id_original(ie):
        return id_unidade(outros[linha_de(ie)]["id"], slot_de(ie))

    return {"projeto": ctx.projeto_codigo, "selecao": (selecao or {}).get("rotulo"),
            "cabos": [{"id": cabos[i]["id"], "ativo": cabos[i].get("ativo") or "", "operacao": cabos[i].get("operacao"),
                       "vinculacao": analise["cabos"].get(i)} for i in sorted(cabos_alvo)],
            "estruturas": [{"id": id_original(i), **r} for i, r in analise["estruturas"].items() if linha_de(i) in outros_alvo],
            "vinculos": [{**l, "cabo": cabos[l["cabo"]]["id"], "estrutura": id_original(l["estrutura"])} for l in links]}


def validar_linhas(ctx: Contexto, a: dict) -> dict:
    """Roda a Totalizadora só nas linhas dadas e diz quais ativos a base não reconhece (sem tocar na tabela nem no plano)."""
    brutas = a.get("linhas")
    if not isinstance(brutas, list) or not brutas:
        return {"erro": "Envie 'linhas': [{tabela, ativo, operacao}]."}
    cabos, outros = [], []
    problemas = []
    for l in brutas[:60]:
        if not isinstance(l, dict) or not _texto(l.get("ativo")):
            continue
        t = _tabela(l.get("tabela"), "outros")
        operacao = _operacao(l.get("operacao"))
        if operacao is None:
            problemas.append({"ativo": _texto(l.get("ativo")), "motivo": "operação inválida"})
            continue
        (cabos if t == "cabos" else outros).append({"entidade": "CABO" if t == "cabos" else "0", "operacao": operacao, "ativo": _texto(l.get("ativo"))})
    tot = _totalizar_bruto(ctx, cabos, outros)
    problemas += [{"ativo": r.get("ativo"), "motivo": r.get("desc") or "não encontrado na base"} for r in tot if r.get("naoEncontrado")]
    return {"linhas_validadas": len(cabos) + len(outros), "problemas": _json_seguro(problemas), "tudo_certo": not problemas}


def _avisos_do_item(tabela: str, ativo: str) -> list:
    av = []
    if tabela == "cabos" and not _RE_COMPRIMENTO_FIM.search(ativo):
        av.append("linha de Cabos costuma terminar no comprimento (ex.: 'C16 3 35 m' ou '3xC16(16) 35 M')")
    if tabela == "outros" and _RE_TERMINA_EM_M.search(ativo):
        av.append("termina em 'm' (formato de cabo) — isso é da tabela Cabos?")
    return av


def propor_alteracoes(ctx: Contexto, a: dict) -> dict:
    """Registra no PLANO as mudanças pedidas (adicionar/editar/remover). Valida ids e formato; o usuário confirma na tela."""
    from services.autonomo.vinculacao import codigos_estrutura
    from services.vinculacao_regras import linha_de, slot_de

    itens = a.get("alteracoes")
    if not isinstance(itens, list) or not itens:
        return {"erro": "Envie 'alteracoes': lista de {acao, tabela, id?, operacao?, ativo?}."}
    if ctx.tabela_incompleta:
        return {"erro": "A cópia das tabelas está incompleta. Edite pelo Resumo."}
    if any(p.get("acao") == "obra" for p in ctx.plano):
        return {"erro": "Já há um comando de obra no plano; não misture obra com outras alterações. Peça ao usuário para confirmar a obra antes."}
    if a.get("substituir"):
        ctx.plano.clear()
    ids = {t: {r.get("id"): r for r in _linhas_da_tela(ctx, t)} for t in ("cabos", "outros")}
    ja = {(p["tabela"], p["id"]) for p in ctx.plano if p.get("id") is not None}
    aceitas, rejeitadas, avisos, novas = 0, [], [], []
    ligacoes = {}
    for cabo in ctx.cabos:
        for unidade in cabo.get("vinculoEstruturas") or []:
            if isinstance(unidade, (int, float)) and math.isfinite(unidade):
                ligacoes.setdefault(linha_de(unidade), []).append(unidade)
    postes = {}
    for cabo in ctx.cabos:
        for linha in cabo.get("vinculoPostes") or []:
            postes[linha] = postes.get(linha, 0) + 1
    for n, it in enumerate(itens, 1):
        if len(ctx.plano) + len(novas) >= LIMITE_PLANO:
            rejeitadas.append({"n": n, "motivo": f"o plano comporta até {LIMITE_PLANO} alterações; proponha o resto depois"})
            break
        if not isinstance(it, dict):
            rejeitadas.append({"n": n, "motivo": "item inválido"})
            continue
        acao, tabela = _texto(it.get("acao")).lower(), _tabela(it.get("tabela"))
        if acao not in ("adicionar", "editar", "remover") or not tabela:
            rejeitadas.append({"n": n, "motivo": "informe acao (adicionar|editar|remover) e tabela (cabos|outros)"})
            continue
        ativo = _texto(it.get("ativo"))
        if len(ativo) > LIMITE_TEXTO:
            rejeitadas.append({"n": n, "motivo": "ativo grande demais"})
            continue
        item = {"n": 0, "acao": acao, "tabela": tabela, "id": None, "operacao": _operacao(it.get("operacao")), "ativo": ativo, "antes": None, "avisos": []}
        if acao in ("editar", "remover"):
            i = _inteiro(it.get("id"))
            if i is None or i not in ids[tabela]:
                rejeitadas.append({"n": n, "motivo": f"id {it.get('id')!r} não existe na tabela {tabela} (use ler_tabela)"})
                continue
            if (tabela, i) in ja:
                rejeitadas.append({"n": n, "motivo": f"a linha {tabela}-{i} já está no plano"})
                continue
            ant = ids[tabela][i]
            if ant.get("bloqueada"):
                rejeitadas.append({"n": n, "motivo": "linha bloqueada"})
                continue
            if acao == "editar" and not _texto(it.get("operacao")):
                item["operacao"] = ant.get("operacao") or "I"
            item["id"], item["antes"] = i, {"operacao": ant.get("operacao") or "", "ativo": ant.get("ativo") or ""}
            if acao == "remover":
                item["ativo"], item["operacao"] = ant.get("ativo") or "", ant.get("operacao") or ""
                estruturas = len(ligacoes.get(i, [])) if tabela == "outros" else len(ant.get("vinculoEstruturas") or [])
                atribuicoes = postes.get(i, 0) if tabela == "outros" else len(ant.get("vinculoPostes") or [])
                if estruturas or atribuicoes:
                    item["avisos"].append(f"A remoção desfaz {estruturas} vínculo(s) com estruturas e {atribuicoes} atribuição(ões) de poste.")
            elif tabela == "outros" and i in ligacoes:
                antes = codigos_estrutura(ant.get("ativo") or "")
                depois = codigos_estrutura(ativo)
                slots = {slot_de(unidade) for unidade in ligacoes[i]}
                if any(antes[k:k + 1] != depois[k:k + 1] for k in slots):
                    rejeitadas.append({"n": n, "motivo": "a edição trocaria uma estrutura já vinculada. Revise os vínculos no mapa."})
                    continue
        if acao in ("adicionar", "editar"):
            if item["operacao"] is None:
                rejeitadas.append({"n": n, "motivo": "operação inválida"})
                continue
            if ativo.startswith(("@", "#")):
                rejeitadas.append({"n": n, "motivo": "use o Resumo para criar comentários ou grupos"})
                continue
            if not ativo:
                rejeitadas.append({"n": n, "motivo": "ativo vazio"})
                continue
            item["avisos"] += _avisos_do_item(tabela, ativo)
        if item["id"] is not None:
            ja.add((tabela, item["id"]))
        novas.append(item)
    # confere na base os ativos que entram (uma só Totalizadora para todos)
    entram = [x for x in novas if x["acao"] in ("adicionar", "editar")]
    if entram:
        try:
            cabos = [{"entidade": "CABO", "operacao": x["operacao"], "ativo": x["ativo"]} for x in entram if x["tabela"] == "cabos"]
            outros = [{"entidade": "0", "operacao": x["operacao"], "ativo": x["ativo"]} for x in entram if x["tabela"] == "outros"]
            nao = {str(r.get("ativo") or "").upper(): (r.get("desc") or "") for r in _totalizar_bruto(ctx, cabos, outros) if r.get("naoEncontrado")}
            for x in entram:
                motivos = [m for n, m in nao.items() if n and n in x["ativo"].upper()]
                if motivos:
                    x["avisos"].append(motivos[0] or "ativo não encontrado na base do projeto (confira com buscar_ativo)")
        except Exception:                              # a conferência é um extra: sem ela o plano segue
            avisos.append("Não foi possível conferir os ativos na base do projeto.")
            for x in entram:
                x["avisos"].append(avisos[-1])
    for x in novas:
        ctx.plano.append(x)
        x["n"] = len(ctx.plano)
        aceitas += 1
        avisos += [f"#{x['n']} {x['ativo']}: {m}" for m in x["avisos"]]
    if a.get("resumo"):
        ctx.resumo_plano = _texto(a["resumo"])[:300]
    return {"aceitas": aceitas, "rejeitadas": rejeitadas, "avisos": avisos, "total_no_plano": len(ctx.plano),
            "proximo_passo": "O usuário verá o plano e confirmará na tela. Explique em 1-2 frases o que propôs; não diga que já aplicou."}


def totalizar(ctx: Contexto, a: dict) -> dict:
    cabos, outros = _tabelas_para_calculo(ctx, bool(a.get("com_plano")))
    tot = _totalizar_bruto(ctx, cabos, outros)
    agreg = {}
    for r in tot:
        chave = (r.get("operacao") or "I", r.get("ativo") or "", r.get("origem") or "")
        e = agreg.setdefault(chave, {"operacao": chave[0], "ativo": chave[1], "origem": chave[2], "qtd": 0.0, "linhas": 0, "desc": r.get("desc") or "", "nao_encontrado": False})
        q = r.get("qtd")
        if isinstance(q, (int, float)) and not (isinstance(q, float) and math.isnan(q)):
            e["qtd"] += q
        e["linhas"] += 1
        e["nao_encontrado"] = e["nao_encontrado"] or bool(r.get("naoEncontrado"))
    itens = sorted(agreg.values(), key=lambda x: (x["origem"], x["operacao"], x["ativo"]))
    for e in itens:
        e["qtd"] = round(e["qtd"], 3)
    return _json_seguro({"com_plano": bool(a.get("com_plano")), "linhas_da_totalizadora": len(tot), "itens": itens[:LIMITE_LINHAS], "cortado": len(itens) > LIMITE_LINHAS,
                         "nao_encontrados": sorted({i["ativo"] for i in itens if i["nao_encontrado"]})})


def _calcular_payload(ctx: Contexto, payload: dict) -> dict:
    from routers.vinculacao_config import perfil_do_projeto, postes_do_projeto
    from services.orcamento_calc import processar_calculo

    p = ctx.pipe()
    if not p.base_orcamento:
        raise ValueError("Base de orçamento indisponível.")
    return processar_calculo(payload["cabos"], payload["outros"], p.projeto_nome or ctx.projeto_nome, p.base_orcamento, obra=ctx.obra,
                             convencao_postes=postes_do_projeto(ctx.projeto_codigo), perfil_exportacao=perfil_do_projeto(ctx.projeto_codigo))


def _orcamento(ctx: Contexto, cabos: list, outros: list) -> dict:
    from services.autonomo.totalizadora import payload_calculo
    return _calcular_payload(ctx, payload_calculo(_totalizar_bruto(ctx, cabos, outros)))


def comparar_orcamento(ctx: Contexto, antes_payload: dict, depois_payload: dict) -> dict:
    antes = _calcular_payload(ctx, antes_payload)
    depois = _calcular_payload(ctx, depois_payload)

    def agrupar(r):
        grupos = {}
        for linha in r.get("resultado") or []:
            total = linha.get("total")
            if not isinstance(total, (int, float)) or not math.isfinite(total):
                continue
            chave = tuple(_texto(linha.get(k)) for k in ("operacao", "mdo", "codigo"))
            grupos[chave] = grupos.get(chave, 0) + total
        return grupos

    a, d = agrupar(antes), agrupar(depois)
    mudancas = []
    for op, mdo, codigo in sorted(a.keys() | d.keys()):
        chave = op, mdo, codigo
        qa, qd = round(a.get(chave, 0), 6), round(d.get(chave, 0), 6)
        if qa != qd:
            mudancas.append({"operacao": op, "mdo": mdo, "codigo": codigo, "antes": qa, "depois": qd})
    return _json_seguro({"mudancas": mudancas[:20], "total_mudancas": len(mudancas),
                         "nao_encontrados": depois.get("nao_encontrados") or [], "avisos": (depois.get("avisos") or [])[:10]})


def calcular_orcamento(ctx: Contexto, a: dict) -> dict:
    """Totalizadora → cálculo do orçamento. `com_plano` simula as alterações propostas."""
    r = _orcamento(ctx, *_tabelas_para_calculo(ctx, bool(a.get("com_plano"))))
    limite = min(max(_inteiro(a.get("limite")) or 60, 1), 200)
    linhas = r.get("resultado") or []
    por_mdo = {}
    for l in linhas:
        k = (l.get("mdo") or "").strip() or "MAO-DE-OBRA"
        por_mdo[k] = por_mdo.get(k, 0) + 1
    return _json_seguro({
        "com_plano": bool(a.get("com_plano")), "linhas_do_orcamento": len(linhas), "linhas_por_mdo": por_mdo,
        "itens": [{"operacao": l.get("operacao"), "mdo": l.get("mdo"), "codigo": l.get("codigo"), "descricao": l.get("desc_codigo"), "total": l.get("total")} for l in linhas[:limite]],
        "cortado": len(linhas) > limite, "nao_encontrados": r.get("nao_encontrados") or [], "avisos": (r.get("avisos") or [])[:20]})


def listar_obras(ctx: Contexto, a: dict) -> dict:
    obras = ctx.carregar_obras() if ctx.carregar_obras else []
    return {"obras": [{"id": o.get("id"), "nome": o.get("nome"), "data": o.get("data"), "tipo": o.get("tipo"), "publica": o.get("publica"), "minha": o.get("minha"), "dono": o.get("dono")}
                      for o in obras[:60]], "total": len(obras)}


def consultar_obra(ctx: Contexto, a: dict) -> dict:
    from services.obras_contexto import resumir_obra
    obra = ctx.buscar_obra(_texto(a.get("obra_id"))) if ctx.buscar_obra else None
    if not obra:
        return {"erro": "Obra não encontrada entre as suas ou as públicas do projeto (use listar_obras)."}
    return {"id": obra.get("id"), "conteudo": resumir_obra(obra)}


def propor_obra(ctx: Contexto, a: dict) -> dict:
    modo = _texto(a.get("modo")).lower()
    if modo not in ("adicionar", "subtrair", "carregar"):
        return {"erro": "modo deve ser adicionar, subtrair ou carregar."}
    obra = ctx.buscar_obra(_texto(a.get("obra_id"))) if ctx.buscar_obra else None
    if not obra:
        return {"erro": "Obra não encontrada (use listar_obras e o id exato)."}
    if ctx.plano:
        return {"erro": "Já há alterações no plano; não misture obra com outras alterações."}
    if obra.get("tipo") == "modelo":
        aviso = "É um MODELO: a tela abrirá a janela de parâmetros e o usuário digita os valores (nunca invente)."
    else:
        aviso = ""
    ctx.plano.append({"n": 1, "acao": "obra", "modo": modo, "obra_id": obra["id"], "nome": obra.get("nome"), "tipo": obra.get("tipo") or "obra", "avisos": [aviso] if aviso else []})
    return {"aceita": True, "aviso": aviso, "proximo_passo": "O usuário confirma na tela. Não diga que já aplicou."}


def ajustar_tela(ctx: Contexto, a: dict) -> dict:
    acao = _texto(a.get("acao")).lower()
    alvo = _texto(a.get("alvo")).lower() or "ambos"
    if alvo not in ("cabos", "outros", "ambos"):
        return {"erro": "alvo deve ser cabos, outros ou ambos."}
    if acao == "ordenar":
        coluna, ordem = _texto(a.get("coluna")), _texto(a.get("ordem")).lower() or "asc"
        if coluna not in ("entidade", "operacao", "ativo", "qtdAtivos") or ordem not in ("asc", "desc"):
            return {"erro": "ordenar: coluna entre entidade|operacao|ativo|qtdAtivos e ordem asc|desc."}
        ctx.acoes_ui.append({"acao_ui": "ordenar", "alvo": alvo, "coluna": coluna, "ordem": ordem})
    elif acao == "filtrar":
        coluna, valor = _texto(a.get("coluna")), _texto(a.get("valor"))
        if coluna not in ("entidade", "operacao", "ativo") or not valor:
            return {"erro": "filtrar: informe coluna (entidade|operacao|ativo) e valor."}
        ctx.acoes_ui.append({"acao_ui": "filtrar", "alvo": alvo, "coluna": coluna, "valor": valor})
    elif acao == "limpar_filtros":
        ctx.acoes_ui.append({"acao_ui": "limpar_filtros", "alvo": alvo})
    else:
        return {"erro": "acao deve ser ordenar, filtrar ou limpar_filtros."}
    return {"ok": True}


# ── registro (o que o modelo enxerga) ─────────────────────────────────────────────────────────────────
def _obj(props: dict, obrigatorios=()) -> dict:
    """JSON Schema de um objeto; `required` só aparece quando há obrigatórios (alguns provedores recusam a lista vazia)."""
    sch = {"type": "object", "properties": props}
    if obrigatorios:
        sch["required"] = list(obrigatorios)
    return sch


_S = {"type": "string"}
_ITEM_ALTERACAO = _obj({
    "acao": {"type": "string", "enum": ["adicionar", "editar", "remover"]},
    "tabela": {"type": "string", "enum": ["cabos", "outros"]},
    "id": {"type": "integer", "description": "id da linha (de ler_tabela); só para editar/remover"},
    "operacao": {"type": "string", "enum": list(OPERACOES), "description": "I instalar, R remover, M manter. * = linha viva. Ao editar, omita para preservar."},
    "ativo": {"type": "string", "description": "texto completo da linha (editar SUBSTITUI a linha inteira)"},
}, ("acao", "tabela"))

FERRAMENTAS = {
    "listar_skills": (listar_skills, "Lista as skills (guias do projeto) com a descrição de cada uma.", _obj({})),
    "carregar_skill": (carregar_skill, "Lê uma skill pelo nome (ex.: ativos, montar_projeto, orcamento, formato_ativos).", _obj({"nome": _S}, ("nome",))),
    "buscar_ativo": (buscar_ativo, "Procura ativos REAIS na base do projeto por código ou palavras da descrição (ex.: 'poste 11 300', 'TR125', 'chave fusível'). Use SEMPRE antes de propor um código que você não tem certeza.",
                     _obj({"texto": _S, "limite": {"type": "integer"}}, ("texto",))),
    "detalhar_ativo": (detalhar_ativo, "Mostra os componentes (códigos de orçamento, fatores, MDO) de UM ativo da base.", _obj({"ativo": _S}, ("ativo",))),
    "ler_tabela": (ler_tabela, "Lê as linhas atuais com IDs. Filtre por texto e use proximo_inicio na página seguinte da mesma tabela.", _obj({
        "tabela": {"type": "string", "enum": ["cabos", "outros", "ambas"]}, "texto": _S, "inicio": {"type": "integer"}, "limite": {"type": "integer"}})),
    "consultar_vinculos": (consultar_vinculos, "Explica os vínculos e pendências do poste ou vão selecionado no mapa, ou de um ID da tabela. Estados gr completo, k sem vínculo, pu parcial, rd inválido.",
                           _obj({"cabo_id": {"type": "integer"}, "poste_id": {"type": "integer"}})),
    "validar_linhas": (validar_linhas, "Confere se linhas candidatas são reconhecidas pela base/Totalizadora, sem alterar nada.",
                       _obj({"linhas": {"type": "array", "items": _obj({"tabela": {"type": "string", "enum": ["cabos", "outros"]}, "ativo": _S, "operacao": _S}, ("tabela", "ativo"))}}, ("linhas",))),
    "propor_alteracoes": (propor_alteracoes, "Propõe adicionar/editar/remover linhas. NADA é aplicado: o usuário vê o plano e confirma. Use o id de ler_tabela para editar/remover.",
                          _obj({"resumo": _S, "alteracoes": {"type": "array", "items": _ITEM_ALTERACAO}, "substituir": {"type": "boolean", "description": "descarta o plano anterior"}}, ("alteracoes",))),
    "totalizar": (totalizar, "Roda a Totalizadora (regras de conversão do projeto) e devolve os ativos totais. com_plano=true simula as alterações propostas.", _obj({"com_plano": {"type": "boolean"}})),
    "calcular_orcamento": (calcular_orcamento, "Calcula o orçamento (itens ODI/ODD) da tabela atual; com_plano=true inclui as alterações propostas.", _obj({"com_plano": {"type": "boolean"}, "limite": {"type": "integer"}})),
    "listar_obras": (listar_obras, "Lista as obras salvas visíveis ao usuário neste projeto.", _obj({})),
    "consultar_obra": (consultar_obra, "Resume o conteúdo de uma obra salva.", _obj({"obra_id": _S}, ("obra_id",))),
    "propor_obra": (propor_obra, "Propõe adicionar, subtrair ou carregar uma obra salva (o usuário confirma na tela).",
                    _obj({"obra_id": _S, "modo": {"type": "string", "enum": ["adicionar", "subtrair", "carregar"]}}, ("obra_id", "modo"))),
    "ajustar_tela": (ajustar_tela, "Ordena ou filtra as tabelas na tela (não altera dados).",
                     _obj({"acao": {"type": "string", "enum": ["ordenar", "filtrar", "limpar_filtros"]}, "alvo": {"type": "string", "enum": ["cabos", "outros", "ambos"]},
                           "coluna": _S, "ordem": {"type": "string", "enum": ["asc", "desc"]}, "valor": _S}, ("acao",))),
}


def especificacoes() -> list:
    """[{name, description, parameters}] neutros: cada provedor converte para o seu formato."""
    return [{"name": n, "description": d, "parameters": s} for n, (_f, d, s) in FERRAMENTAS.items()]


def executar(ctx: Contexto, nome: str, args) -> dict:
    """Roda uma ferramenta; sempre devolve um dict (erro incluso) para o modelo ler."""
    if nome not in FERRAMENTAS:
        return {"erro": f"Ferramenta '{nome}' não existe.", "disponiveis": sorted(FERRAMENTAS)}
    if isinstance(args, str):
        try:
            args = json.loads(args or "{}")
        except ValueError:
            return {"erro": "argumentos não são um JSON válido."}
    if not isinstance(args, dict):
        args = {}
    try:
        return _json_seguro(FERRAMENTAS[nome][0](ctx, args))
    except Exception as e:             # a conversa nunca cai por causa de uma ferramenta
        return {"erro": f"{type(e).__name__}: {e}"[:300]}
