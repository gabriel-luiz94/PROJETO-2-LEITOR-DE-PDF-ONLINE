"""TASK-171 — a estrutura FICTÍCIA do cabo sugere os componentes avulsos (alças, conectores…) para a tabela Outros.

O cabo novo ancora numa estrutura existente que não está na obra (estrutura fictícia). A fictícia só completa a ponta do cabo; o material dela (os grupos da base que
CHAMAM ativos: ``N4(cabo=%,cabo2=%):=3*ALCA(cabo) + 3*ALCA(cabo2)``) não saía. Aqui se calcula o que sugerir — só o lado do cabo NOVO — e a tela deixa o usuário
adicionar com um clique (nada entra no cálculo sem o clique).

Este arquivo tem a MESMA lógica de static/avulsos_ficticia.js (paridade testada em tests/test_avulsos_ficticia_171.py).
"""
from services import vinculacao_regras as vr
from services.ativo_parametros import analisar_ativo, compostos_parametrizados, linha_vale, nome_tela


def grupos_de_chamada(linhas, projeto) -> dict:
    """{BASE: [{"restricoes", "itens": [{"ativo", "qtd"}]}]} — os grupos da base que chamam ativos (TASK-134), para o projeto."""
    return compostos_parametrizados(linhas, projeto)["grupos"]


def _nome_do_alvo(a, k, assinaturas):
    """Nome do k-ésimo argumento (posicional) do ativo chamado, pela assinatura dele na base; sem assinatura = sem nome."""
    sig = (assinaturas or {}).get(a["base"])
    nomes = [n for n in (sig or {}).get("nomes", [])]
    return nome_tela(nomes[k]).lower() if k < len(nomes) else None


def avulsos_da_ficticia(cabo, ficticia, cfg, prefixo_fn, assinaturas, grupos) -> list:
    """Itens avulsos sugeridos para o CABO ligado à estrutura fictícia `ficticia`: [{"ativo": "ALCA(cabo=3XA35P-13.8KV)", "qtd": 3.0, "operacao": "I", "completar": False}].
    1. a fictícia recebe o cabo nos parâmetros que a Tabela de Estruturas dá ao 1º cabo dela (mesmas fontes da TASK-100); os demais ficam no padrão da base;
    2. valem as linhas de grupo que chamam ativos da fictícia cuja condição casa com esses valores;
    3. de cada linha, só os itens que usam APENAS o parâmetro do cabo (ou valores fixos); item que usa o do cabo E outro sem valor entra com o outro em branco
       (``completar``: complete o cabo existente); item só do outro lado (``3*ALCA(cabo2)``) não entra. Itens iguais somam."""
    base = str(ficticia or "").strip().upper()
    sig = (assinaturas or {}).get(base)
    if not base or not sig or base not in (grupos or {}):
        return []
    cfg = vr.normalizar_config(cfg)
    params_cabo = vr.parametros_do_cabo_na_estrutura(base, cabo, cfg, prefixo_fn)
    if not params_cabo:
        return []
    pedido = {}
    for n in sig["nomes"]:
        if n in params_cabo and params_cabo[n]:
            pedido[n] = params_cabo[n][0]
        elif n in sig["padroes"]:
            pedido[n] = sig["padroes"][n]
    do_cabo = {nome_tela(n) for n in sig["nomes"] if n in params_cabo and params_cabo[n]}
    e = vr.info_estrutura(base, cfg)
    de_cabos = {nome_tela(p["nome"]) for p in (e["parametros"] or [])} or ({e["parametro"]} if e["parametro"] else set())          # nomes que a Tabela de Estruturas dá a CABOS
    outro_lado = de_cabos - do_cabo                                   # os cabos que já existem (cabo2…): o que a fictícia NÃO sabe
    valores = {nome_tela(n): v for n, v in pedido.items()}
    conhecidos = {nome_tela(n) for n in sig["nomes"]}
    op = (cabo.get("operacao") or "I")
    soma, ordem = {}, []
    for g in grupos[base]:
        if not linha_vale(g["restricoes"], pedido):
            continue
        for it in g["itens"]:
            a = analisar_ativo(it["ativo"])
            if a["erro"] or a["tipo"] != "grupo":
                continue
            partes, usa_cabo, usa_outro = [], False, False
            for k, q in enumerate(a["params"]):
                v = q["valor"].upper()
                if v in conhecidos:                               # o argumento é o NOME de um parâmetro da fictícia
                    val = valores.get(v)
                    if v in do_cabo:
                        usa_cabo = True
                    elif v in outro_lado or val is None:          # o cabo existente (ou parâmetro sem valor): em branco
                        usa_outro, val = True, None
                    texto_v = val if val is not None else "?"
                else:
                    texto_v = q["valor"]
                nome = nome_tela(q["nome"]).lower() if q["nome"] else _nome_do_alvo(a, k, assinaturas)
                partes.append(f"{nome}={texto_v}" if nome else texto_v)
            if usa_outro and not usa_cabo:
                continue                                          # só o outro lado (o cabo existente)
            chave = (f"{a['base']}({', '.join(partes)})", usa_outro)
            if chave not in soma:
                soma[chave] = 0.0
                ordem.append(chave)
            soma[chave] += float(it["qtd"])
    return [{"ativo": c[0], "qtd": soma[c], "operacao": op, "completar": c[1]} for c in ordem]


def avulsos_do_cabo(cabo, cfg, prefixo_fn, assinaturas, grupos) -> list:
    """Uma lista por estrutura fictícia do cabo: [{"ficticia", "itens"}] (só as que têm sugestão)."""
    saida = []
    for cod in vr.estruturas_ficticias_de(cabo):
        itens = avulsos_da_ficticia(cabo, cod, cfg, prefixo_fn, assinaturas, grupos)
        if itens:
            saida.append({"ficticia": cod, "itens": itens})
    return saida


def chave_origem(idx_cabo, ficticia) -> str:
    """Marca guardada na linha de Outros (`origemAvulso`): "<índice do cabo>:<fictícia>"."""
    return f"{idx_cabo}:{str(ficticia).upper()}"
