"""
services/orcamento_detalhes.py — Descrição, filtro e tipo de MDO de códigos, tirados da Base de Orçamento (TASK-129).

Usado pelo Resultado do Orçamento ao importar/somar/subtrair um REC, colar itens e no modo comparação: o REC só traz código e quantidade,
então descrição e filtro vêm da base (a mesma do cálculo: master/nuvem, não a cópia pessoal de cada usuário).
Vários ativos podem ter o mesmo código: vale a linha do PROJETO informado, depois as genéricas (projeto vazio), depois qualquer uma;
a descrição é da primeira linha que tem descrição e o filtro é o primeiro filtro preenchido nessa mesma ordem de preferência.
"""


def _norm(v) -> str:
    return str(v if v is not None else "").strip().upper()


def _prioridade(linha: dict, projeto: str) -> int:
    lp = _norm(linha.get("projeto"))
    if projeto and lp and projeto in [x.strip() for x in lp.split("/")]:
        return 0
    return 1 if not lp else 2


def detalhes_por_codigo(linhas, codigos, projeto=None) -> dict:
    """{código pedido: {"desc_codigo", "mdo", "filtro"}} só para os códigos que existem na base."""
    proj = _norm(projeto)
    pedidos = {}
    for c in codigos or []:
        k = _norm(c)
        if k:
            pedidos.setdefault(k, []).append(str(c).strip())
    por_codigo = {}
    for ordem, l in enumerate(linhas or []):
        k = _norm(l.get("codigo"))
        if k in pedidos:
            por_codigo.setdefault(k, []).append((_prioridade(l, proj), ordem, l))
    saida = {}
    for k, cand in por_codigo.items():
        cand.sort(key=lambda t: (t[0], t[1]))
        ordenadas = [l for _, _, l in cand]
        com_desc = next((l for l in ordenadas if str(l.get("desc_codigo") or "").strip()), ordenadas[0])
        filtro = next((str(l.get("filtro")).strip() for l in ordenadas if str(l.get("filtro") or "").strip()), "")
        det = {"desc_codigo": com_desc.get("desc_codigo") or "", "mdo": com_desc.get("mdo"), "filtro": filtro}
        for original in pedidos[k]:
            saida[original] = det
    return saida
