"""
services/cava_rocha.py — TASK-248: cava em rocha (SDEMU3007II) igual à aba CAVA_ROCHA da ICP.

A parte `cava_rocha` do pacote do projeto traz os dados (ativo, serviço, códigos contados e o índice de cada localidade).
`cavas`/`calcular` são a tradução célula a célula da aba (anexo .ai/tasks/anexos/TASK-248/cava_icp.py), SEM "consertar" a conta:
    F     = estais em solo + postes de 9 m para cima (os códigos de `contagem`, nos serviços de instalação da ODI)
    cavas = F se ARREDMULTB(MARRED(F × índice; 0,1); 1) > F, senão ARREDMULTB(MARRED(F × índice; 0,1); 1)
Localidade fora da tabela → 0. static/cava_rocha.js tem a mesma função (paridade testada em tests/test_cava_rocha_248.py).
"""
import re
from decimal import Decimal, ROUND_FLOOR, ROUND_HALF_UP

FORMATO = "cava-rocha"


def _d(x):
    return Decimal(format(x, '.15g'))                    # o Excel trabalha com 15 algarismos significativos


def mround(x, m):
    """MARRED: múltiplo de m mais próximo, meio para longe do zero."""
    if x == 0:
        return 0.0
    q = (_d(x) / _d(m)).quantize(Decimal(1), rounding=ROUND_HALF_UP)
    return float(q * _d(m))


def floor1(x):
    """ARREDMULTB(x; 1) para x >= 0."""
    return float(_d(x).quantize(Decimal(1), rounding=ROUND_FLOOR))


def cavas(postes_estais, indice):
    """Quantidade de cavas em rocha de UMA localidade (linha da aba)."""
    f = postes_estais
    g = floor1(mround(f * indice, 0.1))
    return f if g > f else g


def calcular(localidade, qtd_servicos, tabela, contagem=("SDEMU3002II", "SDEMU1004II", "SDEMU1005II")):
    """localidade = DASHBORD!C4; qtd_servicos = {codigo: quantidade}; tabela = [(localidade, indice), ...] na ordem da aba.
    Devolve {'f', 'cavas'} — cavas = quantidade de SDEMU3007II (0 quando a localidade não está na tabela)."""
    f = 0
    for c in contagem:
        f += qtd_servicos.get(c, 0) or 0
    total = 0.0
    for loc, indice in tabela:
        if loc == localidade:                            # a aba soma todas as linhas iguais à localidade
            total += cavas(f, indice)
    return {"f": f, "cavas": total}


# ── dados da parte `cava_rocha` ───────────────────────────────────────────────────────────────
def validar(v) -> list:
    """Erros da parte `cava_rocha` (lista vazia = ok)."""
    if not isinstance(v, dict):
        return ["a parte cava_rocha precisa ser um objeto JSON"]
    erros = []
    if v.get("formato") != FORMATO:
        erros.append(f"formato deve ser '{FORMATO}'")
    if not str(v.get("ativo") or "").strip():
        erros.append("ativo precisa ser o nome do ativo da base (ex.: CAVA_ROCHA)")
    cont = v.get("contagem")
    if not isinstance(cont, list) or not cont or not all(isinstance(c, str) and c.strip() for c in cont):
        erros.append("contagem precisa da lista de códigos de serviço contados")
    locs = v.get("localidades")
    if not isinstance(locs, list) or not locs:
        erros.append("localidades precisa de pelo menos uma localidade")
    else:
        for i, l in enumerate(locs, 1):
            ind = l.get("indice") if isinstance(l, dict) else None
            if not isinstance(l, dict) or not str(l.get("localidade") or "").strip() or not isinstance(ind, (int, float)) or isinstance(ind, bool) or ind < 0:
                erros.append(f"localidades, item {i}: precisa de localidade e indice numérico ≥ 0")
                break
    return erros


def dados_do_projeto(projeto) -> dict:
    """A parte `cava_rocha` do projeto (nome ou código): a importada no pacote ou, na falta, a do pacote de partida; {} se não há."""
    chave = str(projeto or "").strip()
    if not chave:
        return {}
    from services import pacote_projeto as pp
    proj = pp.projeto_do_cadastro(chave)
    v = (pp.parte_guardada(proj["codigo"], "cava_rocha") if proj else None) or pp.parte_inicial(chave, "cava_rocha")
    return v if isinstance(v, dict) and not validar(v) else {}


def _norm(t) -> str:
    from services.obra_params import normalizar
    return normalizar(t).replace("_", " ")


def resolver_localidade(obra, dados) -> dict:
    """Localidade da obra → {localidade, indice, relacao, origem}: pelo bairro; senão pelo município. localidade '' = fora da tabela."""
    obra = obra or {}
    locs = (dados or {}).get("localidades") or []
    for campo, rotulo in (("bairro", "bairro"), ("municipio", "município")):
        alvo = _norm(obra.get(campo))
        if not alvo:
            continue
        for l in locs:
            if _norm(l.get("localidade")) == alvo:
                return {"localidade": l["localidade"], "indice": l["indice"], "relacao": l.get("relacao") or "", "origem": f"{rotulo} {obra.get(campo)}"}
    return {"localidade": "", "indice": None, "relacao": "", "origem": ""}


_RE_LINHA = re.compile(r"^\s*\*?[0-9]+(?:[.,][0-9]+)?[-xX]\s*([^\s(,]+)")


def eh_linha_cava(ativo, dados) -> bool:
    """A linha de Outros é a cava em rocha (de um cálculo anterior)?"""
    nome_ativo = str((dados or {}).get("ativo") or "").strip().upper()
    if not nome_ativo:
        return False
    t = str(ativo or "").strip().upper()
    m = _RE_LINHA.match(t)
    return (m.group(1) if m else t.split("(")[0].strip()) == nome_ativo


def _qtd_texto(q) -> str:
    s = f"{q:.2f}".rstrip("0").rstrip(".")
    return s or "0"


def linha_para_outros(quantidade, dados) -> str:
    """'6-CAVA_ROCHA'; '' quando a quantidade é 0 (só remove a antiga)."""
    if not quantidade:
        return ""
    return f"{_qtd_texto(quantidade)}-{dados.get('ativo')}"


def qtd_do_resultado(resultado, dados) -> dict:
    """Qtd ODI por código contado: soma do `total` das linhas de instalação (I) do resultado do cálculo."""
    codigos = {str(c) for c in (dados or {}).get("contagem") or []}
    saida = {}
    for r in resultado or []:
        c = str(r.get("codigo") or "").strip()
        if r.get("operacao") == "I" and c in codigos:
            saida[c] = saida.get(c, 0) + (float(r.get("total") or 0))
    return saida


def ativo_na_base(rows, ativo, projeto) -> bool:
    """A base (linhas da tabela de orçamento) tem o ativo para o projeto (do projeto ou genérico)?"""
    alvo = str(ativo or "").strip().upper()
    proj = str(projeto or "").strip().upper()
    for r in rows or []:
        nome = str(r.get("ativo") or "").split("=")[0].strip().upper()
        if nome != alvo:
            continue
        rp = str(r.get("projeto") or "").strip().upper()
        if not rp or not proj or proj in [p.strip() for p in rp.split("/")]:
            return True
    return False
