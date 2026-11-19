"""
services/assistente/catalogo.py (TASK-191) — O catálogo de ativos de um projeto, lido da Base de Orçamento: índice, busca e o texto da skill `ativos`.

A base (`tabela_orcamento`) tem uma linha por componente de cada ativo; aqui elas viram UM registro por ativo (nome oficial, descrição, quantos componentes/códigos,
formas com parâmetros). Serve às duas coisas que o assistente precisa: BUSCAR um código de verdade (nunca inventar) e CONHECER a estrutura do projeto (skill gerada).
Só funções puras sobre a lista de linhas — quem lê o banco é o chamador.
"""
import re
import unicodedata

from services.ativo_parametros import nome_base

LIMITE_LISTA_SKILL = 250          # a skill gerada mostra até tantos ativos (o resto vem pela ferramenta de busca)
_RE_FAMILIA = re.compile(r"^\d*([A-Z]+)")


def _norm(txt) -> str:
    t = unicodedata.normalize("NFKD", str(txt or "")).encode("ascii", "ignore").decode()
    return re.sub(r"\s+", " ", t).strip().upper()


def _vale(row, projeto_nome: str, projeto_codigo: str) -> bool:
    """A linha vale no projeto: sem projeto (genérica) ou o projeto está na lista "A/B/C" (por nome ou por código)."""
    p = (row.get("projeto") or "").strip().upper()
    if not p:
        return True
    if not projeto_nome and not projeto_codigo:
        return True
    alvos = {x.strip() for x in p.split("/")}
    return (projeto_nome or "").strip().upper() in alvos or (projeto_codigo or "").strip().upper() in alvos


def indexar(rows, projeto_nome: str = "", projeto_codigo: str = "") -> dict:
    """{NOME: {nome, descricao, componentes, codigos, formas}} só com as linhas que valem no projeto. `componentes` conta linhas; `codigos`, os códigos distintos."""
    indice = {}
    for r in rows or []:
        celula = (r.get("ativo") or "").strip()
        if not celula or not _vale(r, projeto_nome, projeto_codigo):
            continue
        try:
            nome = nome_base(celula)
        except Exception:                      # célula malformada: ainda assim aparece pelo texto
            nome = celula.upper()
        if not nome:
            continue
        e = indice.setdefault(nome, {"nome": nome, "descricao": "", "componentes": 0, "codigos": set(), "formas": []})
        e["componentes"] += 1
        if not e["descricao"] and (r.get("desc_ativo") or "").strip():
            e["descricao"] = r["desc_ativo"].strip()
        if (r.get("codigo") or "").strip():
            e["codigos"].add(r["codigo"].strip())
        forma = celula.upper() if celula.upper() != nome else ""
        if forma and forma not in e["formas"] and len(e["formas"]) < 3:
            e["formas"].append(forma)
    for e in indice.values():
        e["codigos"] = len(e["codigos"])
    return indice


def buscar(indice: dict, texto: str, limite: int = 15) -> list:
    """Ativos que casam com `texto` (nome ou descrição), do mais exato ao mais frouxo: igual > começa com > contém no nome > contém na descrição.
    Cada palavra do texto precisa aparecer (nome ou descrição); "poste 11 300" acha DT11/300."""
    t = _norm(texto)
    if not t:
        return []
    palavras = [p for p in re.split(r"[\s,;]+", t) if p]
    achados = []
    for nome, e in indice.items():
        alvo_nome, alvo_desc = _norm(nome), _norm(e["descricao"])
        juntos = alvo_nome + " " + alvo_desc
        if alvo_nome == t:
            rank = 0
        elif alvo_nome.startswith(t):
            rank = 1
        elif t in alvo_nome:
            rank = 2
        elif all(p in juntos for p in palavras):
            rank = 3 if t in alvo_desc else 4
        else:
            continue
        achados.append((rank, len(nome), nome, e))
    achados.sort(key=lambda x: x[:3])
    return [x[3] for x in achados[:max(1, int(limite or 15))]]


def componentes_de(rows, nome: str, projeto_nome: str = "", projeto_codigo: str = "", limite: int = 40) -> list:
    """Os componentes (código, descrição, fatores) de UM ativo, só do projeto."""
    alvo = _norm(nome)
    saida = []
    for r in rows or []:
        celula = (r.get("ativo") or "").strip()
        if not celula or not _vale(r, projeto_nome, projeto_codigo):
            continue
        try:
            nb = nome_base(celula)
        except Exception:
            nb = celula.upper()
        if _norm(nb) != alvo:
            continue
        saida.append({"componente": r.get("componente") or "", "codigo": r.get("codigo") or "", "descricao": r.get("desc_codigo") or "",
                      "mdo": r.get("mdo") or "", "fator_i": r.get("fator_i"), "fator_r": r.get("fator_r"), "filtro": r.get("filtro") or ""})
        if len(saida) >= limite:
            break
    return saida


def familia(nome: str) -> str:
    """"DT11/300" → "DT"; "2ESTR" → "ESTR"; "C16" → "C"; sem letras → "#"."""
    m = _RE_FAMILIA.match(_norm(nome))
    return m.group(1) if m else "#"


def gerar_markdown(indice: dict, projeto_nome: str) -> str:
    """O texto da skill `ativos` do projeto (determinístico: sem data, ordem alfabética)."""
    nomes = sorted(indice)
    por_familia = {}
    for n in nomes:
        por_familia.setdefault(familia(n), []).append(n)
    cab = [
        f"descricao: Catálogo dos {len(nomes)} ativos cadastrados na Base de Orçamento do projeto {projeto_nome}: famílias, códigos válidos e descrições. Use antes de propor qualquer ativo.",
        f"# Catálogo de ativos — {projeto_nome}",
        "",
        "> Gerado automaticamente a partir da Base de Orçamento (botão «Gerar catálogo» nas Skills). Não edite à mão: gere de novo quando a base mudar.",
        "> Só os códigos listados aqui existem. Para confirmar um código, ou ver os componentes dele, use `buscar_ativo` / `detalhar_ativo`.",
        "",
        f"Total: {len(nomes)} ativos em {len(por_familia)} famílias (família = letras iniciais do código).",
        "",
        "## Famílias",
        "",
    ]
    for fam in sorted(por_familia, key=lambda f: (-len(por_familia[f]), f)):
        lista = por_familia[fam]
        exemplos = ", ".join(lista[:12]) + (", …" if len(lista) > 12 else "")
        cab.append(f"- **{fam}** ({len(lista)}): {exemplos}")
    cab += ["", "## Ativos (código — descrição)", ""]
    for n in nomes[:LIMITE_LISTA_SKILL]:
        e = indice[n]
        extra = f" · formas: {'; '.join(e['formas'])}" if e["formas"] else ""
        cab.append(f"- `{n}` — {(e['descricao'] or '(sem descrição)')[:80]} ({e['componentes']} comp.){extra}")
    if len(nomes) > LIMITE_LISTA_SKILL:
        cab.append(f"- … e mais {len(nomes) - LIMITE_LISTA_SKILL} ativos (use `buscar_ativo`).")
    return "\n".join(cab) + "\n"
