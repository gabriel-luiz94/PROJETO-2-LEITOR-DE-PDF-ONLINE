"""
services/modelos_obra.py — Modelos de obra com variáveis `V` (TASK-058). Funções puras: sem banco, sem rede.

Um MODELO é uma obra (Cabos + Outros) em que quantidades podem ser a variável `V`:
  - Cabos:  o COMPRIMENTO, ex. `CAA 2 ABC V m` ou `CAA 2 ABC V(vão) m`;
  - Outros: a QUANTIDADE, ex. `V-U4`, `V(postes)-DT11/300`, `*V-CFU` (o `*` = quantidade negativa, decisão do modelo).
Só o token exato `V` (ou `V(nome)`) é variável: `CAV`, `V1`, `1-VA` não são. `V(nome)` repetido = UM campo; `V` sem nome = um campo por
ocorrência. Quem usa o modelo digita sempre números positivos; o sinal vem do modelo.

ATIVO VARIÁVEL (TASK-060, só em Outros): no lugar do código do ativo, `X(A,B,C)` = lista de 2 a 30 opções (quem usa o modelo ESCOLHE uma).
`X(poste:A,B)` dá nome ao campo e `X(poste:)` reaproveita a mesma escolha; uma opção `@GRUPO` expande para os ativos do grupo do projeto
(no momento do uso). Combina com a quantidade: `V(qtd)-X(U3,U4)`, `*V-X(CFU,CFUA)`, `2-X(U3,U4)`. `X`, `X9`, `1-X` e `CAX(A,B)` não são variáveis.
Campos de ativo têm chave `ativo:<nome>` ou `ativo:#<n>` (as chaves das quantidades não mudam).

`gerar` devolve tabelas PADRÃO (sem `V`), no mesmo formato de ativo do contrato (RULES Regra 5), validadas pela camada 1.
"""
import copy
import fnmatch
import json
import re

from services.regras_dominio import _casa
from services.validacao_planilhas import validar_planilhas

LIMITE_VALOR = 1_000_000
LIMITE_ROTULO = 80
LIMITE_OPCOES_DIGITADAS = 30
LIMITE_OPCOES = 50

_RE_CABO = re.compile(r"^(?P<pre>.*\S)\s+V(?:\((?P<nome>[^()\s]+)\))?(?P<m>\s+[mM])?\s*$")
# item de Outros: [*]<qtd>-<ativo>; qtd = V | V(nome) | número; ativo = X(lista) | X(nome:lista) | X(nome:) | código
_RE_ITEM = re.compile(r"(?<!\S)(?P<neg>\*?)(?P<q>V(?:\((?P<vnome>[^()\s]+)\))?|\d+(?:[.,]\d+)?)-"
                      r"(?P<a>X\((?:(?P<xnome>[^():\s,]+):)?(?P<xlista>[^()\s]*)\)|\S+)")
_RE_VALOR = re.compile(r"^\d+(\.\d+)?$")
_RE_CODIGO = re.compile(r"^[^\s(),:\-*@]+$")


class ErroModelo(Exception):
    def __init__(self, mensagens):
        self.mensagens = [mensagens] if isinstance(mensagens, str) else list(mensagens)
        super().__init__("; ".join(self.mensagens))


def _e_variavel_q(m) -> bool:
    return m.group("q").startswith("V")


def _e_variavel_x(m) -> bool:
    return m.group("xlista") is not None          # só o formato completo X(...); `X(A,B` sem fechar é um código comum


def tem_variavel(ativo, tabela) -> bool:
    """A linha (`tabela` = 'cabos' | 'outros') ainda tem alguma variável (`V` ou `X(...)`)?"""
    texto = str(ativo or "").strip()
    if tabela == "cabos":
        return bool(_RE_CABO.match(texto))
    return any(_e_variavel_q(m) or _e_variavel_x(m) for m in _RE_ITEM.finditer(texto))


def _linhas(valor):
    dados = valor.get("data") if isinstance(valor, dict) else valor
    return dados if isinstance(dados, list) else []


def _ativo(linha):
    return str(linha.get("ativo") or "") if isinstance(linha, dict) else ""


def detectar_variaveis(cabos, outros) -> list:
    """Variáveis das duas tabelas, na ordem em que aparecem (Cabos, depois Outros; em cada item a quantidade antes do ativo).

    Cada item: {"chave", "nome" (None se sem nome), "tipo": "quantidade"|"ativo", "ocorrencias": [{"tabela","indice","ativo"}]}; os de ativo levam
    também `definicoes` (as listas escritas, uma por ocorrência que traz lista).
    `chave`: quantidade = nome em minúsculas ou `#<n>`; ativo = `ativo:` + nome em minúsculas ou `ativo:#<n>`."""
    achadas, por_chave, anon = [], {}, {"q": 0, "x": 0}

    def registrar(tipo, nome, tabela, indice, ativo, lista=None):
        if nome:
            chave = ("ativo:" if tipo == "ativo" else "") + nome.casefold()
        else:
            anon["x" if tipo == "ativo" else "q"] += 1
            chave = ("ativo:" if tipo == "ativo" else "") + f"#{anon['x' if tipo == 'ativo' else 'q']}"
        if chave not in por_chave:
            por_chave[chave] = {"chave": chave, "nome": nome or None, "tipo": tipo, "ocorrencias": []}
            if tipo == "ativo":
                por_chave[chave]["definicoes"] = []
            achadas.append(por_chave[chave])
        por_chave[chave]["ocorrencias"].append({"tabela": tabela, "indice": indice, "ativo": ativo})
        if tipo == "ativo" and lista:
            por_chave[chave]["definicoes"].append(lista)

    for i, linha in enumerate(_linhas(cabos)):
        m = _RE_CABO.match(_ativo(linha).strip())
        if m:
            registrar("quantidade", m.group("nome"), "cabos", i, _ativo(linha))
    for i, linha in enumerate(_linhas(outros)):
        for m in _RE_ITEM.finditer(_ativo(linha).strip()):
            if _e_variavel_q(m):
                registrar("quantidade", m.group("vnome"), "outros", i, _ativo(linha))
            if _e_variavel_x(m):
                registrar("ativo", m.group("xnome"), "outros", i, _ativo(linha), m.group("xlista") or None)
    return achadas


def _rotulo_padrao(v):
    if v["nome"]:
        return v["nome"]
    n = v["chave"].split("#")[-1]
    return f"Ativo {n}" if v["tipo"] == "ativo" else f"Valor {n}"


# ── opções de um campo de ativo ─────────────────────────────────────────────
def _membros(sel, grupos, universo, pilha, avisos):
    """Códigos de UM membro de grupo: literal entra direto; curinga/regex/@outro é resolvido contra o universo (base técnica) quando existe."""
    if isinstance(sel, list):
        return [c for s in sel for c in _membros(s, grupos, universo, pilha, avisos)]
    if isinstance(sel, str):
        t = sel.strip()
        if t.startswith("@"):
            g = t[1:].upper()
            if g in pilha:
                return []
            return [c for m in _como_lista((grupos or {}).get(g, [])) for c in _membros(m, grupos, universo, pilha + (g,), avisos)]
        if not any(ch in t for ch in "*?["):
            return [t.upper()]
    if not universo:
        avisos.append("Um grupo usa curinga/regex e a base técnica do projeto não está disponível: parte das opções do grupo pode faltar.")
        return []
    return [u for u in sorted(universo) if _casa(u, sel, grupos)]


def _como_lista(v):
    return v if isinstance(v, list) else [v]


def expandir_opcoes(brutas, grupos=None, universo=None):
    """(opções em maiúsculas sem repetição, erros, avisos) de uma lista escrita `A,B,@GRUPO`."""
    erros, avisos, opcoes = [], [], []
    itens = [x.strip() for x in str(brutas or "").split(",")]
    if len(itens) > LIMITE_OPCOES_DIGITADAS:
        erros.append(f"Opções demais na lista (máx. {LIMITE_OPCOES_DIGITADAS} escritas).")
        itens = itens[:LIMITE_OPCOES_DIGITADAS]
    for it in itens:
        if not it:
            erros.append("Há uma opção vazia na lista (vírgula sobrando).")
        elif it.startswith("@"):
            g = it[1:].upper()
            if grupos is not None and g not in grupos:
                erros.append(f"O grupo '{it}' não existe neste projeto.")
                continue
            achados = _membros(it, grupos, universo, (), avisos)
            if not achados and grupos is not None and (grupos.get(g) is not None):
                avisos.append(f"O grupo '{it}' não resultou em nenhum ativo.")
            opcoes += achados
        elif not _RE_CODIGO.match(it):
            erros.append(f"Opção '{it}' inválida (um código de ativo não tem espaço, hífen, vírgula, dois-pontos, asterisco ou parênteses).")
        else:
            opcoes.append(it.upper())
    opcoes = list(dict.fromkeys(opcoes))
    if len(opcoes) > LIMITE_OPCOES:
        erros.append(f"Opções demais depois de expandir os grupos ({len(opcoes)}; máx. {LIMITE_OPCOES}).")
    return opcoes, erros, avisos


def _opcoes_do_campo(v, grupos, universo):
    """(opções, erros, avisos) de um campo de ativo, conferindo se todas as definições escritas concordam."""
    erros = []
    defs = v.get("definicoes") or []
    if not defs:
        return [], [f"O campo '{_rotulo_padrao(v)}' (X(...)) não tem nenhuma lista de opções: escreva-a em pelo menos uma ocorrência, ex. X({v['nome'] or 'nome'}:A,B)."], []
    primeira = defs[0]
    for d in defs[1:]:
        if sorted(x.strip().upper() for x in d.split(",")) != sorted(x.strip().upper() for x in primeira.split(",")):
            erros.append(f"O campo '{_rotulo_padrao(v)}' foi escrito com listas diferentes ({primeira} × {d}); repita só o nome: X({v['nome']}:).")
            break
    opcoes, e2, avisos = expandir_opcoes(primeira, grupos, universo)
    erros += e2
    if len(opcoes) < 2 and not e2 and not any("base técnica" in a for a in avisos):      # sem a base não dá para expandir um curinga: só avisa
        erros.append(f"O campo '{_rotulo_padrao(v)}' precisa de pelo menos 2 opções (tem {len(opcoes)}).")
    return opcoes, erros, avisos


def montar_parametros(variaveis, configurados=None, grupos=None, universo=None) -> list:
    """Parâmetros do modelo = variáveis detectadas + rótulo/padrão já configurados (casados pela `chave`). Campo de ativo traz `opcoes`."""
    cfg = {c.get("chave"): c for c in (configurados or []) if isinstance(c, dict)}
    saida = []
    for v in variaveis:
        c = cfg.get(v["chave"], {})
        rotulo = str(c.get("rotulo") or "").strip()[:LIMITE_ROTULO] or _rotulo_padrao(v)
        onde = [f"{'Cabos' if o['tabela'] == 'cabos' else 'Outros'} linha {o['indice'] + 1}: {o['ativo']}" for o in v["ocorrencias"]]
        if v["tipo"] == "ativo":
            opcoes, _, _ = _opcoes_do_campo(v, grupos, universo)
            padrao = str(c.get("padrao") or "").strip().upper()
            saida.append({"chave": v["chave"], "nome": v["nome"], "tipo": "ativo", "rotulo": rotulo, "opcoes": opcoes,
                          "padrao": padrao if padrao in opcoes else None, "onde": onde})
        else:
            saida.append({"chave": v["chave"], "nome": v["nome"], "tipo": "quantidade", "rotulo": rotulo, "padrao": _numero(c.get("padrao")), "onde": onde})
    return saida


def analisar(cabos, outros, configurados=None, grupos=None, universo=None) -> dict:
    """Tudo o que a tela de "Salvar como modelo" precisa: {parametros, erros (impedem salvar), avisos (não impedem)}.
    Avisos: opção de ativo que não existe na base técnica do projeto (decisão do usuário: só avisar)."""
    variaveis = detectar_variaveis(cabos, outros)
    cfg = {c.get("chave"): c for c in (configurados or []) if isinstance(c, dict)}
    erros, avisos = [], []
    for v in variaveis:
        if v["tipo"] != "ativo":
            continue
        opcoes, e, a = _opcoes_do_campo(v, grupos, universo)
        erros += e
        avisos += a
        padrao = str((cfg.get(v["chave"]) or {}).get("padrao") or "").strip().upper()
        if padrao and padrao not in opcoes:
            erros.append(f"O padrão '{padrao}' de '{_rotulo_padrao(v)}' não está entre as opções.")
        if universo:
            fora = [o for o in opcoes if o not in universo]
            if fora:
                avisos.append(f"Campo '{_rotulo_padrao(v)}': {', '.join(fora[:8])}{'…' if len(fora) > 8 else ''} não encontrado(s) na base técnica do projeto (o orçamento acusaria ativo não encontrado).")
    return {"parametros": montar_parametros(variaveis, configurados, grupos, universo), "erros": list(dict.fromkeys(erros)), "avisos": list(dict.fromkeys(avisos))}


def _numero(valor):
    """Número positivo (aceita vírgula) ou None."""
    if valor is None or isinstance(valor, bool):
        return None
    texto = str(valor).strip().replace(",", ".")
    if not _RE_VALOR.match(texto):
        return None
    n = float(texto)
    return n if 0 < n <= LIMITE_VALOR else None


def formatar(n: float) -> str:
    texto = f"{round(n, 4):.4f}".rstrip("0").rstrip(".")
    return texto or "0"


def validar_configuracao(parametros) -> list:
    """Confere a lista de parâmetros que o cliente manda ao salvar (formato e limites). Devolve a lista limpa.
    `padrao` é número positivo (campo de quantidade) ou código de ativo (campo de ativo; a conferência com as opções é feita em `analisar`)."""
    if parametros is None:
        return []
    if not isinstance(parametros, list) or len(parametros) > 200:
        raise ErroModelo("Parâmetros do modelo inválidos.")
    limpos, erros = [], []
    for p in parametros:
        if not isinstance(p, dict) or not isinstance(p.get("chave"), str):
            raise ErroModelo("Parâmetros do modelo inválidos.")
        padrao = p.get("padrao")
        if p["chave"].startswith("ativo:"):
            texto = "" if padrao is None else str(padrao).strip().upper()
            if isinstance(padrao, bool) or (texto and (len(texto) > 60 or not _RE_CODIGO.match(texto))):
                erros.append(f"Padrão de '{p.get('rotulo') or p['chave']}' inválido: use um dos códigos da lista.")
            padrao_ok = texto or None
        else:
            if padrao not in (None, "") and _numero(padrao) is None:
                erros.append(f"Valor padrão de '{p.get('rotulo') or p['chave']}' inválido: use um número maior que zero.")
            padrao_ok = _numero(padrao)
        rot = p.get("rotulo")
        if rot is not None and (not isinstance(rot, str) or len(rot.strip()) > LIMITE_ROTULO):
            erros.append(f"Rótulo de '{p['chave']}' inválido (texto de até {LIMITE_ROTULO} caracteres).")
        limpos.append({"chave": p["chave"], "rotulo": rot if isinstance(rot, str) else "", "padrao": padrao_ok})
    if erros:
        raise ErroModelo(erros)
    return limpos


def gerar(cabos, outros, configurados, valores, grupos=None, universo=None) -> dict:
    """Aplica `valores` ({chave: número ou código}) ao modelo e devolve {"cabos": [...], "outros": [...]} sem nenhuma variável.

    Quantidade: faltou valor -> padrão do modelo; sem padrão, erro; precisa ser número > 0 (vírgula ou ponto).
    Ativo: precisa ser UMA das opções do campo (faltou -> padrão; sem padrão, erro); nunca texto livre."""
    valores = valores if isinstance(valores, dict) else {}
    variaveis = detectar_variaveis(cabos, outros)
    analise = analisar(cabos, outros, configurados, grupos, universo)
    if analise["erros"]:
        raise ErroModelo(analise["erros"])
    parametros = {p["chave"]: p for p in analise["parametros"]}
    erros, escolhidos = [], {}
    for chave in valores:
        if chave not in parametros:
            erros.append(f"Variável desconhecida: '{chave}'.")
    for chave, p in parametros.items():
        bruto = valores.get(chave)
        vazio = bruto is None or str(bruto).strip() == ""
        if p["tipo"] == "ativo":
            if vazio:
                if p["padrao"] is None:
                    erros.append(f"Escolha o ativo de '{p['rotulo']}'.")
                else:
                    escolhidos[chave] = p["padrao"]
            elif isinstance(bruto, bool) or str(bruto).strip().upper() not in p["opcoes"]:
                erros.append(f"Escolha inválida para '{p['rotulo']}': use uma das opções ({', '.join(p['opcoes'][:10])}{'…' if len(p['opcoes']) > 10 else ''}).")
            else:
                escolhidos[chave] = str(bruto).strip().upper()
            continue
        if vazio:
            if p["padrao"] is None:
                erros.append(f"Informe o valor de '{p['rotulo']}'.")
            else:
                escolhidos[chave] = p["padrao"]
            continue
        n = _numero(bruto)
        if n is None:
            erros.append(f"Valor inválido para '{p['rotulo']}': use um número maior que zero (até {LIMITE_VALOR:,}).".replace(",", "."))
        else:
            escolhidos[chave] = n
    if erros:
        raise ErroModelo(erros)

    novos_cabos = [copy.deepcopy(l) for l in _linhas(cabos)]
    novos_outros = [copy.deepcopy(l) for l in _linhas(outros)]
    # reaplica a mesma numeração das anônimas na mesma ordem da detecção
    contador = {"q": 0, "x": 0}

    def chave_de(tipo, nome):
        if nome:
            return ("ativo:" if tipo == "x" else "") + nome.casefold()
        contador[tipo] += 1
        return ("ativo:" if tipo == "x" else "") + f"#{contador[tipo]}"

    tocadas = {"cabos": set(), "outros": set()}
    for i, linha in enumerate(novos_cabos):
        if not isinstance(linha, dict):
            continue
        m = _RE_CABO.match(_ativo(linha).strip())
        if m:
            valor = formatar(escolhidos[chave_de("q", m.group("nome"))])
            linha["ativo"] = f"{m.group('pre')} {valor}" + (f" {m.group('m').strip()}" if m.group("m") else "")
            tocadas["cabos"].add(i)
    for i, linha in enumerate(novos_outros):
        if not isinstance(linha, dict):
            continue

        def troca(m):
            eh_q, eh_x = _e_variavel_q(m), _e_variavel_x(m)
            if not (eh_q or eh_x):
                return m.group(0)
            qtd = formatar(escolhidos[chave_de("q", m.group("vnome"))]) if eh_q else m.group("q")
            ativo = escolhidos[chave_de("x", m.group("xnome"))] if eh_x else m.group("a")
            return f"{m.group('neg')}{qtd}-{ativo}"
        novo = _RE_ITEM.sub(troca, _ativo(linha).strip())
        if novo != _ativo(linha).strip():
            linha["ativo"] = novo
            tocadas["outros"].add(i)

    achados = validar_planilhas(
        [dict(l, id=f"CABOS-{i}") if isinstance(l, dict) else {"id": f"CABOS-{i}"} for i, l in enumerate(novos_cabos)],
        [dict(l, id=f"OUTROS-{i}") if isinstance(l, dict) else {"id": f"OUTROS-{i}"} for i, l in enumerate(novos_outros)])
    ruins = [a["mensagem"] for a in achados if a["severidade"] == "erro"
             and any(a["linha_id"] == f"{t.upper()}-{i}" for t in tocadas for i in tocadas[t])]
    if ruins:
        raise ErroModelo(["O modelo gera linhas inválidas com esses valores:"] + ruins[:10])
    return {"cabos": novos_cabos, "outros": novos_outros}


def separar_dados(dados_json):
    """`dados_json` salvo -> (cabos, outros, configurados). Aceita o formato da tela ({cabos:{data}}) e listas diretas."""
    try:
        snap = json.loads(dados_json) if isinstance(dados_json, str) else (dados_json or {})
    except (TypeError, ValueError):
        snap = {}
    if not isinstance(snap, dict):
        snap = {}
    modelo = snap.get("modelo") if isinstance(snap.get("modelo"), dict) else {}
    return _linhas(snap.get("cabos")), _linhas(snap.get("outros")), modelo.get("parametros") or []
