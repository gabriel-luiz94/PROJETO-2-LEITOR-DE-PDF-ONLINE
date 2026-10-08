"""
services/modelos_obra.py — Modelos de obra com variáveis `V` (TASK-058). Funções puras: sem banco, sem rede.

Um MODELO é uma obra (Cabos + Outros) em que quantidades podem ser a variável `V`:
  - Cabos:  o COMPRIMENTO, ex. `CAA 2 ABC V m` ou `CAA 2 ABC V(vão) m`;
  - Outros: a QUANTIDADE, ex. `V-U4`, `V(postes)-DT11/300`, `*V-CFU` (o `*` = quantidade negativa, decisão do modelo).
Só o token exato `V` (ou `V(nome)`) é variável: `CAV`, `V1`, `1-VA` não são. `V(nome)` repetido = UM campo; `V` sem nome = um campo por
ocorrência. Quem usa o modelo digita sempre números positivos; o sinal vem do modelo.

`gerar` devolve tabelas PADRÃO (sem `V`), no mesmo formato de ativo do contrato (RULES Regra 5), validadas pela camada 1.
"""
import copy
import json
import re

from services.validacao_planilhas import validar_planilhas

LIMITE_VALOR = 1_000_000
LIMITE_ROTULO = 80

_RE_CABO = re.compile(r"^(?P<pre>.*\S)\s+V(?:\((?P<nome>[^()\s]+)\))?(?P<m>\s+[mM])?\s*$")
_RE_OUTROS = re.compile(r"(?<!\S)(?P<neg>\*?)V(?:\((?P<nome>[^()\s]+)\))?(?=-\S)")
_RE_VALOR = re.compile(r"^\d+(\.\d+)?$")


class ErroModelo(Exception):
    def __init__(self, mensagens):
        self.mensagens = [mensagens] if isinstance(mensagens, str) else list(mensagens)
        super().__init__("; ".join(self.mensagens))


def tem_variavel(ativo, tabela) -> bool:
    """A linha (`tabela` = 'cabos' | 'outros') ainda tem alguma variável `V`?"""
    texto = str(ativo or "").strip()
    return bool(_RE_CABO.match(texto)) if tabela == "cabos" else bool(_RE_OUTROS.search(texto))


def _linhas(valor):
    dados = valor.get("data") if isinstance(valor, dict) else valor
    return dados if isinstance(dados, list) else []


def _ativo(linha):
    return str(linha.get("ativo") or "") if isinstance(linha, dict) else ""


def detectar_variaveis(cabos, outros) -> list:
    """Variáveis das duas tabelas, na ordem em que aparecem (Cabos, depois Outros).

    Cada item: {"chave", "nome" (None se sem nome), "ocorrencias": [{"tabela", "indice", "ativo"}]}.
    `chave`: o nome em minúsculas (campos nomeados com o mesmo nome são um só) ou `#<n>` (sem nome, n = ordem)."""
    achadas, por_chave, anonimas = [], {}, 0

    def registrar(nome, tabela, indice, ativo):
        nonlocal anonimas
        if nome:
            chave = nome.casefold()
        else:
            anonimas += 1
            chave = f"#{anonimas}"
        if chave not in por_chave:
            por_chave[chave] = {"chave": chave, "nome": nome or None, "ocorrencias": []}
            achadas.append(por_chave[chave])
        por_chave[chave]["ocorrencias"].append({"tabela": tabela, "indice": indice, "ativo": ativo})

    for i, linha in enumerate(_linhas(cabos)):
        m = _RE_CABO.match(_ativo(linha).strip())
        if m:
            registrar(m.group("nome"), "cabos", i, _ativo(linha))
    for i, linha in enumerate(_linhas(outros)):
        for m in _RE_OUTROS.finditer(_ativo(linha).strip()):
            registrar(m.group("nome"), "outros", i, _ativo(linha))
    return achadas


def _rotulo_padrao(v):
    return v["nome"] or f"Valor {v['chave'][1:]}"


def montar_parametros(variaveis, configurados=None) -> list:
    """Parâmetros do modelo = variáveis detectadas + rótulo/padrão já configurados (casados pela `chave`)."""
    cfg = {c.get("chave"): c for c in (configurados or []) if isinstance(c, dict)}
    saida = []
    for v in variaveis:
        c = cfg.get(v["chave"], {})
        rotulo = str(c.get("rotulo") or "").strip()[:LIMITE_ROTULO] or _rotulo_padrao(v)
        saida.append({"chave": v["chave"], "nome": v["nome"], "rotulo": rotulo, "padrao": _numero(c.get("padrao")),
                      "onde": [f"{'Cabos' if o['tabela'] == 'cabos' else 'Outros'} linha {o['indice'] + 1}: {o['ativo']}" for o in v["ocorrencias"]]})
    return saida


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
    """Confere a lista de parâmetros que o cliente manda ao salvar (formato e limites). Devolve a lista limpa."""
    if parametros is None:
        return []
    if not isinstance(parametros, list) or len(parametros) > 200:
        raise ErroModelo("Parâmetros do modelo inválidos.")
    limpos, erros = [], []
    for p in parametros:
        if not isinstance(p, dict) or not isinstance(p.get("chave"), str):
            raise ErroModelo("Parâmetros do modelo inválidos.")
        padrao = p.get("padrao")
        if padrao not in (None, "") and _numero(padrao) is None:
            erros.append(f"Valor padrão de '{p.get('rotulo') or p['chave']}' inválido: use um número maior que zero.")
        rot = p.get("rotulo")
        if rot is not None and (not isinstance(rot, str) or len(rot.strip()) > LIMITE_ROTULO):
            erros.append(f"Rótulo de '{p['chave']}' inválido (texto de até {LIMITE_ROTULO} caracteres).")
        limpos.append({"chave": p["chave"], "rotulo": rot if isinstance(rot, str) else "", "padrao": _numero(padrao)})
    if erros:
        raise ErroModelo(erros)
    return limpos


def gerar(cabos, outros, configurados, valores) -> dict:
    """Aplica `valores` ({chave: número}) ao modelo e devolve {"cabos": [...], "outros": [...]} sem nenhum `V`.

    Faltou valor -> usa o padrão do modelo; sem padrão, erro. Valor precisa ser número > 0 (vírgula ou ponto)."""
    valores = valores if isinstance(valores, dict) else {}
    variaveis = detectar_variaveis(cabos, outros)
    parametros = {p["chave"]: p for p in montar_parametros(variaveis, configurados)}
    erros, escolhidos = [], {}
    for chave in valores:
        if chave not in parametros:
            erros.append(f"Variável desconhecida: '{chave}'.")
    for chave, p in parametros.items():
        bruto = valores.get(chave)
        if bruto is None or str(bruto).strip() == "":
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
    contador = {"n": 0}

    def chave_de(nome):
        if nome:
            return nome.casefold()
        contador["n"] += 1
        return f"#{contador['n']}"

    tocadas = {"cabos": set(), "outros": set()}
    for i, linha in enumerate(novos_cabos):
        if not isinstance(linha, dict):
            continue
        m = _RE_CABO.match(_ativo(linha).strip())
        if m:
            valor = formatar(escolhidos[chave_de(m.group("nome"))])
            linha["ativo"] = f"{m.group('pre')} {valor}" + (f" {m.group('m').strip()}" if m.group("m") else "")
            tocadas["cabos"].add(i)
    for i, linha in enumerate(novos_outros):
        if not isinstance(linha, dict):
            continue

        def troca(m):
            return f"{m.group('neg')}{formatar(escolhidos[chave_de(m.group('nome'))])}"
        novo = _RE_OUTROS.sub(troca, _ativo(linha).strip())
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
