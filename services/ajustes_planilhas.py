"""
services/ajustes_planilhas.py — Motor de AJUSTES das planilhas Cabos e Outros (TASK-022, ADR-006).

Função pura: sem I/O, sem IA. Recebe ações declarativas (JSON) e as tabelas como aparecem na tela e devolve um DIFF
previsível — o motor NUNCA altera os dados de entrada; quem aplica é o frontend, depois do aceite do usuário.

AÇÕES (cada uma exige `acao` e `tabela`: "cabos" | "outros" | "ambos"; `id` e `descricao` são livres):

  substituir        {"de": "SUP-L", "para": "SUPL", "modo": "texto"(padrão)|"item", "palavra_inteira": true}
                      texto: troca no texto do ativo (literal, sem distinguir maiúsculas, ou {"regex": ".."});
                      item (só Outros): o ativo de cada item <qtd>-<ativo> que casa com o seletor `de` (código, curinga,
                      @GRUPO, lista, {regex}) passa a `para`, mantendo a quantidade.
  normalizar        {"regras": ["espacos", "maiusculas", "poste"]}   poste: "1-DT11/300" → "DT11/300" no início da linha
  ordenar           {"por": [{"coluna": "operacao"|"ativo"|"entidade", "ordem": "asc"|"desc", "valores": ["I","*I",..]}]}
  excluir_linhas    {"onde": {"vazias": true} | {"duplicadas": true} | {"condicao": COND} | {"texto": "regex"}}
  adicionar_linha   {"valores": {"operacao": "I", "ativo": "1-RA2", "entidade": "0"},
                     "posicao": "fim"|"inicio"|{"depois_de": COND}|{"antes_de": COND}, "apenas_se_nao_existir": true}
  adicionar_ativo   (Outros) {"ativo": "SUPL", "qtd": 1, "se_ja_existe": "ignorar"|"somar"|"substituir"}
                      qtd negativo gera o prefixo "*" (linha viva/retirada, mesma leitura de Outros em geral: "*1-PR").
  remover_ativo     (Outros) {"ativo": SELETOR}
  mesclar_duplicadas(Outros) soma as quantidades do mesmo ativo repetido na linha

Campos comuns opcionais: `operacoes` (só linhas com essas operações), `quando` (condição da linguagem de regras v2, só
Outros e só tem/soma/texto/combinadores; em Cabos só `texto`) e `operacao_nova` — a operação da linha só muda quando a ação
a determina explicitamente (e apenas nas linhas em que a ação de fato alterou algo); sem ele a operação é mantida.

GARANTIAS
  • Camada 1 (contrato) é a guarda: uma edição/inserção que deixe a linha com um ERRO de contrato novo é descartada e
    listada em `descartadas` (a linha fica como estava).
  • Idempotência: reaplicar o resultado não produz novas operações (nas ações acima, desde que `para` não contenha `de`).
  • O resultado traz as tabelas finais (com `id` por linha) e `operacoes` (editar/inserir/excluir/mover) para a pré-visualização.
"""
import copy
import re

from services.orcamento_calc import extrair_ativos_outros
from services.regras_dominio import (_casa, _cond, _frase_sel, _frase_situacao, _Linha, _n, _numero_ok,
                                     _regex_ok, _validar_cond, _validar_sel, validar_grupos)
from services.validacao_planilhas import OPERACOES_VALIDAS, validar_planilhas

LIMITE_ACOES = 50
LIMITE_LINHAS = 5000
TABELAS = ("cabos", "outros", "ambos")
COMUNS = {"acao", "id", "descricao", "tabela"}
ESPECIFICOS = {
    "substituir": {"de", "para", "modo", "palavra_inteira", "operacao_nova", "operacoes", "quando"},
    "normalizar": {"regras", "operacoes", "quando"},
    "ordenar": {"por"},
    "excluir_linhas": {"onde", "operacoes", "quando"},
    "adicionar_linha": {"valores", "posicao", "apenas_se_nao_existir"},
    "adicionar_ativo": {"ativo", "qtd", "se_ja_existe", "operacao_nova", "operacoes", "quando"},
    "remover_ativo": {"ativo", "operacao_nova", "operacoes", "quando"},
    "mesclar_duplicadas": {"operacao_nova", "operacoes", "quando"},
}
SO_OUTROS = {"adicionar_ativo", "remover_ativo", "mesclar_duplicadas"}
SO_UMA_TABELA = {"adicionar_linha"}
REGRAS_NORMALIZAR = ("espacos", "maiusculas", "poste")
COLUNAS = ("operacao", "ativo", "entidade")
_RE_TOKEN = re.compile(r"\S+")
_RE_QTD_ATIVO = re.compile(r"^(\*?[\d.,]+)-(.+)$")
_RE_NUM = re.compile(r"^[\d]+([.,]\d+)?$")


# ═══════════════════════════════════════════════════════════════════════════
# VALIDAÇÃO DO SCHEMA
# ═══════════════════════════════════════════════════════════════════════════
def _usa_itens(c) -> bool:
    """A condição olha itens <qtd>-<ativo> (tem/soma)? Isso só existe em Outros."""
    if isinstance(c, dict):
        if "tem" in c or "soma" in c:
            return True
        return any(_usa_itens(v) for k, v in c.items() if k in ("todos", "algum", "nenhum", "nao"))
    if isinstance(c, list):
        return any(_usa_itens(x) for x in c)
    return False


def _validar_condicao(c, campo, tabela, erros, grupos):
    _validar_cond(c, campo, erros, grupos, [0])
    if tabela in ("cabos", "ambos") and _usa_itens(c):
        erros.append(f"{campo}: em Cabos só vale a condição 'texto' (tem/soma olham itens de Outros).")


def _validar_acao(a, pre, erros, grupos):
    if not isinstance(a, dict):
        erros.append(f"{pre}: precisa ser um objeto.")
        return
    nome = a.get("acao")
    if nome not in ESPECIFICOS:
        erros.append(f"{pre}: 'acao' precisa ser uma de {sorted(ESPECIFICOS)}.")
        return
    pre = f"{pre} ({nome})"
    for k in a:
        if k not in COMUNS and k not in ESPECIFICOS[nome]:
            erros.append(f"{pre}: campo desconhecido ou não permitido '{k}'.")
    tabela = a.get("tabela")
    modo = a.get("modo", "texto") if nome == "substituir" else None
    permitidas = TABELAS
    if nome in SO_OUTROS or (nome == "substituir" and modo == "item"):
        permitidas = ("outros",)
    elif nome in SO_UMA_TABELA:
        permitidas = ("cabos", "outros")
    if tabela not in permitidas:
        erros.append(f"{pre}: 'tabela' precisa ser {' ou '.join(permitidas)}.")
        tabela = None
    if "operacoes" in a:
        ops = a["operacoes"]
        if not isinstance(ops, list) or not ops or any(str(o).strip().upper() not in OPERACOES_VALIDAS for o in ops):
            erros.append(f"{pre}: 'operacoes' precisa ser uma lista com valores de {sorted(OPERACOES_VALIDAS)}.")
    if "operacao_nova" in a and str(a["operacao_nova"]).strip().upper() not in OPERACOES_VALIDAS:
        erros.append(f"{pre}: 'operacao_nova' precisa ser um de {sorted(OPERACOES_VALIDAS)}.")
    if a.get("quando") is not None:
        _validar_condicao(a["quando"], f"{pre}.quando", tabela or "ambos", erros, grupos)

    if nome == "substituir":
        if modo not in ("texto", "item"):
            erros.append(f"{pre}: 'modo' precisa ser 'texto' ou 'item'.")
        de, para = a.get("de"), a.get("para")
        if not isinstance(para, str):
            erros.append(f"{pre}: 'para' precisa ser um texto.")
        elif modo == "item" and (not para.strip() or " " in para.strip()):
            erros.append(f"{pre}: em modo item, 'para' precisa ser um código sem espaços.")
        if modo == "item":
            _validar_sel(de, f"{pre}.de", erros, grupos)
        elif isinstance(de, dict):
            if set(de) != {"regex"}:
                erros.append(f"{pre}.de: use um texto ou {{\"regex\": \"...\"}}.")
            else:
                _regex_ok(de["regex"], f"{pre}.de.regex", erros)
                try:
                    re.compile(de["regex"]).sub(para if isinstance(para, str) else "", "")
                except (re.error, IndexError) as e:
                    erros.append(f"{pre}.para: referência inválida ({e}).")
        elif not isinstance(de, str) or not de:
            erros.append(f"{pre}.de: texto obrigatório (ou {{\"regex\": ...}}).")
        if "palavra_inteira" in a and not isinstance(a["palavra_inteira"], bool):
            erros.append(f"{pre}: 'palavra_inteira' precisa ser true ou false.")
    elif nome == "normalizar":
        regras = a.get("regras", ["espacos"])
        if not isinstance(regras, list) or not regras or any(r not in REGRAS_NORMALIZAR for r in regras):
            erros.append(f"{pre}: 'regras' precisa ser uma lista de {list(REGRAS_NORMALIZAR)}.")
    elif nome == "ordenar":
        por = a.get("por")
        if not isinstance(por, list) or not por:
            erros.append(f"{pre}: 'por' precisa ser uma lista de critérios.")
        else:
            for i, c in enumerate(por):
                if not isinstance(c, dict) or c.get("coluna") not in COLUNAS or set(c) - {"coluna", "ordem", "valores"}:
                    erros.append(f"{pre}.por[{i}]: use {{\"coluna\": {'|'.join(COLUNAS)}, \"ordem\": asc|desc, \"valores\": [..]}}.")
                    continue
                if c.get("ordem", "asc") not in ("asc", "desc"):
                    erros.append(f"{pre}.por[{i}].ordem: 'asc' ou 'desc'.")
                v = c.get("valores")
                if v is not None and (not isinstance(v, list) or not v or not all(isinstance(x, str) for x in v)):
                    erros.append(f"{pre}.por[{i}].valores: lista de textos.")
    elif nome == "excluir_linhas":
        onde = a.get("onde")
        chaves = [k for k in ("vazias", "duplicadas", "condicao", "texto") if isinstance(onde, dict) and k in onde]
        if not isinstance(onde, dict) or len(chaves) != 1 or len(onde) != 1:
            erros.append(f"{pre}: 'onde' precisa ter exatamente um de vazias, duplicadas, condicao, texto.")
        else:
            k = chaves[0]
            if k in ("vazias", "duplicadas") and onde[k] is not True:
                erros.append(f"{pre}.onde.{k}: use true.")
            elif k == "condicao":
                _validar_condicao(onde[k], f"{pre}.onde.condicao", tabela or "ambos", erros, grupos)
            elif k == "texto":
                _regex_ok(onde[k], f"{pre}.onde.texto", erros)
    elif nome == "adicionar_linha":
        v = a.get("valores")
        if not isinstance(v, dict) or set(v) - {"operacao", "ativo", "entidade"}:
            erros.append(f"{pre}: 'valores' precisa ser {{operacao, ativo, entidade?}}.")
        else:
            if str(v.get("operacao", "")).strip().upper() not in OPERACOES_VALIDAS:
                erros.append(f"{pre}.valores.operacao: um de {sorted(OPERACOES_VALIDAS)}.")
            if not isinstance(v.get("ativo"), str) or not v["ativo"].strip():
                erros.append(f"{pre}.valores.ativo: texto obrigatório.")
        pos = a.get("posicao", "fim")
        if isinstance(pos, dict):
            if len(pos) != 1 or next(iter(pos)) not in ("depois_de", "antes_de"):
                erros.append(f"{pre}.posicao: use 'fim', 'inicio', {{\"depois_de\": COND}} ou {{\"antes_de\": COND}}.")
            else:
                _validar_condicao(next(iter(pos.values())), f"{pre}.posicao", tabela or "outros", erros, grupos)
        elif pos not in ("fim", "inicio"):
            erros.append(f"{pre}.posicao: use 'fim', 'inicio', {{\"depois_de\": COND}} ou {{\"antes_de\": COND}}.")
        if "apenas_se_nao_existir" in a and not isinstance(a["apenas_se_nao_existir"], bool):
            erros.append(f"{pre}: 'apenas_se_nao_existir' precisa ser true ou false.")
    elif nome == "adicionar_ativo":
        ativo = a.get("ativo")
        if not isinstance(ativo, str) or not ativo.strip() or re.search(r"[\s]", ativo.strip()):
            erros.append(f"{pre}: 'ativo' precisa ser um código sem espaços.")
        if "qtd" in a and not (_numero_ok(a["qtd"]) and a["qtd"] != 0):
            erros.append(f"{pre}: 'qtd' precisa ser um número diferente de zero (negativo gera o prefixo '*', linha viva).")
        if a.get("se_ja_existe", "ignorar") not in ("ignorar", "somar", "substituir"):
            erros.append(f"{pre}: 'se_ja_existe' precisa ser ignorar, somar ou substituir.")
    elif nome == "remover_ativo":
        _validar_sel(a.get("ativo"), f"{pre}.ativo", erros, grupos)


def validar_acoes(acoes, grupos=None, limite=LIMITE_ACOES) -> list:
    """Lista de erros (vazia = válido). `grupos` = grupos efetivos (para @GRUPO). `limite` = máximo de ações (50 nas rotas; o modo autônomo usa mais)."""
    if not isinstance(acoes, list):
        return ["O payload de ações precisa ser uma lista."]
    erros = validar_grupos(grupos) if grupos else []
    if len(acoes) > limite:
        erros.append(f"Ações demais (máx. {limite}).")
    for i, a in enumerate(acoes):
        _validar_acao(a, f"Ação #{i + 1}", erros, grupos or {})
    return erros


_RE_ID_RECEITA = re.compile(r"^[A-Za-z0-9_.\-]{1,60}$")
CAMPOS_RECEITA = {"id", "nome", "descricao", "ativa", "acoes", "regras", "origem", "oculta", "frases"}


def validar_receitas(receitas, grupos=None) -> list:
    """Erros (vazia = válido) do cadastro de ajustes: lista de receitas {id, nome, ativa, acoes, regras?} (TASK-023)."""
    if not isinstance(receitas, list):
        return ["O payload de ajustes precisa ser uma lista."]
    erros, ids = [], set()
    for i, r in enumerate(receitas):
        pre = f"Ajuste #{i + 1}"
        if not isinstance(r, dict):
            erros.append(f"{pre}: precisa ser um objeto.")
            continue
        rid = r.get("id")
        if not isinstance(rid, str) or not _RE_ID_RECEITA.match(rid):
            erros.append(f"{pre}: 'id' obrigatório (letras, números, _ . -; até 60).")
        elif rid in ids:
            erros.append(f"{pre}: 'id' duplicado ({rid}).")
        else:
            ids.add(rid)
            pre = f"Ajuste {rid}"
        for k in r:
            if k not in CAMPOS_RECEITA:
                erros.append(f"{pre}: campo desconhecido '{k}'.")
        if not isinstance(r.get("nome"), str) or not r["nome"].strip() or len(r["nome"]) > 120:
            erros.append(f"{pre}: 'nome' obrigatório (até 120 caracteres).")
        if not isinstance(r.get("ativa"), bool):
            erros.append(f"{pre}: 'ativa' precisa ser true ou false.")
        if "descricao" in r and not isinstance(r["descricao"], str):
            erros.append(f"{pre}: 'descricao' precisa ser um texto.")
        if "regras" in r and (not isinstance(r["regras"], list) or any(not isinstance(x, str) or not x.strip() for x in r["regras"])):
            erros.append(f"{pre}: 'regras' precisa ser uma lista de ids de regras.")
        acoes = r.get("acoes")
        if not isinstance(acoes, list) or not acoes:
            erros.append(f"{pre}: 'acoes' precisa ser uma lista com ao menos uma ação.")
        else:
            erros += [f"{pre}: {e}" for e in validar_acoes(acoes, grupos)]
    return erros


# ═══════════════════════════════════════════════════════════════════════════
# AUXILIARES DE TEXTO
# ═══════════════════════════════════════════════════════════════════════════
def _parse_token(tok: str):
    """'1-SUPL' → ('1', 'SUPL'); 'DT11/300' → (None, 'DT11/300')."""
    m = _RE_QTD_ATIVO.match(tok)
    return (m.group(1), m.group(2)) if m else (None, tok)


def _juntar(qtd, ativo):
    return f"{qtd}-{ativo}" if qtd is not None else ativo


def _num(qtd):
    """'1' → 1.0; '*1' → -1.0 (negativo/linha viva, mesma leitura de orcamento_calc.tokenizar_outros)."""
    if qtd is None:
        return None
    neg, s = (True, qtd[1:]) if qtd.startswith("*") else (False, qtd)
    if not _RE_NUM.match(s):
        return None
    v = float(s.replace(",", "."))
    return -v if neg else v


def _fmt_qtd(x):
    """Inverso de _num: quantidade negativa volta como '*<abs>' (nunca '-<abs>', que o tokenizador não lê como sinal)."""
    return f"*{_n(-x)}" if x < 0 else _n(x)


def _chave_linha(operacao, ativo):
    return ((operacao or "").strip().upper(), re.sub(r"\s+", " ", (ativo or "").strip().upper()))


def _contexto(tabela, texto, grupos):
    itens = extrair_ativos_outros({"ativo": texto, "operacao": ""}) if tabela == "outros" else []
    return _Linha(itens, texto, grupos)


def _quando(cond, tabela, texto, grupos) -> bool:
    return cond is None or _cond(cond, _contexto(tabela, texto, grupos))[0]


def _natural(s):
    return [int(t) if t.isdigit() else t for t in re.split(r"(\d+)", (s or "").lower())]


# ═══════════════════════════════════════════════════════════════════════════
# TRANSFORMAÇÕES DE TEXTO (uma por ação; devolvem o novo texto)
# ═══════════════════════════════════════════════════════════════════════════
def _t_substituir(a, tabela, texto, grupos):
    de, para = a["de"], a["para"]
    if a.get("modo", "texto") == "item":
        def f(m):
            qtd, ativo = _parse_token(m.group(0))
            return _juntar(qtd, para) if _casa(ativo.upper(), de, grupos) else m.group(0)
        return _RE_TOKEN.sub(f, texto)
    if isinstance(de, dict):
        return re.compile(de["regex"], re.IGNORECASE).sub(para, texto)
    padrao = re.escape(de)
    if a.get("palavra_inteira", True):
        padrao = rf"(?<![A-Za-z0-9]){padrao}(?![A-Za-z0-9])"
    return re.compile(padrao, re.IGNORECASE).sub(lambda m: para, texto)


def _t_normalizar(a, tabela, texto, grupos):
    for r in a.get("regras", ["espacos"]):
        if r == "espacos":
            texto = re.sub(r"\s+", " ", texto).strip()
        elif r == "maiusculas":
            texto = texto.upper()
        elif r == "poste":
            texto = re.sub(r"^(\s*)1-(?=(?:DT|CV))", r"\1", texto, flags=re.IGNORECASE)
    return texto


def _t_adicionar_ativo(a, tabela, texto, grupos):
    ativo, qtd = a["ativo"].strip(), a.get("qtd", 1)
    modo = a.get("se_ja_existe", "ignorar")
    achou = False

    def f(m):
        nonlocal achou
        q, nome = _parse_token(m.group(0))
        if q is None or nome.upper() != ativo.upper() or achou:
            return m.group(0)
        achou = True
        atual = _num(q)
        if modo == "somar" and atual is not None:
            return _juntar(_fmt_qtd(atual + qtd), nome)
        if modo == "substituir":
            return _juntar(_fmt_qtd(qtd), nome)
        return m.group(0)
    novo = _RE_TOKEN.sub(f, texto)
    if achou:
        return novo
    return f"{texto.rstrip()} {_fmt_qtd(qtd)}-{ativo}".strip()


def _t_remover_ativo(a, tabela, texto, grupos):
    def f(m):
        _, ativo = _parse_token(m.group(0))
        return "" if _casa(ativo.upper(), a["ativo"], grupos) else m.group(0)
    novo = _RE_TOKEN.sub(f, texto)
    return re.sub(r"\s+", " ", novo).strip() if novo != texto else texto


def _t_mesclar(a, tabela, texto, grupos):
    tokens = [(m, *_parse_token(m.group(0))) for m in _RE_TOKEN.finditer(texto)]
    soma, primeira, contagem = {}, {}, {}
    for m, q, ativo in tokens:
        n = _num(q)
        if n is not None:
            k = ativo.upper()
            soma[k] = soma.get(k, 0) + n
            contagem[k] = contagem.get(k, 0) + 1
            primeira.setdefault(k, m.start())
    if not any(c > 1 for c in contagem.values()):
        return texto

    def f(m):
        q, ativo = _parse_token(m.group(0))
        k = ativo.upper()
        if _num(q) is None or contagem.get(k, 0) < 2:
            return m.group(0)
        return _juntar(_n(soma[k]), ativo) if m.start() == primeira[k] else ""
    return re.sub(r"\s+", " ", _RE_TOKEN.sub(f, texto)).strip()


TRANSFORMACOES = {
    "substituir": _t_substituir, "normalizar": _t_normalizar, "adicionar_ativo": _t_adicionar_ativo,
    "remover_ativo": _t_remover_ativo, "mesclar_duplicadas": _t_mesclar,
}


# ═══════════════════════════════════════════════════════════════════════════
# MOTOR
# ═══════════════════════════════════════════════════════════════════════════
def _erros_c1(tabela, row) -> dict:
    linha = {"id": row["id"], "operacao": row.get("operacao", ""), "ativo": row.get("ativo", "")}
    achados = validar_planilhas([linha], []) if tabela == "cabos" else validar_planilhas([], [linha])
    return {a["regra_id"]: a["mensagem"] for a in achados if a["severidade"] == "erro"}


def _novo_erro_c1(tabela, antes, depois):
    """Mensagem do primeiro erro de contrato que a linha NÃO tinha e passou a ter; None se a mudança é segura."""
    ea, ed = _erros_c1(tabela, antes), _erros_c1(tabela, depois)
    for rid, msg in ed.items():
        if rid not in ea:
            return f"{rid}: {msg}"
    return None


class _Estado:
    def __init__(self, cabos, outros, grupos):
        self.grupos = grupos or {}
        self.tabelas = {"cabos": self._com_ids(cabos, "CABOS"), "outros": self._com_ids(outros, "OUTROS")}
        self.originais = {t: copy.deepcopy(r) for t, r in self.tabelas.items()}
        self.tocadas = {"cabos": {}, "outros": {}}
        self.descartadas, self.avisos = [], []
        self.novas = 0

    @staticmethod
    def _com_ids(linhas, prefixo):
        saida = []
        for i, l in enumerate(linhas or []):
            r = copy.deepcopy(l)
            r["id"] = l.get("id") or f"{prefixo}-{i}"
            r["operacao"] = r.get("operacao") or ""
            r["ativo"] = r.get("ativo") or ""
            saida.append(r)
        return saida

    def marcar(self, tabela, linha_id, idx):
        self.tocadas[tabela].setdefault(linha_id, set()).add(idx)

    def descartar(self, idx, linha_id, motivo):
        self.descartadas.append({"acao": idx, "linha_id": linha_id, "motivo": motivo})


def _ops_filtro(a):
    return {str(o).strip().upper() for o in a["operacoes"]} if a.get("operacoes") else None


def _alvos(a):
    return ("cabos", "outros") if a["tabela"] == "ambos" else (a["tabela"],)


def _aplicar_edicao(est, idx, a):
    fn = TRANSFORMACOES[a["acao"]]
    filtro = _ops_filtro(a)
    for tabela in _alvos(a):
        for row in est.tabelas[tabela]:
            texto = row["ativo"]
            if not texto.strip():
                continue
            if filtro is not None and row["operacao"].strip().upper() not in filtro:
                continue
            if not _quando(a.get("quando"), tabela, texto, est.grupos):
                continue
            novo = fn(a, tabela, texto, est.grupos)
            if novo == texto:
                continue
            cand = {**row, "ativo": novo}
            if a.get("operacao_nova"):
                cand["operacao"] = str(a["operacao_nova"]).strip().upper()
            erro = _novo_erro_c1(tabela, row, cand)
            if erro:
                est.descartar(idx, row["id"], f"o ajuste deixaria a linha inválida ({erro})")
                continue
            row.update(cand)
            est.marcar(tabela, row["id"], idx)


def _aplicar_ordenar(est, idx, a):
    for tabela in _alvos(a):
        linhas = est.tabelas[tabela]
        if tabela == "cabos":
            est.avisos.append("Ordenar Cabos pode mudar o cálculo: linha de cabo sem fase herda a fase da linha seguinte.")
        for crit in reversed(a["por"]):
            col, desc, valores = crit["coluna"], crit.get("ordem", "asc") == "desc", crit.get("valores")
            if valores:
                pos = {v.strip().upper(): i for i, v in enumerate(valores)}
                chave = lambda r, c=col, p=pos: p.get(str(r.get(c, "")).strip().upper(), len(p))
            else:
                chave = lambda r, c=col: _natural(str(r.get(c, "")))
            linhas.sort(key=chave, reverse=desc)


def _aplicar_excluir(est, idx, a):
    onde, filtro = a["onde"], _ops_filtro(a)
    for tabela in _alvos(a):
        vistos, manter = set(), []
        for row in est.tabelas[tabela]:
            texto = row["ativo"]
            apaga = False
            if filtro is None or row["operacao"].strip().upper() in filtro:
                aplica = _quando(a.get("quando"), tabela, texto, est.grupos) if texto.strip() else a.get("quando") is None
                if aplica:
                    if onde.get("vazias"):
                        apaga = not texto.strip()
                    elif onde.get("duplicadas"):
                        if texto.strip():
                            k = _chave_linha(row["operacao"], texto)
                            apaga = k in vistos
                            vistos.add(k)
                    elif "condicao" in onde:
                        apaga = bool(texto.strip()) and _quando(onde["condicao"], tabela, texto, est.grupos)
                    else:
                        apaga = bool(re.compile(onde["texto"], re.IGNORECASE).search(texto))
            if apaga:
                est.marcar(tabela, row["id"], idx)
            else:
                manter.append(row)
        est.tabelas[tabela] = manter


def _aplicar_adicionar_linha(est, idx, a):
    tabela = a["tabela"]
    v = a["valores"]
    pos = a.get("posicao", "fim")
    unico = a.get("apenas_se_nao_existir", True)
    linhas = est.tabelas[tabela]
    chave_nova = _chave_linha(v["operacao"], v["ativo"])

    def criar():
        est.novas += 1
        return {"id": f"NOVA-{est.novas}", "operacao": str(v["operacao"]).strip().upper(), "ativo": v["ativo"].strip(),
                "entidade": v.get("entidade", "0")}

    def tentar(indice, vizinhos):
        if unico and any(_chave_linha(r["operacao"], r["ativo"]) == chave_nova for r in vizinhos):
            return False
        nova = criar()
        erro = _novo_erro_c1(tabela, {"id": nova["id"], "operacao": "I", "ativo": ""}, nova)
        if erro:
            est.descartar(idx, nova["id"], f"a linha nova seria inválida ({erro})")
            return False
        est.tabelas[tabela].insert(indice, nova)
        est.marcar(tabela, nova["id"], idx)
        return True

    if pos in ("fim", "inicio"):
        tentar(len(linhas) if pos == "fim" else 0, linhas)
        return
    modo, cond = next(iter(pos.items()))
    for ref in [r["id"] for r in list(linhas)]:
        atual = est.tabelas[tabela]
        i = next(k for k, r in enumerate(atual) if r["id"] == ref)
        if not atual[i]["ativo"].strip() or not _quando(cond, tabela, atual[i]["ativo"], est.grupos):
            continue
        if modo == "depois_de":
            vizinho = atual[i + 1:i + 2]
            tentar(i + 1, vizinho)
        else:
            vizinho = atual[i - 1:i] if i > 0 else []
            tentar(i, vizinho)


EXECUTORES = {
    "ordenar": _aplicar_ordenar, "excluir_linhas": _aplicar_excluir, "adicionar_linha": _aplicar_adicionar_linha,
}


def _visivel(r):
    v = {"operacao": r.get("operacao", ""), "ativo": r.get("ativo", "")}
    if "entidade" in r:
        v["entidade"] = r["entidade"]
    return v


def _diff(tabela, originais, finais, tocadas):
    ops = []
    o_por_id = {r["id"]: r for r in originais}
    f_por_id = {r["id"]: r for r in finais}
    anterior = None
    for r in finais:
        if r["id"] not in o_por_id:
            ops.append({"op": "inserir", "tabela": tabela, "linha_id": r["id"], "depois": _visivel(r),
                        "depois_de": anterior, "acoes": sorted(tocadas.get(r["id"], ()))})
        anterior = r["id"]
    for r in originais:
        f = f_por_id.get(r["id"])
        if f is None:
            ops.append({"op": "excluir", "tabela": tabela, "linha_id": r["id"], "antes": _visivel(r),
                        "acoes": sorted(tocadas.get(r["id"], ()))})
        elif _visivel(r) != _visivel(f):
            ops.append({"op": "editar", "tabela": tabela, "linha_id": r["id"], "antes": _visivel(r),
                        "depois": _visivel(f), "acoes": sorted(tocadas.get(r["id"], ()))})
    antes = [r["id"] for r in originais if r["id"] in f_por_id]
    depois = [r["id"] for r in finais if r["id"] in o_por_id]
    if antes != depois:
        ops.append({"op": "mover", "tabela": tabela, "ordem_antes": antes, "ordem_depois": depois,
                    "acoes": []})
    return ops


def ajustar(acoes: list, cabos: list, outros: list, grupos: dict = None, limite_acoes: int = LIMITE_ACOES) -> dict:
    """Aplica as ações em ordem (cada uma vê o resultado da anterior) e devolve o diff. Não altera os argumentos."""
    erros = validar_acoes(acoes, grupos, limite_acoes)
    if erros:
        raise ValueError(erros)
    if len(cabos or []) > LIMITE_LINHAS or len(outros or []) > LIMITE_LINHAS:
        raise ValueError([f"Tabelas grandes demais (máx. {LIMITE_LINHAS} linhas cada)."])
    est = _Estado(cabos, outros, grupos)
    for idx, a in enumerate(acoes):
        if a["acao"] in TRANSFORMACOES:
            _aplicar_edicao(est, idx, a)
        else:
            EXECUTORES[a["acao"]](est, idx, a)
        if a["acao"] == "ordenar":
            for t in _alvos(a):
                est.marcar(t, "__ordem__", idx)
    operacoes = []
    for t in ("cabos", "outros"):
        ops = _diff(t, est.originais[t], est.tabelas[t], est.tocadas[t])
        for o in ops:
            if o["op"] == "mover":
                o["acoes"] = sorted(est.tocadas[t].get("__ordem__", ()))
        operacoes += ops
    resumo = {k: sum(1 for o in operacoes if o["op"] == k) for k in ("editar", "inserir", "excluir", "mover")}
    return {"operacoes": operacoes, "cabos": est.tabelas["cabos"], "outros": est.tabelas["outros"],
            "descartadas": est.descartadas, "avisos": list(dict.fromkeys(est.avisos)), "resumo": resumo}


# ═══════════════════════════════════════════════════════════════════════════
# DESCRIÇÃO EM PORTUGUÊS (editor e pré-visualização)
# ═══════════════════════════════════════════════════════════════════════════
_NOME_TABELA = {"cabos": "Cabos", "outros": "Outros", "ambos": "Cabos e Outros"}


def _frase_filtros(a):
    partes = []
    if a.get("operacoes"):
        partes.append("com operação " + ", ".join(str(o).upper() for o in a["operacoes"]))
    if a.get("quando"):
        partes.append(f"em que a linha {_frase_situacao(a['quando'])}")
    return f" (só nas linhas {' e '.join(partes)})" if partes else ""


def descrever_acao(a: dict) -> str:
    """Frase em português da ação (supõe ação válida)."""
    t = _NOME_TABELA.get(a.get("tabela"), "?")
    nome = a["acao"]
    op = f" e mudar a operação para {str(a['operacao_nova']).upper()}" if a.get("operacao_nova") else ""
    if nome == "substituir":
        de = a["de"]
        de_txt = _frase_sel(de) if a.get("modo") == "item" else (f"/{de['regex']}/" if isinstance(de, dict) else f"'{de}'")
        return f"Em {t}: substituir {de_txt} por '{a['para']}'{_frase_filtros(a)}{op}."
    if nome == "normalizar":
        rot = {"espacos": "espaços", "maiusculas": "maiúsculas", "poste": "poste sem '1-' no início"}
        return f"Em {t}: normalizar ({', '.join(rot[r] for r in a.get('regras', ['espacos']))}){_frase_filtros(a)}."
    if nome == "ordenar":
        crit = ", ".join(f"{c['coluna']} {'decrescente' if c.get('ordem') == 'desc' else 'crescente'}" for c in a["por"])
        return f"Em {t}: ordenar por {crit}."
    if nome == "excluir_linhas":
        onde = a["onde"]
        if onde.get("vazias"):
            alvo = "linhas vazias"
        elif onde.get("duplicadas"):
            alvo = "linhas duplicadas (mantém a primeira)"
        elif "condicao" in onde:
            alvo = f"linhas em que a linha {_frase_situacao(onde['condicao'])}"
        else:
            alvo = f"linhas cujo texto casa com /{onde['texto']}/"
        return f"Em {t}: excluir {alvo}{_frase_filtros(a)}."
    if nome == "adicionar_linha":
        v = a["valores"]
        pos = a.get("posicao", "fim")
        onde = {"fim": "no fim", "inicio": "no início"}.get(pos) if isinstance(pos, str) else \
            ("depois" if "depois_de" in pos else "antes") + f" de cada linha em que a linha {_frase_situacao(next(iter(pos.values())))}"
        unico = ", só se ainda não existir" if a.get("apenas_se_nao_existir", True) else ""
        return f"Em {t}: adicionar a linha {str(v['operacao']).upper()} '{v['ativo']}' {onde}{unico}."
    if nome == "adicionar_ativo":
        modo = {"ignorar": "", "somar": "; se já existir, soma a quantidade", "substituir": "; se já existir, troca a quantidade"}[a.get("se_ja_existe", "ignorar")]
        return f"Em {t}: adicionar {_fmt_qtd(a.get('qtd', 1))}-{a['ativo']} à linha{_frase_filtros(a)}{modo}{op}."
    if nome == "remover_ativo":
        return f"Em {t}: remover {_frase_sel(a['ativo'])} da linha{_frase_filtros(a)}{op}."
    return f"Em {t}: somar as quantidades do mesmo ativo repetido na linha{_frase_filtros(a)}{op}."
