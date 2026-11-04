"""
services/fator_formula.py — Fator da base de orçamento em função de parâmetro (TASK-109).

FATOR I / FATOR R da base continuam aceitando um número (``1,5``) e agora também uma FÓRMULA que começa com ``=`` e usa os parâmetros do pedido:

    =NUM(fios)           "8F" → 8        =NUM(fios)/8        =SE(tipo="A"; 2; 1)        =MAX(NUM(cabo); 35) * 2

Linguagem (sem ``eval``): números (``.`` ou ``,`` decimal), textos entre aspas duplas, nomes de parâmetro (sem diferença de caixa; ``CABO#2`` = 2º ``cabo``),
``+ - * /``, parênteses, comparações ``= <> < > <= >=`` e as funções ``NUM(x)`` (1º número do texto), ``ABS``, ``MAX``, ``MIN``, ``ARRED(x; casas)``, ``CIMA``,
``BAIXO``, ``SE(cond; a; b)``, ``E``, ``OU`` e ``NAO``. Argumentos separados por ``;``. Parâmetro de texto só entra em conta se for número puro (``8``);
use ``NUM(nome)`` para ``8F``. Erro (parâmetro não informado, texto sem número, divisão por zero) → ``FormulaErro``; o cálculo avisa e a linha não soma.
"""
import functools
import math
import re

TOKEN = re.compile(r'\s*(?:(?P<num>\d+(?:[.,]\d+)?)|(?P<str>"[^"]*")|(?P<id>[A-Za-z_][A-Za-z0-9_.]*(?:#\d+)?)|(?P<op><>|<=|>=|[-+*/()=<>;]))')


class FormulaErro(ValueError):
    pass


def _tokens(txt):
    saida, pos = [], 0
    txt = txt.rstrip()
    while pos < len(txt):
        m = TOKEN.match(txt, pos)
        if not m or m.end() == pos:
            raise FormulaErro(f"caractere inesperado em '{txt[pos:pos + 10].strip()}'")
        pos = m.end()
        for tipo in ("num", "str", "id", "op"):
            if m.group(tipo) is not None:
                saida.append((tipo, m.group(tipo)))
                break
    return saida


class _Parser:
    def __init__(self, tokens):
        self.t, self.i = tokens, 0

    def ver(self):
        return self.t[self.i] if self.i < len(self.t) else (None, None)

    def pegar(self, esperado=None):
        tipo, val = self.ver()
        if tipo is None:
            raise FormulaErro("fórmula incompleta")
        if esperado is not None and val != esperado:
            raise FormulaErro(f"esperado '{esperado}' e veio '{val}'")
        self.i += 1
        return tipo, val

    def expr(self):
        e = self.soma()
        tipo, val = self.ver()
        if tipo == "op" and val in ("=", "<>", "<", ">", "<=", ">="):
            self.pegar()
            e = ("cmp", val, e, self.soma())
        return e

    def soma(self):
        e = self.produto()
        while self.ver()[1] in ("+", "-") and self.ver()[0] == "op":
            op = self.pegar()[1]
            e = ("bin", op, e, self.produto())
        return e

    def produto(self):
        e = self.unario()
        while self.ver()[1] in ("*", "/") and self.ver()[0] == "op":
            op = self.pegar()[1]
            e = ("bin", op, e, self.unario())
        return e

    def unario(self):
        if self.ver() == ("op", "-"):
            self.pegar()
            return ("neg", self.unario())
        return self.primario()

    def primario(self):
        tipo, val = self.pegar()
        if tipo == "num":
            return ("num", float(val.replace(",", ".")))
        if tipo == "str":
            return ("str", val[1:-1])
        if tipo == "id":
            if self.ver() == ("op", "("):
                self.pegar()
                args = []
                if self.ver() != ("op", ")"):
                    args.append(self.expr())
                    while self.ver() == ("op", ";"):
                        self.pegar()
                        args.append(self.expr())
                self.pegar(")")
                nome = val.upper()
                if nome not in FUNCOES:
                    raise FormulaErro(f"função desconhecida '{val}'")
                return ("fn", nome, args)
            return ("id", val.upper())
        if val == "(":
            e = self.expr()
            self.pegar(")")
            return e
        raise FormulaErro(f"inesperado '{val}'")


def _corpo(formula):
    f = str(formula if formula is not None else "").strip()
    if f.startswith("="):
        f = f[1:]
    if not f.strip():
        raise FormulaErro("fórmula vazia")
    return f


@functools.lru_cache(maxsize=2048)
def compilar(formula):
    """Fórmula ("=NUM(fios)/8" ou sem o "=") → árvore. Sintaxe inválida → FormulaErro."""
    p = _Parser(_tokens(_corpo(formula)))
    ast = p.expr()
    if p.i != len(p.t):
        raise FormulaErro(f"sobrou '{p.ver()[1]}'")
    return ast


def identificadores(formula) -> set:
    """Nomes de parâmetro usados (maiúsculos, sem '#2'), para a validação da base."""
    saida = set()

    def anda(n):
        if n[0] == "id":
            saida.add(n[1].split("#")[0])
        elif n[0] in ("bin", "cmp"):
            anda(n[2]); anda(n[3])
        elif n[0] == "neg":
            anda(n[1])
        elif n[0] == "fn":
            for a in n[2]:
                anda(a)
    anda(compilar(formula))
    return saida


def _num(v, nome=None):
    if isinstance(v, str):
        try:
            return float(v.strip().replace(",", "."))
        except ValueError:
            raise FormulaErro(f"'{v}' não é número" + (f" (parâmetro {nome}); use NUM({nome.lower()})" if nome else ""))
    return float(v)


def _fn_num(args):
    m = re.search(r"\d+(?:[.,]\d+)?", str(args[0]))
    if not m:
        raise FormulaErro(f"NUM: '{args[0]}' não tem número")
    return float(m.group(0).replace(",", "."))


def _arred(args):
    x, casas = _num(args[0]), int(_num(args[1])) if len(args) > 1 else 0
    f = 10 ** casas
    return float(math.floor(round(x * f, 9) + 0.5) / f)


FUNCOES = {
    "NUM": (1, 1, lambda a: _fn_num(a)),
    "ABS": (1, 1, lambda a: abs(_num(a[0]))),
    "MAX": (1, 99, lambda a: max(_num(x) for x in a)),
    "MIN": (1, 99, lambda a: min(_num(x) for x in a)),
    "ARRED": (1, 2, _arred),
    "CIMA": (1, 1, lambda a: float(math.ceil(round(_num(a[0]), 9)))),
    "BAIXO": (1, 1, lambda a: float(math.floor(round(_num(a[0]), 9)))),
    "NAO": (1, 1, lambda a: 0.0 if _verdade(a[0]) else 1.0),
}


def _verdade(v):
    return bool(v) if not isinstance(v, str) else v.strip() != ""


def _ev(n, params):
    tipo = n[0]
    if tipo == "num" or tipo == "str":
        return n[1]
    if tipo == "id":
        chave = n[1]
        if chave in params:
            return params[chave]
        if "#" not in chave and chave + "#1" in params:
            return params[chave + "#1"]
        raise FormulaErro(f"parâmetro {chave.split('#')[0].lower()} não informado no pedido")
    if tipo == "neg":
        return -_num(_ev(n[1], params))
    if tipo == "bin":
        a, b = _num(_ev(n[2], params), _nome(n[2])), _num(_ev(n[3], params), _nome(n[3]))
        if n[1] == "+":
            return a + b
        if n[1] == "-":
            return a - b
        if n[1] == "*":
            return a * b
        if b == 0:
            raise FormulaErro("divisão por zero")
        return a / b
    if tipo == "cmp":
        a, b = _ev(n[2], params), _ev(n[3], params)
        if isinstance(a, str) or isinstance(b, str):
            try:
                a, b = _num(a), _num(b)                                          # "8" = 8 (texto numérico compara como número)
            except FormulaErro:
                a, b = str(a).strip().upper(), str(b).strip().upper()
        op = n[1]
        return 1.0 if {"=": a == b, "<>": a != b, "<": a < b, ">": a > b, "<=": a <= b, ">=": a >= b}[op] else 0.0
    if tipo == "fn":
        nome, args = n[1], n[2]
        if nome == "SE":
            if len(args) != 3:
                raise FormulaErro("SE precisa de 3 argumentos: SE(cond; então; senão)")
            return _ev(args[1], params) if _verdade(_ev(args[0], params)) else _ev(args[2], params)
        if nome in ("E", "OU"):
            vals = [_verdade(_ev(a, params)) for a in args]
            return 1.0 if (all(vals) if nome == "E" else any(vals)) else 0.0
        minimo, maximo, f = FUNCOES[nome]
        if not minimo <= len(args) <= maximo:
            raise FormulaErro(f"{nome} recebeu {len(args)} argumento(s)")
        return f([_ev(a, params) for a in args])
    raise FormulaErro("fórmula inválida")


def _nome(n):
    return n[1].split("#")[0].lower() if n[0] == "id" else None


FUNCOES["SE"] = FUNCOES["E"] = FUNCOES["OU"] = (1, 99, None)         # só para o reconhecimento do nome (a avaliação é em _ev)


def avaliar(formula, params=None) -> float:
    """Valor da fórmula com os parâmetros do pedido ({NOME: VALOR}, nomes em maiúsculas; CABO#2 = 2º cabo). Erro → FormulaErro."""
    p = {str(k).upper(): v for k, v in (params or {}).items()}
    v = _ev(compilar(formula), p)
    if isinstance(v, str):
        v = _num(v)
    v = float(v)
    if not math.isfinite(v):
        raise FormulaErro("resultado não é um número")
    return v


def eh_formula(valor) -> bool:
    return str(valor if valor is not None else "").strip().startswith("=")


def normalizar_fator(valor, rotulo=""):
    """Célula FATOR I/R (CSV/tela) → (número, fórmula). Número: como sempre (vírgula decimal; ilegível = 0). Fórmula ("=…"): número 0 e a fórmula
    guardada com o "="; sintaxe inválida → ValueError (a importação recusa)."""
    if valor is None or isinstance(valor, bool):
        return 0.0, ""
    if isinstance(valor, (int, float)):
        return float(valor), ""
    t = str(valor).strip()
    if t.startswith("="):
        try:
            compilar(t)
        except FormulaErro as e:
            raise ValueError(f"{rotulo + ': ' if rotulo else ''}fórmula inválida '{t}': {e}.")
        return 0.0, "=" + _corpo(t).strip()
    try:
        return float(t.replace(",", ".")), ""
    except ValueError:
        return 0.0, ""
