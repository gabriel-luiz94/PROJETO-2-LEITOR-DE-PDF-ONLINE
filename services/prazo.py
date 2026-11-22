"""Interrupção cooperativa de cálculos executados em uma thread."""
from contextvars import ContextVar
from collections.abc import Callable


class PrazoExcedido(RuntimeError):
    pass


_cancelado: ContextVar[Callable[[], bool] | None] = ContextVar("cancelado", default=None)


def verificar_prazo():
    cancelado = _cancelado.get()
    if cancelado is not None and cancelado():
        raise PrazoExcedido


def executar_com_prazo(cancelado: Callable[[], bool], funcao, *args):
    token = _cancelado.set(cancelado)
    try:
        verificar_prazo()
        resultado = funcao(*args)
        verificar_prazo()
        return resultado
    finally:
        _cancelado.reset(token)
