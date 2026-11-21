"""TASK-215 — a ponte só abre visualizações conhecidas e fecha as janelas dependentes."""
from types import SimpleNamespace
from unittest.mock import Mock

from services.janelas_auxiliares import ApiProjetoComJanelas


class Evento:
    def __init__(self):
        self.ouvintes = []

    def __iadd__(self, callback):
        self.ouvintes.append(callback)
        return self

    def disparar(self):
        return [callback() for callback in self.ouvintes]


def janela_falsa():
    eventos = SimpleNamespace(**{nome: Evento() for nome in ("closed", "closing", "maximized", "restored")})
    janela = Mock(events=eventos)
    janela.destroy.side_effect = eventos.closed.disparar
    return janela


def test_abre_so_as_visualizacoes_da_mesma_origem_e_reutiliza_o_processo():
    principal = janela_falsa()
    resultado, base = janela_falsa(), janela_falsa()
    criar = Mock(side_effect=[resultado, base])
    api = ApiProjetoComJanelas("http://127.0.0.1:8123", criar)
    api.vincular(principal)
    canal = "a" * 32

    assert not api.janela_auxiliar_abrir("https://outro", canal)["ok"]
    assert not api.janela_auxiliar_abrir("resultado", "x")["ok"]
    assert api.janela_auxiliar_abrir("resultado", canal)["ok"]
    kw = criar.call_args.kwargs
    assert kw["url"] == "http://127.0.0.1:8123/static/janela_auxiliar.html#resultado=" + canal
    assert kw["frameless"] and not kw["easy_drag"]
    assert not api.janela_auxiliar_abrir("resultado", "b" * 32)["ok"]

    controle = kw["js_api"]
    controle.janela_minimizar()
    resultado.minimize.assert_called_once()
    assert controle.janela_maximizar()
    resultado.maximize.assert_called_once()
    resultado.events.maximized.disparar()
    assert not controle.janela_maximizar()
    resultado.restore.assert_called_once()

    api.janela_auxiliar_fechar(canal)
    resultado.destroy.assert_called_once()
    principal.evaluate_js.assert_called_once_with(f"window.AbrirNaApp?.nativaFechada('{canal}')")
    assert api.janela_auxiliar_abrir("base", "b" * 32)["ok"]
    principal.events.closed.disparar()
    base.destroy.assert_called_once()
    principal.evaluate_js.assert_called_once()
