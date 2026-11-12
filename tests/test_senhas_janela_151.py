"""TASK-151 — a janela do programa liga 'salvar senhas'/'preenchimento automático' mesmo quando o WebView2 ainda está inicializando."""
import sys
import types

from services import janela_desktop


class Settings:
    IsPasswordAutosaveEnabled = False
    IsGeneralAutofillEnabled = False


class Core:
    def __init__(self):
        self.Settings = Settings()


class Web:
    def __init__(self, pronto_apos):
        self.chamadas, self.pronto_apos, self._core = 0, pronto_apos, Core()

    @property
    def CoreWebView2(self):
        self.chamadas += 1
        return self._core if self.chamadas > self.pronto_apos else None            # None enquanto o WebView2 inicializa


class Form:
    def __init__(self, web):
        self.browser = types.SimpleNamespace(webview=web)

    def Invoke(self, acao):
        acao()


def janela_com(pronto_apos):
    web = Web(pronto_apos)
    return types.SimpleNamespace(native=Form(web)), web


def sistema_falso(monkeypatch):
    m = types.ModuleType("System")
    m.Action = lambda f: f
    monkeypatch.setitem(sys.modules, "System", m)


def test_espera_o_webview2_ficar_pronto_e_liga(monkeypatch):
    sistema_falso(monkeypatch)
    win, web = janela_com(pronto_apos=3)
    assert janela_desktop.ligar_senhas_salvas(win, tentativas=10, intervalo=0) is True
    assert web.CoreWebView2.Settings.IsPasswordAutosaveEnabled and web.CoreWebView2.Settings.IsGeneralAutofillEnabled and web.chamadas >= 4      # tentou de novo até o controle existir


def test_desiste_depois_das_tentativas_sem_travar(monkeypatch):
    sistema_falso(monkeypatch)
    win, web = janela_com(pronto_apos=999)
    assert janela_desktop.ligar_senhas_salvas(win, tentativas=4, intervalo=0) is False


def test_fora_do_windows_nao_tenta_de_novo(monkeypatch):
    monkeypatch.delitem(sys.modules, "System", raising=False)
    win, web = janela_com(pronto_apos=0)
    assert janela_desktop.ligar_senhas_salvas(win, tentativas=50, intervalo=0) is False and web.chamadas == 0            # sem pythonnet: ImportError, sai na hora
    assert janela_desktop.ligar_senhas_salvas(types.SimpleNamespace(native=types.SimpleNamespace(browser=None)), tentativas=2, intervalo=0) is False
