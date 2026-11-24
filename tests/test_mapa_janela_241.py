"""TASK-241: o Mapa (antigo «Vincular cabos») abre como as janelas de Resultado e Base de Orçamento — área principal, rodapé,
maximizar/restaurar e «Abrir em nova janela» do Windows (o Resumo em modo janela, só com o Mapa, sincronizado com a janela principal).
Esc no Mapa só deseleciona. O fluxo completo (popup como janela nativa, sincronização nos dois sentidos, retornar, fechar) foi
conferido no Chromium; tests/test_resumo_janelas_213.py cobre o mesmo onde houver Playwright para Python."""
import os
from types import SimpleNamespace
from unittest.mock import Mock

import config
from services.janelas_auxiliares import ApiProjetoComJanelas

S = config.STATIC_DIR
ler = lambda nome: open(os.path.join(S, nome), encoding="utf-8").read()


def test_ponte_nativa_abre_o_mapa():
    eventos = lambda: SimpleNamespace(closed=_Evento(), closing=_Evento())
    principal, mapa = Mock(events=eventos()), Mock(events=eventos())
    criar = Mock(return_value=mapa)
    api = ApiProjetoComJanelas("http://127.0.0.1:8123", criar)
    api.vincular(principal)
    assert api.janela_auxiliar_abrir("mapa", "d" * 32)["ok"]
    assert criar.call_args.kwargs["url"] == "http://127.0.0.1:8123/static/janela_auxiliar.html#mapa=" + "d" * 32
    assert criar.call_args.kwargs["title"] == "OrçaCore · Mapa"


class _Evento:
    def __init__(self):
        self.ouvintes = []

    def __iadd__(self, f):
        self.ouvintes.append(f)
        return self


def test_moldura_nativa_conhece_o_mapa_e_repassa_os_dados():
    js = ler("janela_auxiliar.js")
    assert "mapa: ['/resumo', 'Mapa']" in js and "^(resultado|base|mapa)=" in js
    assert "raiz.__enviarDadosMapa = dados => { if (ligada) enviar('dados', { dados }); };" in js
    assert "if (data.acao === 'dados') conteudo()?.dados?.(data.dados);" in js


def test_visualizacao_de_elemento_vivo():
    js = ler("abrir_na_app.js")
    for nome in ("abrirConteudo", "removerConteudo", "trazerConteudo", "nativaAtiva", "enviarNativa"):
        assert nome in js.split("raiz.AbrirNaApp = {")[1]
    assert "const capturarDe = p => (p.capturar ? p.capturar() : conteudoDa(p)?.capturar());" in js
    assert "if (msg.acao === 'dados' && p.conteudo) p.aoDados?.(msg.dados);" in js


def test_resumo_em_modo_janela_do_mapa():
    js = ler("resumo.js")
    assert "const MODO_JANELA_MAPA = new URLSearchParams(location.search).has('janela_transferida');" in js
    assert "if (!MODO_JANELA_MAPA) { try { localStorage.setItem('processar_dados'" in js       # a cópia não grava a obra
    assert "consultarAssistente: MODO_JANELA_MAPA ? undefined" in js
    assert "instantaneo: instantaneoDoMapa," in js and "aplicarInstantaneo: aplicarInstantaneoDoMapa" in js
    assert "assinaturaMapa = JSON.stringify(instantaneoDoMapa());" in js                          # sem eco para a principal


def test_mapa_renomeado_abre_na_aba_mapa_e_esc_so_deseleciona():
    ui, html, barra = ler("vinculacao_ui.js"), ler("resumo.html"), ler("barra_modais.js")
    assert "aba: (opcoes && opcoes.aba) || 'mapa'," in ui
    assert "const abas = [['mapa', 'Mapa', ''], ['vinculo', 'Vínculo', '']" in ui
    assert "el('h3', { text: '🗺 Mapa' })" in ui and "Vinculação cabo ⇄ estrutura" not in ui
    assert ">MAPA</button>" in html and "{ id: 'vincular', nome: 'MAPA', icone: '🗺' }" in barra
    assert "else if (st.aba !== 'mapa' && !overlay.classList.contains('vz-destacada')) fechar();" in ui
    assert "Esc no Mapa só deseleciona" in ui and "caboEmFoco = null; atualizarSelecao();" in ui
    assert "!svg.getClientRects().length" in ui                                                   # minimizado não pega o Esc do Resumo
