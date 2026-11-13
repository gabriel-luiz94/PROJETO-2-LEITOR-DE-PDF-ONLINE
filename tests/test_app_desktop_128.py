"""TASK-128 — janela do desktop só carrega o login com o servidor pronto, campos de chave fora das heurísticas de senha do navegador, telas abertas dentro do programa (sem nova aba) e dica sem 'não encontrado' enquanto a base carrega."""
import os
import re
import socket
import threading
import time

import config
from services import inicio_janela as ij


def _ler(nome):
    with open(os.path.join(config.STATIC_DIR, nome), encoding="utf-8") as f:
        return f.read()


# ───────────────────────── abertura da janela ─────────────────────────
def test_espera_o_servidor_ficar_pronto():
    s = socket.socket(); s.bind(("127.0.0.1", 0)); porta = s.getsockname()[1]; s.close()        # porta livre (ninguém escutando ainda)
    assert ij.esperar_servidor("127.0.0.1", porta, limite=0.4, intervalo=0.05) is False

    def sobe():
        time.sleep(0.5)
        srv = socket.socket(); srv.setsockopt(socket.SOL_SOCKET, socket.SO_REUSEADDR, 1)
        srv.bind(("127.0.0.1", porta)); srv.listen(5); time.sleep(2.0); srv.close()
    threading.Thread(target=sobe, daemon=True).start()
    assert ij.esperar_servidor("127.0.0.1", porta, limite=5.0, intervalo=0.05) is True


def test_telas_de_inicio_e_erro():
    assert "Iniciando" in ij.TELA_INICIAL and "<html" in ij.TELA_INICIAL
    assert "Não foi possível iniciar" in ij.TELA_ERRO


def test_app_abre_a_janela_com_a_tela_inicial_e_carrega_o_login_depois():
    src = open(os.path.join(os.path.dirname(config.STATIC_DIR), "app.py"), encoding="utf-8").read()
    assert "html=TELA_INICIAL" in src and "win.load_url(login_url)" in src and "esperar_servidor(\"127.0.0.1\", port)" in src
    assert "url=login_url" not in src          # a janela não abre mais direto no login (o servidor ainda podia estar subindo)


# ───────────────────────── campos de chave / senha ─────────────────────────
def test_so_o_login_tem_campo_de_senha():
    for nome in os.listdir(config.STATIC_DIR):
        if nome.endswith(".html") and nome != "login.html":
            assert 'type="password"' not in _ler(nome), nome
    assert 'autocomplete="username"' in _ler("login.html") and 'autocomplete="current-password"' in _ler("login.html")


def test_campos_de_chave_usam_campo_segredo():
    resumo, index, admin = _ler("resumo.html"), _ler("index.html"), _ler("admin.html")
    for i in ("input-apikey", "input-apikey-openai", "input-apikey-claude"):
        assert re.search(rf'class="campo-segredo"[^>]*id="{i}"', resumo)
    assert 'class="campo-segredo"' in index and 'id="input-cloudconvert-key"' in index
    assert re.search(r'id="newPassword"[^>]*class="campo-segredo[ "]', admin)          # TASK-153: a tag tinha class= duas vezes (o 2º era ignorado); agora é um só, com as classes de estilo junto
    for pag in (resumo, index, admin):
        assert "/static/campos_segredo.js" in pag


# ───────────────────────── tudo dentro do programa ─────────────────────────
def test_nada_abre_em_nova_aba_no_desktop():
    assert "window.open('/resultado_orcamento'" not in _ler("resumo.js") and "AbrirNaApp.abrir('/resultado_orcamento'" in _ler("resumo.js")
    assert "window.open('/resumo'" not in _ler("script.js") and "AbrirNaApp.navegar('/resumo')" in _ler("script.js")
    assert "window.open('/orcamento'" not in _ler("resumo.html") and "AbrirNaApp.abrir('/orcamento'" in _ler("resumo.html")
    assert "AbrirNaApp.abrir('/manual'" in _ler("painel_regras.js")
    for pag in ("resumo.html", "index.html"):
        assert "/static/abrir_na_app.js" in _ler(pag)
    for pag in ("resultado_orcamento.html", "orcamento.html"):
        assert "function voltarOuFechar()" in _ler(pag) and 'onclick="voltarOuFechar()"' in _ler(pag)


# ───────────────────────── dica: base carregando ─────────────────────────
def test_dica_nao_desenha_com_a_base_vazia_e_a_base_e_pedida_ao_abrir():
    dica = _ler("ativo_dica.js")
    assert "if (!linhas.length) {" in dica and "if (cfg.carregarBase && !carregando)" in dica
    resumo = _ler("resumo.js")
    assert "carregarBase: () => carregarBaseOrcamento()" in resumo and re.search(r"^carregarBaseOrcamento\(\);", resumo, re.M)
