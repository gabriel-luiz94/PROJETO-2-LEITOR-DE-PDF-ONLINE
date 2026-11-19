"""TASK-189 — regressões do uso real: (1) abrir o Resumo com muitas linhas já gravadas (ler um DXF e abrir a aba) quebrava TODOS os botões
(«Cannot access 'grade' before initialization»: o 1º renderTable repinta a seleção ao montar as linhas e a `const grade` vinha depois no arquivo);
(2) fechar o programa gerava traceback do trabalhador do leitor (ConnectionResetError no Windows)."""
import os
import socket
import threading

import config
from services.autonomo import trabalhador_js


def _js():
    with open(os.path.join(config.STATIC_DIR, "resumo.js"), encoding="utf-8") as f:
        return f.read()


def test_estado_da_grade_e_declarado_antes_do_primeiro_render():
    js = _js()
    assert js.count("const grade = {") == 1
    assert js.index("const grade = {") < js.index("function renderTable(type)")
    # tudo que o render/hidratação lê durante a inicialização também fica no topo
    for nome in ["const _lazy =", "let _cacheCores", "const LIMIAR_LAZY"]:
        assert js.index(nome) < js.index("function renderTable(type)"), nome


def test_trabalhador_encerra_quieto_quando_o_principal_derruba_a_conexao():
    srv = socket.socket()
    srv.bind(("127.0.0.1", 0))
    srv.listen(1)
    porta = srv.getsockname()[1]
    erros = []

    def trabalhador():
        try:
            trabalhador_js.main([str(porta), "token"])
        except BaseException as e:  # noqa: BLE001
            erros.append(e)

    t = threading.Thread(target=trabalhador)
    t.start()
    conn, _ = srv.accept()
    conn.setsockopt(socket.SOL_SOCKET, socket.SO_LINGER, b"\x01\x00\x00\x00\x00\x00\x00\x00")      # fecha com RST (ConnectionReset no outro lado)
    conn.close()
    srv.close()
    t.join(timeout=20)
    assert not t.is_alive()
    assert erros == [], erros


def test_main_do_trabalhador_engole_so_erros_de_conexao():
    src = open(os.path.join(os.path.dirname(config.STATIC_DIR), "services", "autonomo", "trabalhador_js.py"), encoding="utf-8").read()
    assert "except (ConnectionError, BrokenPipeError, OSError):" in src
