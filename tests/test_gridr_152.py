"""TASK-152 — ícone (grade com o caminho laranja ligando os nós). O NOME passou de GridR para OrçaCore na TASK-156 (tests/test_orcacore_156.py); o ícone não mudou."""
import glob
import os
import struct

import config

ESTATICOS = config.STATIC_DIR


def ler(*partes):
    return open(os.path.join(*partes), encoding="utf-8").read()


def test_icone_ico_tem_todos_os_tamanhos_com_png_dentro():
    d = open(os.path.join(ESTATICOS, "gridr.ico"), "rb").read()
    assert d[:4] == b"\x00\x00\x01\x00"
    n = struct.unpack("<H", d[4:6])[0]
    tamanhos = [(d[6 + 16 * i] or 256) for i in range(n)]
    assert tamanhos == [16, 24, 32, 48, 64, 128, 256]
    for i in range(n):                                                                 # cada entrada aponta para um PNG de verdade
        tam, off = struct.unpack("<II", d[6 + 16 * i + 8:6 + 16 * i + 16])
        assert d[off:off + 8] == b"\x89PNG\r\n\x1a\n" and off + tam <= len(d)
    assert open(os.path.join(ESTATICOS, "gridr-512.png"), "rb").read()[:8] == b"\x89PNG\r\n\x1a\n"
    svg = ler(ESTATICOS, "gridr.svg")
    assert "#d97706" in svg and "<svg" in svg


def test_paginas_com_icone_e_nome_novo():
    paginas = glob.glob(os.path.join(ESTATICOS, "*.html"))
    assert len(paginas) >= 8
    for p in paginas:
        html = ler(p)
        assert "/static/gridr.svg" in html and "/static/gridr.ico" in html, p
        assert "GridReader" not in html, p
    marca = ler(ESTATICOS, "marca.js")
    assert "/static/gridr.svg" in marca and "window.GridR = window.GridReader" in marca


def test_nome_na_janela_e_no_build():
    app = ler(config.BASE_DIR, "app.py")
    bat = ler(config.BASE_DIR, "build.bat")
    assert '--icon "static\\gridr.ico"' in bat


def test_plugin_continua_reconhecendo_o_programa(client_plugin=None):
    """O plugin do nanoCAD já instalado confere "app" == "GridReader": o identificador não muda, só ganha "nome"."""
    src = ler(config.BASE_DIR, "routers", "plugin_cad.py")
    assert '"app": "GridReader"' in src
