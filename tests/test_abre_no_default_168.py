"""TASK-168 — o programa sempre abre no projeto default: o projeto aberto da execução anterior é esquecido (o usuário abre o projeto)."""
import json
import os
import subprocess

import pytest
from fastapi.testclient import TestClient

import config
from tests import oraculo_tela as o

PROJETO_JS = os.path.join(config.STATIC_DIR, "projeto.js")


def _cliente():
    import app
    return TestClient(app.app)


def test_pagina_traz_o_id_da_sessao_no_programa_instalado(monkeypatch):
    import middleware.nocache_middleware as m
    monkeypatch.setattr(m, "APP_MODE", "desktop")
    c = _cliente()
    r = c.get("/resumo")
    assert r.status_code == 200 and m.COOKIE_SESSAO + "=" + m.SESSAO_ID in r.headers.get("set-cookie", "")
    r2 = c.get("/resumo", cookies={m.COOKIE_SESSAO: m.SESSAO_ID})                  # já tem o id desta abertura: não manda de novo
    assert "set-cookie" not in r2.headers


def test_versao_online_nao_recebe_o_cookie(monkeypatch):
    import middleware.nocache_middleware as m
    monkeypatch.setattr(m, "APP_MODE", "server")
    r = _cliente().get("/resumo")
    assert "set-cookie" not in r.headers or m.COOKIE_SESSAO not in r.headers["set-cookie"]


def test_id_da_sessao_e_estavel_no_processo():
    import middleware.nocache_middleware as m
    assert len(m.SESSAO_ID) == 32


@pytest.mark.skipif(not o.NODE, reason="node não instalado")
def test_js_esquece_o_projeto_so_quando_a_abertura_muda():
    codigo = (f"const P=require({json.dumps(PROJETO_JS)});"
              "const mk=()=>{const d={};return {getItem:k=>(k in d?d[k]:null),setItem:(k,v)=>{d[k]=String(v)},removeItem:k=>{delete d[k]},d};};"
              "const s=mk(); s.setItem(P.CHAVE,'{\"caminho\":\"C:/x.obra\"}');"
              "const r=[];"
              "r.push(P.reiniciarNoDefault('a=1; gridr_sessao=AAA; b=2',s), s.getItem(P.CHAVE));"          # 1ª vez neste armazenamento: esquece
              "s.setItem(P.CHAVE,'{\"caminho\":\"C:/y.obra\"}');"
              "r.push(P.reiniciarNoDefault('gridr_sessao=AAA',s), s.getItem(P.CHAVE));"                      # mesma abertura (recarregou a tela/outra tela): mantém
              "r.push(P.reiniciarNoDefault('gridr_sessao=BBB',s), s.getItem(P.CHAVE));"                      # programa foi aberto de novo: esquece
              "r.push(P.reiniciarNoDefault('',s), P.reiniciarNoDefault('a=1',s));"                          # sem cookie (online): nada
              "console.log(JSON.stringify(r));")
    saida = subprocess.run([o.NODE, "-e", codigo], capture_output=True, text=True, timeout=30)
    assert saida.returncode == 0, saida.stderr
    assert json.loads(saida.stdout) == [True, None, False, '{"caminho":"C:/y.obra"}', True, None, False, False]
