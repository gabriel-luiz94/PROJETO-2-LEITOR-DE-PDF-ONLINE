"""tests/test_voz_chat.py — voz no chat de IA do Resumo (TASK-027). Frontend puro: aqui só as garantias estáticas;
o comportamento (ditado, erros, leitura) foi verificado no Chromium com reconhecimento/síntese de voz simulados."""
import pytest
from fastapi.testclient import TestClient

import app as appmod
import database


@pytest.fixture
def client(tmp_path, monkeypatch):
    monkeypatch.setattr(database, "DB_PATH", str(tmp_path / "teste.db"))
    database.init_db()
    return TestClient(appmod.app)


def test_voz_js_e_servido_sem_cache_e_referenciado_no_resumo(client):
    r = client.get("/static/voz.js")
    assert r.status_code == 200 and "javascript" in r.headers["content-type"] and "no-cache" in r.headers.get("cache-control", "")
    html = open("static/resumo.html", encoding="utf-8").read()
    for ident in ("btn-mic-chat", "btn-voz-resposta", "btn-voz-parar", "voz-estado"):
        assert f'id="{ident}"' in html, ident
    assert "/static/voz.js" in html


def test_ditado_nunca_envia_sozinho_nem_usa_innerhtml_e_a_leitura_vem_desligada():
    js = open("static/voz.js", encoding="utf-8").read()
    assert "pt-BR" in js and "interimResults = true" in js
    assert "btn-send-chat" not in js and "btnEnviar" not in js and ".click()" not in js      # não aciona o envio
    assert "innerHTML" not in js                                                              # texto falado/ditado só por textContent/value
    assert "localStorage.getItem(PREF_LER) === 'sim'" in js                                   # padrão = desligada
