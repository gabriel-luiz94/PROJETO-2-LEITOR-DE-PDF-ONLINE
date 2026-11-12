"""TASK-142 — ponte de arquivos da janela desktop; espelho da master na cópia pessoal do admin; recarga da base e da barra de modais."""
import base64
import json
import os
import shutil
import subprocess
import sys
import types

import pytest

import config
from services import arquivos_desktop as ad
from services import sync_service
from tests.test_fator_formula_109 import adm, banco, client  # noqa: F401

NODE = shutil.which("node")


class Janela:
    def __init__(self, saida):
        self.saida, self.chamadas = saida, []

    def create_file_dialog(self, tipo, **kw):
        self.chamadas.append((tipo, kw))
        return self.saida


@pytest.fixture
def webview_falso(monkeypatch):
    m = types.ModuleType("webview")
    m.SAVE_DIALOG, m.OPEN_DIALOG = 30, 10
    monkeypatch.setitem(sys.modules, "webview", m)
    return m


def test_filtros_de_arquivo():
    assert ad.filtros_de_arquivo(".pdf, .dxf, .dwg") == ("Arquivos permitidos (*.pdf;*.dxf;*.dwg)", "Todos os arquivos (*.*)")
    assert ad.filtros_de_arquivo(".json,application/json")[0] == "Arquivos permitidos (*.json)"
    assert ad.filtros_de_arquivo("") == ("Todos os arquivos (*.*)",) and ad.filtros_de_arquivo("image/*") == ("Todos os arquivos (*.*)",)


def test_salvar_arquivo_grava_os_bytes(tmp_path, webview_falso):
    alvo = tmp_path / "NOTA_1_ODI"
    api = ad.ApiArquivos(); api.vincular(Janela(str(alvo)))
    r = api.salvar_arquivo("NOTA_1_ODI.pdf", base64.b64encode(b"%PDF-1.4 teste").decode())
    assert r == {"ok": True, "caminho": str(alvo) + ".pdf"} and (tmp_path / "NOTA_1_ODI.pdf").read_bytes() == b"%PDF-1.4 teste"       # sem extensão escolhida: mantém a do arquivo
    assert api._janela.chamadas[0][1]["save_filename"] == "NOTA_1_ODI.pdf" and api._janela.chamadas[0][0] == 30
    api2 = ad.ApiArquivos(); api2.vincular(Janela(None))
    assert api2.salvar_arquivo("a.csv", "QQ==") == {"cancelado": True}
    assert ad.ApiArquivos().salvar_arquivo("a.csv", "QQ==")["ok"] is False                  # sem janela


def test_abrir_arquivos_devolve_nome_e_conteudo(tmp_path, webview_falso):
    f = tmp_path / "base.csv"; f.write_bytes("a;b\n1;ç".encode("utf-8"))
    api = ad.ApiArquivos(); api.vincular(Janela((str(f),)))
    r = api.abrir_arquivos(".csv", False)
    assert r["ok"] and r["arquivos"][0]["nome"] == "base.csv" and base64.b64decode(r["arquivos"][0]["base64"]) == "a;b\n1;ç".encode("utf-8")
    assert api._janela.chamadas[0][1]["allow_multiple"] is False and api._janela.chamadas[0][0] == 10
    api2 = ad.ApiArquivos(); api2.vincular(Janela([]))
    assert api2.abrir_arquivos() == {"cancelado": True}


def test_api_nao_expoe_a_janela():
    publicos = [n for n in dir(ad.ApiArquivos()) if not n.startswith("_")]
    assert sorted(publicos) == ["abrir_arquivos", "salvar_arquivo", "vincular"]


@pytest.mark.skipif(not NODE, reason="node não instalado")
def test_base64_js_ida_e_volta():
    mod = os.path.join(config.STATIC_DIR, "desktop_arquivos.js")
    codigo = f"const D=require({json.dumps(mod)});const b=new Uint8Array(70000).map((_,i)=>i%256);const r=D.deBase64(D.paraBase64(b.buffer));console.log(r.length===b.length&&r.every((x,i)=>x===b[i]));"
    assert subprocess.run([NODE, "-e", codigo], capture_output=True, text=True, timeout=30).stdout.strip() == "true"


def test_paginas_carregam_a_ponte_e_ha_rota():
    for pg in ("resumo", "resultado_orcamento", "orcamento", "index", "admin", "autonomo", "manual"):
        assert "/static/desktop_arquivos.js" in open(os.path.join(config.STATIC_DIR, pg + ".html"), encoding="utf-8").read(), pg
    assert "desktop_arquivos.js" in open(os.path.join(config.BASE_DIR, "app.py"), encoding="utf-8").read()


# ───────── cópia pessoal do admin acompanha a master (sincronizar com a nuvem = banco local na hora) ─────────
def linha(ativo, cod, proj="PERNAMBUCO", fi=1.0):
    return {"ativo": ativo, "desc_ativo": "d", "componente": ativo, "projeto": proj, "mdo": "MATERIAL", "origem": "", "codigo": cod, "desc_codigo": "d " + cod,
            "fator_i": fi, "fator_r": 0, "filtro": "", "lv": "A", "arred": "NORMAL", "formula_i": "", "formula_r": ""}


def total_local(client, ativo):
    """O que enxerga quem usa a cópia pessoal (plugin do nanoCAD, modo autônomo): get_merged_orcamento('u1')."""
    return sorted((x["codigo"], x["fator_i"]) for x in sync_service.get_merged_orcamento("u1") if x["ativo"] == ativo)


def test_sincronizar_atualiza_a_copia_pessoal(client):
    # o admin salvou uma cópia pessoal com a definição ANTIGA do ativo
    assert client.post("/api/orcamento/salvar", json={"dados": [linha("SUH", "111"), linha("OUTRO", "9", "PARAIBA")], "projeto": ""}, headers=adm()).status_code == 200
    assert total_local(client, "SUH") == [("111", 1.0)]
    # edita o ativo e sincroniza com a nuvem (master) — sem clicar em "Salvar"
    r = client.post("/api/admin/sync-master-all", json={"dados": [linha("SUH", "222", fi=3.0)], "projeto": "PERNAMBUCO"}, headers=adm())
    assert r.status_code == 200, r.text
    assert total_local(client, "SUH") == [("222", 3.0)]                                            # a definição nova já vale (antes ficava a antiga)
    assert sorted((x["codigo"], x["fator_i"]) for x in client.get("/api/orcamento/dados", headers=adm()).json()["dados"] if x["ativo"] == "SUH") == [("222", 3.0)]   # e a tela (master)
    assert total_local(client, "OUTRO") == [("9", 1.0)]                                            # outros projetos da cópia pessoal ficam como estavam


def test_sem_copia_pessoal_nao_cria_uma(client):
    assert client.post("/api/admin/sync-master-all", json={"dados": [linha("SUH", "222")], "projeto": "PERNAMBUCO"}, headers=adm()).status_code == 200
    conn = sync_service.get_row_connection()
    n = conn.execute("SELECT COUNT(*) FROM tabela_orcamento WHERE user_id = 'u1'").fetchone()[0]
    conn.close()
    assert n == 0 and total_local(client, "SUH") == [("222", 1.0)]


def test_upload_master_e_add_row_tambem_acompanham(client):
    assert client.post("/api/orcamento/salvar", json={"dados": [linha("SUH", "111")], "projeto": ""}, headers=adm()).status_code == 200
    r = client.post("/api/admin/add-row", json={"ativo": "SUH", "desc_ativo": "d", "componente": "SUH", "projeto": "PERNAMBUCO", "mdo": "MATERIAL", "codigo": "333",
                                                "desc_codigo": "d", "fator_i": 2, "fator_r": 0, "filtro": "", "lv": "A", "arred": "NORMAL"}, headers=adm())
    assert r.status_code == 200, r.text
    assert ("333", 2.0) in total_local(client, "SUH")


def test_resumo_recarrega_barra_ao_trocar_de_projeto_e_base_ao_fechar_a_tabela():
    html = open(os.path.join(config.STATIC_DIR, "resumo.html"), encoding="utf-8").read()
    assert "await BarraModais.recarregar()" in html
    js = open(os.path.join(config.STATIC_DIR, "resumo.js"), encoding="utf-8").read()
    assert "'tela-app-fechada'" in js and "visibilitychange" in js
    assert "tela-app-fechada" in open(os.path.join(config.STATIC_DIR, "abrir_na_app.js"), encoding="utf-8").read()
