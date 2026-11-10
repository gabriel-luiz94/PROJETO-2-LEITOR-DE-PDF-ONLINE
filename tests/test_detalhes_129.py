"""TASK-129 — descrição e filtro dos códigos vêm da Base de Orçamento ao importar REC / modo comparação (/api/orcamento/detalhes); ")" completado no ativo (Enter / sair do campo)."""
import json
import os
import shutil
import subprocess

import pytest
from fastapi.testclient import TestClient

import config
import database
from services.orcamento_detalhes import detalhes_por_codigo

NODE = shutil.which("node")


def L(codigo, desc="", filtro="", projeto="", mdo="MATERIAL"):
    return {"codigo": codigo, "desc_codigo": desc, "filtro": filtro, "projeto": projeto, "mdo": mdo}


BASE = [
    L("111", "Poste genérico", "1.POSTE"),
    L("111", "Poste de Pernambuco", "9.PE", projeto="PERNAMBUCO"),
    L("222", "", "", projeto="PERNAMBUCO"),
    L("222", "Cabo X", "2.CABO"),
    L("333", "Só descrição", "", mdo="MAO-DE-OBRA"),
    L("abc", "Minúsculo", "A.F"),
]


def test_prefere_o_projeto_depois_o_generico():
    d = detalhes_por_codigo(BASE, ["111"], "Pernambuco")
    assert d["111"] == {"desc_codigo": "Poste de Pernambuco", "mdo": "MATERIAL", "filtro": "9.PE"}
    assert detalhes_por_codigo(BASE, ["111"], "OUTRO")["111"]["desc_codigo"] == "Poste genérico"          # projeto sem linha própria → genérica
    assert detalhes_por_codigo(BASE, ["111"])["111"]["filtro"] == "1.POSTE"


def test_descricao_e_filtro_de_linhas_diferentes():
    d = detalhes_por_codigo(BASE, ["222"], "PERNAMBUCO")["222"]
    assert d["desc_codigo"] == "Cabo X" and d["filtro"] == "2.CABO"           # a do projeto não tinha descrição nem filtro


def test_codigo_ausente_nao_vem_e_chave_e_a_pedida():
    d = detalhes_por_codigo(BASE, ["999", "ABC", " 333 "])
    assert "999" not in d and d["ABC"]["desc_codigo"] == "Minúsculo" and d["333"]["mdo"] == "MAO-DE-OBRA"
    assert detalhes_por_codigo(BASE, []) == {} and detalhes_por_codigo([], ["111"]) == {}


@pytest.fixture
def client(tmp_path, monkeypatch):
    monkeypatch.setattr(database, "DB_PATH", str(tmp_path / "t.db"))
    database.init_db()
    conn = database.get_connection()
    conn.execute("DELETE FROM tabela_orcamento_master")
    for r in BASE:
        conn.execute("INSERT INTO tabela_orcamento_master (codigo, desc_codigo, filtro, projeto, mdo) VALUES (?,?,?,?,?)", (r["codigo"], r["desc_codigo"], r["filtro"], r["projeto"], r["mdo"]))
    conn.commit(); conn.close()
    import app as appmod
    return TestClient(appmod.app)


def test_endpoint_devolve_descricao_e_filtro_da_base(client):
    r = client.post("/api/orcamento/detalhes", json={"codigos": ["111", "222", "999"], "projeto": "PERNAMBUCO"})
    assert r.status_code == 200
    d = r.json()
    assert d["111"] == {"desc_codigo": "Poste de Pernambuco", "mdo": "MATERIAL", "filtro": "9.PE"} and d["222"]["filtro"] == "2.CABO" and "999" not in d
    assert client.post("/api/orcamento/detalhes", json={"codigos": []}).json() == {}
    assert client.post("/api/orcamento/detalhes", json={"codigos": ["111"]}).json()["111"]["filtro"] == "1.POSTE"       # sem projeto: continua funcionando


# ───────────────────────── tela: filtro no REC, comparação e fantasmas ─────────────────────────
def _html():
    with open(os.path.join(config.STATIC_DIR, "resultado_orcamento.html"), encoding="utf-8") as f:
        return f.read()


def test_tela_usa_a_base_para_filtro_em_rec_comparacao_e_fantasma():
    h = _html()
    assert h.count("buscarDetalhesCodigos(") >= 4 and "fetch('/api/orcamento/detalhes'" in h and h.count("fetch('/api/orcamento/detalhes'") == 1      # um único ponto de busca
    assert "filtro: det && det.filtro" in h and "i.filtro = (det[i.codigo] && det[i.codigo].filtro)" in h
    assert "<td>${escHtml(f.filtro || '')}</td>" in h and 'data-campo="filtro">${escHtml(f.filtro' in h
    assert "filtro: item.filtro || ''" in h and "if (!map[key].filtro && item.filtro)" in h


# ───────────────────────── ")" que falta ─────────────────────────
@pytest.mark.skipif(not NODE, reason="node indisponível")
@pytest.mark.parametrize("entrada,saida", [
    ("N1(D, X", "N1(D, X)"), ("N1(D,", "N1(D)"), ("N1(", "N1()"), ("N1(D,A)", "N1(D,A)"), ("N1", "N1"),
    ("[CAZ7,94]", "[CAZ7,94]"), ("N1(D, IT_2(A", "N1(D, IT_2(A))"), ("2-N1(D, X", "2-N1(D, X)"), ("", ""),
])
def test_completar_parenteses(entrada, saida):
    cod = "const A=require('./static/ativo_dica.js');console.log(JSON.stringify(A.completarParenteses(JSON.parse(process.argv[1]))));"
    r = subprocess.run([NODE, "-e", cod, json.dumps(entrada)], capture_output=True, text=True, cwd=os.path.dirname(config.STATIC_DIR), timeout=30)
    assert r.returncode == 0, r.stderr
    assert json.loads(r.stdout) == saida


def test_dica_completa_no_enter_e_ao_sair_do_campo():
    with open(os.path.join(config.STATIC_DIR, "ativo_dica.js"), encoding="utf-8") as f:
        js = f.read()
    assert "e.key === 'Enter' && !e.isComposing && ehCampo(e.target)) completar(e.target)" in js
    assert "doc.addEventListener('blur', e => { if (ehCampo(e.target)) completar(e.target); }, true)" in js
