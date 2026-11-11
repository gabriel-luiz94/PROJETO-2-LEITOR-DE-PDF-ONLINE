"""TASK-137 — a base de orçamento aceita CSV salvo em ANSI (Windows-1252), como o Excel grava, além de UTF-8."""
from tests.test_fator_formula_109 import adm, banco, client  # noqa: F401


def _csv(enc):
    txt = "ATIVO;DESC ATIVO;COMPONENTE;PROJETO;MDO;ORIGEM;CODIGO;DESC CODIGO;FATOR I;FATOR R;FILTRO;LV;ARRED\n" \
          "CON_X;CONEXÃO CUNHA;CON_X;P;MATERIAL;;1;CONETOR AÇO;1;0;99.CONECTOR;A;NORMAL\n"
    return txt.encode(enc)


def test_upload_orcamento_aceita_ansi_e_utf8(client):
    for enc in ("cp1252", "utf-8", "utf-8-sig"):
        r = client.post("/api/orcamento/upload", files={"file": ("b.csv", _csv(enc), "text/csv")}, headers=adm())
        assert r.status_code == 200, (enc, r.text)
    r = client.post("/api/admin/upload-master", files={"file": ("b.csv", _csv("cp1252"), "text/csv")}, headers=adm())
    assert r.status_code == 200, r.text
