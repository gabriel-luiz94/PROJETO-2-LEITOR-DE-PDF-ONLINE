"""TASK-145 — publicar / desfazer UM ativo (local + nuvem), prévia com diff, passo anterior e detecção de edição simultânea."""
import pytest

import database
from services import publicar_ativo as pub
from services import sync_service
from tests.test_fator_formula_109 import adm, banco, client  # noqa: F401

ADMIN = {"user_id": "u1", "email": "a@x.com"}


def L(a, comp="", cod="", d="", fi=0, proj="PE", mdo=None, **k):
    r = {"ativo": a, "desc_ativo": "d", "componente": comp, "projeto": proj, "mdo": mdo if mdo is not None else ("MATERIAL" if cod else ""), "codigo": cod, "desc_codigo": d, "fator_i": fi,
         "fator_r": 0, "filtro": "", "origem": "", "lv": "A", "arred": "NORMAL", "formula_i": "", "formula_r": ""}
    r.update(k)
    return r


BASE = [L("TERM=TERM(cabo=S04)"), L("TERM(cabo=%S04%)", "T", "2418089", "TERM", 1), L("OUTRO", "O", "9", "x", 1), L("OUTRO", "O", "8", "y", 2, proj="PARAIBA")]
NOVO = [L("X=X(cabo=S04), XX"), L("X", "X", "500109", "CH", 3), L("X(cabo=%):=3*TERM(cabo)")]


@pytest.fixture
def db(banco):
    database.init_db()
    conn = database.get_connection()
    conn.execute("DELETE FROM tabela_orcamento_master")
    for r in BASE:
        pub._inserir_local(conn.cursor(), [r])
    conn.commit(); conn.close()
    return banco


def local(nome, proj="PE"):
    return [(r["ativo"], r["codigo"], r["fator_i"]) for r in pub._locais() if r["id"] in set(pub.ids_do_ativo(pub._locais(), nome, proj))]


class Nuvem:
    """Supabase de mentira: a tabela como lista de dicts com id, e os poucos métodos usados."""
    def __init__(self, linhas, falha_insert=False, falha_delete=False):
        self.t, self.prox, self.falha_insert, self.falha_delete = [dict(r, id=i + 1) for i, r in enumerate(linhas)], len(linhas) + 1, falha_insert, falha_delete
        self.op, self.args, self.chamadas = None, None, []

    def table(self, nome):
        self.op, self.filtro = None, None
        return self

    def select(self, *_): self.op = "select"; return self
    def ilike(self, col, v): self.filtro = ("ilike", col, v); return self
    def order(self, *_): return self
    def range(self, a, b): self.faixa = (a, b); return self
    def insert(self, linhas): self.op, self.args = "insert", linhas; return self
    def delete(self): self.op = "delete"; return self
    def in_(self, col, ids): self.args = ids; return self

    def execute(self):
        class R: pass
        r = R()
        if self.op == "select":
            r.data = [dict(x) for x in self.t if str(x.get("projeto", "")).upper() == self.filtro[2].upper()][self.faixa[0]:self.faixa[1] + 1]
        elif self.op == "insert":
            if self.falha_insert:
                raise RuntimeError("insert falhou")
            r.data = []
            for x in self.args:
                x = dict(x, id=self.prox); self.prox += 1; self.t.append(x); r.data.append(dict(x))
        else:
            if self.falha_delete:
                self.falha_delete = False                                    # só a 1ª exclusão falha (a de rollback passa)
                raise RuntimeError("delete falhou")
            self.t = [x for x in self.t if x["id"] not in set(self.args)]
            r.data = []
        return r

    def ativo(self, nome, proj="PE"): return [(x["ativo"], x["codigo"]) for x in self.t if x["id"] in set(pub.ids_do_ativo(self.t, nome, proj))]


def test_ids_do_ativo_com_linha_em_branco_e_projeto():
    base = [dict(L("N1=N1(c=A)"), id=1), dict(L("", "N1", "1", "x", 1), id=2), dict(L("N2", "N2", "2", "x", 1), id=3), dict(L("N1", "N1", "3", "x", 1, proj="PARAIBA"), id=4)]
    assert pub.ids_do_ativo(base, "n1", "pe") == [1, 2] and pub.ids_do_ativo(base, "N1", "PARAIBA") == [4]


def test_publicar_ativo_novo_so_troca_as_linhas_dele(db):
    r = pub.publicar(ADMIN, "PE", "X", NOVO, esperada=pub.assinatura([]), sb=None)
    assert r["status"] == "ok" and r["linhas"] == 3 and r["removidas"] == 0 and r["nuvem"] is False
    assert [a for a, _, _ in local("X")] == ["X=X(cabo=S04), XX", "X", "X(cabo=%):=3*TERM(cabo)"]
    assert len(pub._locais()) == len(BASE) + 3 and local("OUTRO") == [("OUTRO", "9", 1.0)]               # o resto da base não foi tocado
    # edita: troca o fator e remove a chamada
    novo2 = [L("X, XX"), L("X", "X", "500109", "CH", 5)]
    p = pub.previa("PE", "X", novo2)
    d = p["diff"]
    assert p["erros"] == [] and [a["campos"] for a in d["alteradas"]] == [["fator_i"]] and len(d["removidas"]) == 2 and len(d["novas"]) == 1 and d["iguais"] == 0
    pub.publicar(ADMIN, "PE", "X", novo2, esperada=p["assinatura"], sb=None)
    assert local("X") == [("X, XX", "", 0.0), ("X", "500109", 5.0)]


def test_desfazer_um_passo_e_refazer(db):
    pub.publicar(ADMIN, "PE", "X", NOVO, esperada=pub.assinatura([]), sb=None)
    assert pub.versao("PE", "X")["tem_anterior"] is True
    pub.publicar(ADMIN, "PE", "X", [L("X", "X", "1", "a", 1)], esperada=pub.versao("PE", "X")["assinatura"], sb=None)
    assert [c for _, c, _ in local("X")] == ["1"]
    pub.desfazer(ADMIN, "PE", "X", sb=None)                                                              # volta à 1ª publicação
    assert [c for _, c, _ in local("X")] == ["", "500109", ""]
    pub.desfazer(ADMIN, "PE", "X", sb=None)                                                              # de novo: refaz
    assert [c for _, c, _ in local("X")] == ["1"]
    with pytest.raises(pub.ErroPublicacao) as e:
        pub.desfazer(ADMIN, "PE", "NUNCA", sb=None)
    assert e.value.status == 404


def test_desfazer_a_criacao_remove_o_ativo(db):
    pub.publicar(ADMIN, "PE", "X", NOVO, esperada=pub.assinatura([]), sb=None)
    pub.desfazer(ADMIN, "PE", "X", sb=None)
    assert local("X") == [] and len(pub._locais()) == len(BASE)


def test_edicao_simultanea_409_e_forcar(db):
    sig = pub.assinatura([])
    pub.publicar(ADMIN, "PE", "X", NOVO, esperada=sig, sb=None)                                         # outro admin criou antes
    with pytest.raises(pub.ErroPublicacao) as e:
        pub.publicar(ADMIN, "PE", "X", [L("X", "X", "1", "a", 1)], esperada=sig, sb=None)
    assert e.value.status == 409 and e.value.extra["conflito"] and len(e.value.extra["linhas_atuais"]) == 3
    pub.publicar(ADMIN, "PE", "X", [L("X", "X", "1", "a", 1)], esperada=sig, forcar=True, sb=None)
    assert local("X") == [("X", "1", 1.0)]


def test_validacao_recusa_e_nao_grava(db):
    with pytest.raises(pub.ErroPublicacao) as e:
        pub.publicar(ADMIN, "PE", "Y", [L("Y=Y(cabo=S04)"), L("Y(foo=%A%)", "", "1", "d", 1)], sb=None)
    assert e.value.status == 400 and local("Y") == []
    with pytest.raises(pub.ErroPublicacao):
        pub.publicar(ADMIN, "PE", "Z", [L("OUTRO_NOME", "c", "1", "d", 1)], sb=None)                    # linha de outro ativo
    with pytest.raises(pub.ErroPublicacao):
        pub.publicar(ADMIN, "", "Z", [L("Z", "c", "1", "d", 1)], sb=None)


def test_previa_mostra_quem_chama_o_ativo(db):
    pub.publicar(ADMIN, "PE", "X", NOVO, esperada=pub.assinatura([]), sb=None)
    p = pub.previa("PE", "TERM", [L("TERM=TERM(cabo=S04)"), L("TERM(cabo=%S04%)", "T", "2418089", "TERM", 2)])
    assert [c["ativo"] for c in p["chamadores"]] == ["X"] and p["diff"]["alteradas"][0]["campos"] == ["fator_i"]


def test_nuvem_primeiro_e_sem_deixar_nada_pela_metade(db):
    nuvem = Nuvem(BASE)
    pub.publicar(ADMIN, "PE", "X", NOVO, esperada=pub.assinatura([]), sb=nuvem)
    assert len(nuvem.ativo("X")) == 3 and nuvem.ativo("OUTRO") == [("OUTRO", "9")]
    r = pub.publicar(ADMIN, "PE", "X", [L("X", "X", "1", "a", 1)], esperada=pub.versao("PE", "X", nuvem)["assinatura"], sb=nuvem)
    assert r["nuvem"] is True and nuvem.ativo("X") == [("X", "1")] and local("X") == [("X", "1", 1.0)]
    # falha ao inserir na nuvem: nada muda, nem no local
    ruim = Nuvem(BASE, falha_insert=True)
    with pytest.raises(pub.ErroPublicacao) as e:
        pub.publicar(ADMIN, "PE", "W", [L("W", "W", "1", "a", 1)], sb=ruim)
    assert e.value.status == 502 and local("W") == [] and ruim.ativo("W") == []
    # falha ao apagar as antigas: as inseridas são removidas de volta
    meio = Nuvem(BASE + NOVO[:2], falha_delete=True)
    antes = len(meio.t)
    with pytest.raises(pub.ErroPublicacao):
        pub.publicar(ADMIN, "PE", "X", [L("X", "X", "1", "a", 1)], sb=meio)
    assert len(meio.t) == antes and meio.ativo("X") == [("X=X(cabo=S04), XX", ""), ("X", "500109")]


def test_copia_pessoal_do_admin_acompanha(db):
    conn = database.get_connection()
    conn.execute("INSERT INTO tabela_orcamento (user_id, ativo, projeto, codigo) VALUES ('u1', 'VELHO', 'PE', '1')"); conn.commit(); conn.close()
    pub.publicar(ADMIN, "PE", "X", NOVO, esperada=pub.assinatura([]), sb=None)
    assert any(r["ativo"] == "X" for r in sync_service.get_merged_orcamento("u1"))


def test_rotas_admin(client):
    h = adm()
    conn = database.get_connection(); conn.execute("DELETE FROM tabela_orcamento_master"); conn.commit(); conn.close()
    v = client.get("/api/admin/ativo/versao?projeto=PE&nome=X", headers=h).json()
    assert v["linhas"] == [] and v["tem_anterior"] is False
    p = client.post("/api/admin/ativo/previa", json={"projeto": "PE", "nome": "X", "linhas": [L("X", "X", "1", "a", 2)]}, headers=h).json()
    assert p["erros"] == [] and len(p["diff"]["novas"]) == 1
    r = client.post("/api/admin/ativo/publicar", json={"projeto": "PE", "nome": "X", "linhas": [L("X", "X", "1", "a", 2)], "esperada": v["assinatura"]}, headers=h)
    assert r.status_code == 200 and r.json()["linhas"] == 1
    again = client.post("/api/admin/ativo/publicar", json={"projeto": "PE", "nome": "X", "linhas": [L("X", "X", "2", "b", 1)], "esperada": v["assinatura"]}, headers=h)
    assert again.status_code == 409 and again.json()["conflito"] is True and "a@x.com" in again.json()["detail"]
    assert client.post("/api/admin/ativo/desfazer", json={"projeto": "PE", "nome": "X"}, headers=h).status_code == 200
    assert client.get("/api/admin/ativo/versao?projeto=PE&nome=X", headers=h).json()["linhas"] == []
    assert client.post("/api/admin/ativo/publicar", json={}, headers={}).status_code in (401, 403)
