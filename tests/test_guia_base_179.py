"""TASK-179 — guia e modelo para quem cria uma base nova: pacote de exemplo que importa e funciona, frase em português por linha, códigos do validador e aviso de equipamento no grupo geral."""
import json
import os
import re
import shutil
import subprocess

import pytest

import config
import database
from services import casos_referencia as cr
from services import pacote_projeto as pp
from services.ativo_parametros import codigo_do_achado, com_codigo, preencher_componentes, resumo_erros, validar_base

NODE = shutil.which("node")
GUIA = os.path.join(config.BASE_DIR, "docs", "base_guia.md")


def L(ativo, comp, cod, desc, mdo="MATERIAL", fi=1, proj="T"):
    return {"ativo": ativo, "desc_ativo": "", "componente": comp, "projeto": proj, "mdo": mdo, "codigo": cod, "desc_codigo": desc, "fator_i": fi, "fator_r": 0, "filtro": "", "origem": "", "lv": "A", "arred": "NORMAL",
            "formula_i": "", "formula_r": ""}


# CE_2_TR antes da correção: um transformador de catálogo nas linhas GERAIS de um ativo que tem o parâmetro `trafo` (saía com dois trafos); depois: cada trafo no grupo do seu valor
ANTES = [L("CE_2_TR=CE_2_TR(trafo=30)", "CE_2_TR", "9001", "CRUZETA CE"), L("", "CE_2_TR", "9002", "TRANSFORMADOR 45 KVA"), L("", "CE_2_TR", "9003", "PARAFUSO", "MATERIAL"),
         L("CE_2_TR(trafo=30)", "CE_2_TR.30", "9010", "TRANSFORMADOR 30 KVA"), L("CE_2_TR(trafo=45)", "CE_2_TR.45", "9011", "TRANSFORMADOR 45 KVA")]
DEPOIS = [L("CE_2_TR=CE_2_TR(trafo=30)", "CE_2_TR", "9001", "CRUZETA CE"), L("", "CE_2_TR", "9003", "PARAFUSO"),
          L("CE_2_TR(trafo=30)", "CE_2_TR.30", "9010", "TRANSFORMADOR 30 KVA"), L("CE_2_TR(trafo=45)", "CE_2_TR.45", "9011", "TRANSFORMADOR 45 KVA")]


def _avisos(rows):
    rows, _ = preencher_componentes([dict(r) for r in rows])
    r = validar_base(rows)
    assert not r["erros"], r["erros"]
    return [a for a in r["avisos"] if "equipamento" in a]


def test_equipamento_no_grupo_geral_gera_aviso_e_a_base_corrigida_nao():
    av = _avisos(ANTES)
    assert len(av) == 1 and "9002" in av[0] and "CE_2_TR" in av[0] and "trafo" in av[0] and "Não deveria estar no grupo do parâmetro" in av[0]
    assert _avisos(DEPOIS) == []
    assert codigo_do_achado(av[0]) == "B-013"


def test_equipamento_so_vale_para_material_e_para_parametro_do_mesmo_tipo():
    mao = [dict(r, mdo="MAO-DE-OBRA") if r["codigo"] == "9002" else r for r in ANTES]
    assert _avisos(mao) == []
    outro_param = [dict(r, ativo=r["ativo"].replace("trafo", "cabo")) for r in ANTES]
    assert _avisos(outro_param) == []


def test_codigos_do_validador_e_guia_citam_todos():
    amostras = {"B-002": "Linha 2: ativo 'N3(XLPE)': na base os parâmetros levam nome (cabo=XLPE).", "B-003": "Linha 2: ativo 'X(OBRA.Z=1)': OBRA.Z não é parâmetro de obra (existem: OBRA.UTD).",
                "B-004": "Linha 9: apelido 'X' já pertence ao ativo A (linha 2).", "B-005": "Linha 3: CE tem mais de uma declaração (CE=…); use só uma, na primeira linha do ativo.",
                "B-006": "Linha 4: a declaração de CE precisa ser a primeira linha do ativo (linha 2).", "B-007": "Linha 4: parâmetro CABO não está na declaração de N3 (POSTE).",
                "B-008": "Linha 2: padrão CABO=C de N3 não tem nenhum grupo cadastrado.", "B-009": "Linha 5: o COMPONENTE N3 é usado por grupos diferentes (N3(a), N3(b)); cada variante precisa do seu COMPONENTE.",
                "B-010": "Código 7204 aparece nos grupos N3.A (linha 1) e N3.B (linha 2), que podem valer juntos: a quantidade soma duas vezes.",
                "B-012": "Linha 5: composto 'KIT' usa 'X', que não existe na base.", "B-001": "Linha 1: ativo 'N3(': parêntese sem fechar."}
    for cod, msg in amostras.items():
        assert codigo_do_achado(msg) == cod, (cod, msg)
    assert codigo_do_achado("outra coisa") is None and com_codigo("outra coisa") == "outra coisa"
    assert com_codigo(amostras["B-005"]).endswith("[B-005 · docs/base_guia.md]") and com_codigo(com_codigo(amostras["B-005"])).count("B-005") == 1
    assert resumo_erros([amostras["B-005"], amostras["B-008"]]).count("[B-0") == 2
    guia = open(GUIA, encoding="utf-8").read()
    for i in range(1, 14):
        assert f"B-{i:03d}" in guia, i
    assert len(guia.split()) < 2600                                                                   # ~4 páginas
    for secao in ("Ativo simples", "Parâmetro", "Linha viva", "Retirada", "Composto", "OBRA.*", "Armadilhas"):
        assert secao in guia, secao


@pytest.fixture
def banco(tmp_path, monkeypatch):
    monkeypatch.setattr(database, "DB_PATH", str(tmp_path / "t.db"))
    monkeypatch.setattr(config, "DB_PATH", str(tmp_path / "t.db"))
    database.init_db()
    conn = database.get_connection()
    conn.execute("INSERT OR REPLACE INTO projetos (nome, codigo) VALUES ('NOVO', '950')")
    conn.execute("DELETE FROM tabela_orcamento_master")
    conn.commit()
    conn.close()
    return tmp_path


def test_pacote_de_exemplo_importa_e_funciona_sozinho(banco):
    assert "exemplo" in pp.pacotes_embutidos() and pp.pacote_inicial_de("950") is None                  # ponto de partida sugerido, nunca automático
    pacote = pp.ler_pacote(pp.empacotar_embutido("exemplo"))
    assert set(pacote["partes"]) == {"base", "vinculacao", "casos"} and len(pacote["partes"]["casos"]) == 5
    linhas = pp._linhas_do_csv(pacote["partes"]["base"])
    assert 10 <= len({l["ativo"] for l in linhas if l["ativo"]}) <= 20
    prev = pp.previa("950", pacote)
    assert {p["parte"]: p["situacao"] for p in prev["partes"]} == {"base": "novo", "vinculacao": "novo", "casos": "novo"} and not any(p["avisos"] for p in prev["partes"])
    assert prev["casos"]["resumo"] == {"igual": 5, "mudou": 0, "erro": 0}
    pp.aplicar("950", pacote, user_id="u")
    from routers.casos import conferir_projeto
    res = conferir_projeto(pp.projeto_do_cadastro("950"))
    assert res["resumo"] == {"igual": 5, "mudou": 0, "erro": 0}, res
    # o que o guia promete: pedidos de exemplo
    from services import orcamento_calc as calc
    rows = [dict(r) for r in database.get_row_connection().execute("SELECT * FROM tabela_orcamento_master").fetchall()]
    def codigos(texto, op="I"):
        return {(x["operacao"], x["codigo"]): x["total"] for x in calc.processar_calculo([], [{"ativo": texto, "operacao": op, "entidade": "0"}], "NOVO", rows)["resultado"]}
    assert codigos("2-N1")[("I", "7101")] == 2 and codigos("1-N3(A)")[("I", "7204")] == 3 and ("I", "7203") not in codigos("1-N3(A)") and codigos("1-N3")[("I", "7203")] == 3
    assert ("I", "7302") not in codigos("1-N4") and ("I", "7302") in codigos("1-N4", "*I") and ("I", "7304") in codigos("1-N4")
    assert codigos("1-N1", "R")[("R", "7101")] == 1 and codigos("1-KIT_N3")[("I", "7601")] == 2


def test_validador_do_exemplo_sem_erros_e_rotas_dos_embutidos(banco):
    import app as appmod
    from fastapi.testclient import TestClient
    from middleware.auth_middleware import create_jwt_token
    c = TestClient(appmod.app)
    h = {"Authorization": f"Bearer {create_jwt_token('2', 'o@x.com', 'operador')}"}
    ids = {p["id"] for p in c.get("/api/pacote/embutidos", headers=h).json()["pacotes"]}
    assert {"081", "rondonia", "exemplo"} <= ids
    r = c.get("/api/pacote/embutidos/exemplo.zip", headers=h)
    assert r.status_code == 200 and r.content[:2] == b"PK" and c.get("/api/pacote/embutidos/nao.zip", headers=h).status_code == 404


@pytest.mark.skipif(not NODE, reason="sem node")
def test_frase_em_portugues_da_ficha():
    js = ("const E=require('./static/base_explica.js');console.log(JSON.stringify(["
          "E.explicarItem({condicao:'cabo=%A35P%',codigo:'7204',fator_i:3,fator_r:3,mdo:'MATERIAL',lv:'A'}),"
          "E.explicarItem({condicao:'',codigo:'7302',fator_i:1,fator_r:0,mdo:'MATERIAL',lv:'LV'}),"
          "E.explicarItem({itens:[{ativo:'N3(C)',qtd:1},{ativo:'ALCA(C)',qtd:3}]}),"
          "E.explicarItem({condicao:'cabo!=C|A',codigo:'1',fator_i:1.02,fator_r:0,arred:'CIMA',mdo:'MAO-DE-OBRA',filtro:'LV'})]))")
    out = json.loads(subprocess.run([NODE, "-e", js], cwd=config.BASE_DIR, capture_output=True, text=True, timeout=60).stdout)
    assert out[0] == "vale quando cabo contém A35P; soma 3 de 7204 (material) na instalação e na retirada."
    assert "só em linha viva" in out[1] and "não devolve na retirada" in out[1]
    assert out[2] == "composto: pedir este nome equivale a 1 × N3(C) + 3 × ALCA(C)."
    assert "não é nenhum de C ou A" in out[3] and "mão de obra" in out[3] and "arredonda para cima" in out[3] and "filtro: LV" in out[3]


def test_ligacoes_na_tela():
    ler = lambda *p: open(os.path.join(config.STATIC_DIR, *p), encoding="utf-8").read()
    assert "/static/base_explica.js" in ler("orcamento.html") and "O que faz" in ler("base_ficha_ui.js") and "BaseExplica" in ler("base_ficha_ui.js")
    assert "/api/pacote/embutidos" in ler("vinculacao_ui.js") and "Começar de um pacote pronto" in ler("vinculacao_ui.js")
