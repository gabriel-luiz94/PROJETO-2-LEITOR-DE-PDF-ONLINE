"""TASK-031 fase B — pipeline autônomo: ler → montar → ajustar → [confirmar exclusões] → totalizadora → orçamento → obra + pasta. Sem IA."""
import ast
import json
import os

import ezdxf
import pytest

import database
from services.autonomo import execucoes, pipeline as pl, saida
from services.autonomo.leitor_js import LeitorJS
from tests.test_autonomo_leitor import CLS, PROC

def _linha_base(ativo, codigo, fi=1.0, fr=0.5):
    return {"ativo": ativo, "desc_ativo": f"Desc {ativo}", "projeto": "", "codigo": codigo, "desc_codigo": f"Cod {codigo}", "fator_i": fi, "fator_r": fr,
            "mdo": "M1", "componente": ativo, "filtro": "", "origem": ""}


BASE = [_linha_base("U4", "C-U4", 2.0, 1.0), _linha_base("CFU", "C-CFU"), _linha_base("SUPL", "C-SUPL"), _linha_base("U3", "C-U3")]

RECEITA_SUBST = {"id": "AJ1", "nome": "SUP-L", "ativa": True, "regras": [], "acoes": [{"acao": "substituir", "tabela": "outros", "de": "SUP-L", "para": "SUPL"}]}
RECEITA_ADD = {"id": "AJ2", "nome": "SUPL no CFU", "ativa": True, "regras": [],
               "acoes": [{"acao": "adicionar_ativo", "tabela": "outros", "ativo": "SUPL", "qtd": 1,
                          "quando": {"todos": [{"tem": "CFU"}, {"nao": {"tem": "SUPL"}}]}}]}
RECEITA_REMOVE = {"id": "AJ3", "nome": "tira U3", "ativa": True, "regras": [], "acoes": [{"acao": "remover_ativo", "tabela": "outros", "ativo": "U3"}]}
RECEITA_VAZIAS = {"id": "AJ4", "nome": "vazias", "ativa": True, "regras": [], "acoes": [{"acao": "excluir_linhas", "tabela": "ambos", "onde": {"vazias": True}}]}


@pytest.fixture(scope="module")
def leitor():
    lt = LeitorJS()
    yield lt
    lt.fechar()


@pytest.fixture
def ambiente(tmp_path, monkeypatch, leitor):
    monkeypatch.setattr(database, "DB_PATH", str(tmp_path / "t.db"))
    database.init_db()
    conn = database.get_connection()
    conn.execute("INSERT OR IGNORE INTO projetos (nome, codigo) VALUES ('P1', 'P1')")
    conn.commit()
    conn.close()
    return tmp_path


def dxf(caminho, textos):
    doc = ezdxf.new()
    msp = doc.modelspace()
    for i, t in enumerate(textos):
        msp.add_text(t, dxfattribs={"color": 1, "insert": (i * 10, 0), "height": 2})
    doc.saveas(str(caminho))
    return str(caminho)


def contexto(leitor, receitas=(), conversao=None, regras_dominio=()):
    return pl.Contexto(projeto_codigo="P1", projeto_nome="P1", user_id="u1", regras_proc=PROC, regras_cls=CLS,
                       regras_conversao=conversao if conversao is not None else [], base_orcamento=BASE, regras_dominio=list(regras_dominio),
                       receitas=list(receitas), leitor=leitor)


def rodar(ambiente, leitor, textos, receitas=(), nome="proj.dxf", **kw):
    arq = dxf(ambiente / nome, textos)
    ctx = contexto(leitor, receitas, kw.pop("conversao", None))
    ex = pl.processar_arquivo(arq, "P1", "u1", str(ambiente / "saida"), ctx=ctx, **kw)
    return ex, ctx


def ativos(tabela):
    return [r["ativo"] for r in tabela]


def test_pipeline_completo_sem_exclusao_gera_obra_orcamento_e_pasta(ambiente, leitor):
    ex, _ = rodar(ambiente, leitor, ["1-U4 1-SUP-L", "2-CFU"], [RECEITA_SUBST, RECEITA_ADD])
    assert ex["status"] == "ok", ex["mensagem"]
    # tabelas finais = do ajuste (SUP-L→SUPL e SUPL no CFU), vistas pela obra salva
    conn = database.get_row_connection()
    obra = conn.execute("SELECT * FROM obras WHERE id = ?", (ex["obra_id"],)).fetchone()
    conn.close()
    assert obra["projeto"] == "P1" and obra["user_id"] == "u1" and obra["nome"].startswith("[Autônomo]")
    snap = json.loads(obra["dados_json"])
    assert snap["autonomo"]["execucao_id"] == ex["id"]
    outros = [r["ativo"] for r in snap["outros"]["data"]]
    assert "1-SUPL" in " ".join(outros) and not any("SUP-L" in a for a in outros)
    assert snap["totalizadora"]["data"] and {"id", "ativo", "qtd", "origem"} <= set(snap["totalizadora"]["data"][0])
    # pasta de saída
    pasta = ex["pasta_saida"]
    assert {"cabos.csv", "outros.csv", "tabelas.json", "orcamento.json", "orcamento.csv", "relatorio.json", "pendencias.json"} <= set(os.listdir(pasta))
    rel = json.load(open(os.path.join(pasta, "relatorio.json"), encoding="utf-8"))
    assert [e["etapa"] for e in rel["etapas"]] == ["ler", "montar", "ramais", "validar_antes", "ajustar", "aplicar", "validar_depois", "totalizadora", "orcamento", "salvar"]
    assert rel["status"] == "ok" and rel["ajustes"]["aplicados"]
    orc = json.load(open(os.path.join(pasta, "orcamento.json"), encoding="utf-8"))
    assert {l["codigo"] for l in orc["resultado"]} >= {"C-U4", "C-CFU"}


def test_exclusao_fica_aguardando_e_nada_e_salvo_nem_orcado(ambiente, leitor):
    ex, _ = rodar(ambiente, leitor, ["1-U3 1-U4", "2-CFU", "1-U4 1-SUP-L"], [RECEITA_SUBST, RECEITA_REMOVE])
    assert ex["status"] == "aguardando_confirmacao"
    assert ex["obra_id"] is None
    conn = database.get_row_connection()
    assert conn.execute("SELECT COUNT(*) FROM obras").fetchone()[0] == 0
    conn.close()
    pend = ex["decisoes"]["pendentes"]
    assert pend and all("Editar" in ex["diff"]["frases"][i] or "Excluir" in ex["diff"]["frases"][i] for i in pend)
    assert "orcamento.json" not in os.listdir(ex["pasta_saida"])
    lista = json.load(open(os.path.join(ex["pasta_saida"], "pendencias.json"), encoding="utf-8"))
    assert [p["indice"] for p in lista] == pend


def test_sim_para_todos_conclui_com_as_exclusoes(ambiente, leitor):
    ex, ctx = rodar(ambiente, leitor, ["1-U3 1-U4", "1-U4 1-SUP-L"], [RECEITA_SUBST, RECEITA_REMOVE])
    final = pl.confirmar(ex["id"], None, ctx=ctx)           # "Sim para todos" (todas as exclusões deste arquivo)
    assert final["status"] == "ok" and final["obra_id"]
    snap = json.loads(_obra(final["obra_id"])["dados_json"])
    assert not any("U3" in r["ativo"] for r in snap["outros"]["data"])
    assert any("SUPL" in r["ativo"] for r in snap["outros"]["data"])    # e o ajuste que não exclui também entrou
    assert final["decisoes"]["confirmadas"] == ex["decisoes"]["pendentes"]


def _obra(obra_id):
    conn = database.get_row_connection()
    r = conn.execute("SELECT * FROM obras WHERE id = ?", (obra_id,)).fetchone()
    conn.close()
    return r


def test_sim_individual_espera_as_demais_e_so_entao_conclui(ambiente, leitor):
    ex, ctx = rodar(ambiente, leitor, ["1-U3 1-U4", "1-U3 1-CFU", "1-U4"], [RECEITA_REMOVE])
    pend = ex["decisoes"]["pendentes"]
    assert len(pend) == 2
    parcial = pl.confirmar(ex["id"], [pend[0]], ctx=ctx)
    assert parcial["status"] == "aguardando_confirmacao" and parcial["obra_id"] is None
    final = pl.confirmar(ex["id"], [pend[1]], ctx=ctx)
    assert final["status"] == "ok"
    with pytest.raises(pl.ErroPipeline):                    # não dá para decidir duas vezes
        pl.confirmar(ex["id"], None, ctx=ctx)


def test_rejeitar_mantem_a_linha_e_marca_com_pendencias(ambiente, leitor):
    ex, ctx = rodar(ambiente, leitor, ["1-U3 1-U4", "1-U4 1-SUP-L"], [RECEITA_REMOVE, RECEITA_SUBST])
    final = pl.rejeitar(ex["id"], None, ctx=ctx)
    assert final["status"] == "com_pendencias"
    snap = json.loads(_obra(final["obra_id"])["dados_json"])
    assert any("U3" in r["ativo"] for r in snap["outros"]["data"])


def test_indice_sem_pendencia_e_recusado(ambiente, leitor):
    ex, ctx = rodar(ambiente, leitor, ["1-U3 1-U4"], [RECEITA_REMOVE])
    with pytest.raises(pl.ErroPipeline, match="sem pendência"):
        pl.confirmar(ex["id"], [999], ctx=ctx)


def test_reverter_refaz_a_obra_com_as_tabelas_originais(ambiente, leitor):
    ex, ctx = rodar(ambiente, leitor, ["1-U4 1-SUP-L"], [RECEITA_SUBST])
    assert ex["status"] == "ok"
    ver = pl.reverter(ex["id"], ctx=ctx)
    assert ver["status"] == "revertida"
    snap = json.loads(_obra(ver["obra_id"])["dados_json"])
    assert any("SUP-L" in r["ativo"] for r in snap["outros"]["data"]) and snap["autonomo"]["revertida"] is True
    with pytest.raises(pl.ErroPipeline):
        pl.reverter(ex["id"], ctx=ctx)


def test_arquivo_duplicado_nao_e_processado_duas_vezes_mas_pode_forcar(ambiente, leitor):
    ex, ctx = rodar(ambiente, leitor, ["1-U4"], [])
    arq = str(ambiente / "proj.dxf")
    de_novo = pl.processar_arquivo(arq, "P1", "u1", str(ambiente / "saida"), ctx=ctx)
    assert de_novo.get("duplicado") and de_novo["id"] == ex["id"]
    forcado = pl.processar_arquivo(arq, "P1", "u1", str(ambiente / "saida"), ctx=ctx, forcar=True)
    assert forcado["id"] != ex["id"] and not forcado.get("duplicado")


def test_arquivo_corrompido_vira_erro_sem_derrubar(ambiente, leitor):
    ruim = ambiente / "quebrado.dxf"
    ruim.write_text("isto não é um DXF", encoding="utf-8")
    ex = pl.processar_arquivo(str(ruim), "P1", "u1", str(ambiente / "saida"), ctx=contexto(leitor))
    assert ex["status"] == "erro" and "ler o arquivo" in ex["mensagem"]
    assert ruim.exists()                                   # o original nunca é apagado


def test_extensao_nao_suportada_e_projeto_sem_regras_do_leitor(ambiente, leitor):
    txt = ambiente / "a.txt"
    txt.write_text("x", encoding="utf-8")
    assert pl.processar_arquivo(str(txt), "P1", "u1", str(ambiente / "saida"), ctx=contexto(leitor))["status"] == "erro"
    arq = dxf(ambiente / "b.dxf", ["1-U4"])
    ctx = contexto(leitor)
    ctx.regras_proc, ctx.regras_cls = [], []
    ex = pl.processar_arquivo(arq, "P1", "u1", str(ambiente / "saida"), ctx=ctx)
    assert ex["status"] == "erro" and "regras do leitor" in ex["mensagem"]


def test_erro_de_validacao_sem_correcao_salva_com_pendencias(ambiente, leitor):
    ex, _ = rodar(ambiente, leitor, ["1-U4", "1-N1"], [])      # N1 não existe na base de orçamento: nenhum ajuste resolve
    assert ex["status"] == "com_pendencias" and ex["obra_id"]
    assert ex["relatorio"]["orcamento"]["nao_encontrados"] == ["N1"]


def test_ajuste_invalido_e_ignorado_e_relatado(ambiente, leitor):
    ruim = {"id": "AJ-RUIM", "nome": "ruim", "ativa": True, "regras": [], "acoes": [{"acao": "inexistente", "tabela": "outros"}]}
    ex, _ = rodar(ambiente, leitor, ["1-U4"], [ruim, RECEITA_SUBST])
    assert ex["status"] in ("ok", "com_pendencias")
    assert ex["relatorio"]["ajustes_ignorados"][0]["ajuste"] == "AJ-RUIM"


def test_regras_de_conversao_valem_no_orcamento(ambiente, leitor):
    conv = [{"origem": "OUTROS", "op_de": "", "ativo_de": "U4", "acao": "SUBST", "op_para": "", "ativo_para": "", "fator": 3, "arredondamento": "NORMAL",
             "val_min": "", "val_max": ""}]
    ex, _ = rodar(ambiente, leitor, ["2-U4"], [], conversao=conv)
    orc = json.load(open(os.path.join(ex["pasta_saida"], "orcamento.json"), encoding="utf-8"))
    linha = next(l for l in orc["resultado"] if l["codigo"] == "C-U4" and l["operacao"] == "I")
    assert linha["total"] == 12.0          # 2 (qtd) × 3 (conversão) × 2.0 (fator_i da base)


def test_historico_e_listagem(ambiente, leitor):
    rodar(ambiente, leitor, ["1-U4"], [])
    rodar(ambiente, leitor, ["1-U3 1-U4"], [RECEITA_REMOVE], nome="outro.dxf")
    todos = execucoes.listar()
    assert len(todos) == 2 and {x["status"] for x in todos} >= {"aguardando_confirmacao"}
    assert [x["status"] for x in execucoes.listar("aguardando_confirmacao")] == ["aguardando_confirmacao"]
    assert "originais" not in todos[0]


def test_nome_seguro_e_pasta_nao_escapa_da_base(tmp_path):
    p = saida.pasta_da_execucao(str(tmp_path), "../../etc", "../../x/passwd.dxf", "abc")
    assert os.path.commonpath([p, str(tmp_path)]) == str(tmp_path) and ".." not in os.path.relpath(p, str(tmp_path))
    assert saida.nome_seguro("a/b\\c:d*?.dxf") == "a_b_c_d_.dxf" and saida.nome_seguro("..") == "item"


def test_o_pipeline_nao_chama_ia(ambiente, leitor, monkeypatch):
    import services.correcao_ia as ci
    import services.validacao_ia as vi
    chamadas = []
    monkeypatch.setattr(vi, "chamar_gemini", lambda *a, **k: chamadas.append("gemini"))
    monkeypatch.setattr(ci, "ajustar_com_ia", lambda *a, **k: chamadas.append("ajuste"))
    monkeypatch.setattr(ci, "corrigir_com_ia", lambda *a, **k: chamadas.append("correcao"))
    rodar(ambiente, leitor, ["1-U4", "2-CFU"], [RECEITA_ADD])
    assert chamadas == []
    # e nenhum módulo do pipeline importa IA
    proibidos = ("ai_chat", "validacao_ia", "correcao_ia", "genai", "openai")
    for nome in ("pipeline", "aplicar", "montagem", "totalizadora", "saida", "execucoes", "leitor_js"):
        arvore = ast.parse(open(f"services/autonomo/{nome}.py", encoding="utf-8").read())
        mods = [n.module or "" for n in ast.walk(arvore) if isinstance(n, ast.ImportFrom)] + [a.name for n in ast.walk(arvore) if isinstance(n, ast.Import) for a in n.names]
        assert not [m for m in mods if any(p in m for p in proibidos)], nome


def test_carregar_contexto_le_o_banco(ambiente):
    ctx = pl.carregar_contexto("P1", "u1")
    assert ctx.projeto_nome == "P1" and ctx.regras_conversao == [pl.REGRA_PADRAO]
    with pytest.raises(pl.ErroPipeline, match="não cadastrado"):
        pl.carregar_contexto("NAOEXISTE", "u1")


RECEITA_DUPLICADAS = {"id": "AJ5", "nome": "dup", "ativa": True, "regras": [], "acoes": [{"acao": "excluir_linhas", "tabela": "outros", "onde": {"duplicadas": True}}]}


def test_excluir_linhas_tambem_pede_confirmacao_e_o_nao_destrutivo_passa(ambiente, leitor):
    ex, ctx = rodar(ambiente, leitor, ["1-U4", "1-U4", "2-CFU"], [RECEITA_DUPLICADAS, RECEITA_ADD])
    assert ex["status"] == "aguardando_confirmacao"
    frases = [ex["diff"]["frases"][i] for i in ex["decisoes"]["pendentes"]]
    assert len(frases) == 1 and frases[0].startswith("Excluir")
    # a edição não destrutiva (SUPL no CFU) NÃO está pendente: entra sozinha quando termina
    final = pl.confirmar(ex["id"], None, ctx=ctx)
    snap = json.loads(_obra(final["obra_id"])["dados_json"])
    assert len(snap["outros"]["data"]) == 2 and any("SUPL" in r["ativo"] for r in snap["outros"]["data"])


def test_sem_ajustes_habilitados_o_pipeline_so_monta_e_orca(ambiente, leitor):
    ex, _ = rodar(ambiente, leitor, ["1-U4", "2-CFU"], [])
    assert ex["status"] == "ok" and ex["decisoes"]["pendentes"] == [] and ex["relatorio"]["ajustes"]["aplicados"] == []


def _receitas_com(n_acoes_por_receita, n_receitas):
    return [{"id": f"AJ{i}", "nome": "x", "ativa": True, "regras": [], "acoes": [{"acao": "normalizar", "tabela": "outros", "regras": ["espacos"]}] * n_acoes_por_receita}
            for i in range(n_receitas)]


def test_projeto_real_com_mais_de_50_acoes_habilitadas_funciona(ambiente, leitor):
    """Caso real (projeto 027): os ajustes habilitados somavam 53 ações e o limite de 50 das rotas barrava o arquivo inteiro."""
    ex, _ = rodar(ambiente, leitor, ["1-U4 1-SUP-L"], _receitas_com(10, 6) + [RECEITA_SUBST])
    assert ex["status"] in ("ok", "com_pendencias"), ex["mensagem"]
    etapa = next(e for e in ex["relatorio"]["etapas"] if e["etapa"] == "ajustar")
    assert etapa["status"] == "ok" and not any("SUP-L" in r["ativo"] for r in json.loads(_obra(ex["obra_id"])["dados_json"])["outros"]["data"])


def test_ajustes_demais_na_cadeia_viram_erro_claro(ambiente, leitor):
    ex, _ = rodar(ambiente, leitor, ["1-U4"], _receitas_com(10, 31))      # 310 ações > 300
    assert ex["status"] == "erro" and "máx. 300" in ex["mensagem"]


def test_ajuste_descartado_pela_camada_1_marca_com_pendencias(ambiente, leitor):
    quebra = {"id": "AJ-QUEBRA", "nome": "quebra", "ativa": True, "regras": [],
              "acoes": [{"acao": "substituir", "tabela": "outros", "de": "1-U4", "para": "U4"}]}   # deixaria o token sem quantidade
    ex, _ = rodar(ambiente, leitor, ["1-U4"], [quebra])
    assert ex["status"] == "com_pendencias" and ex["relatorio"]["ajustes"]["descartados"]
    snap = json.loads(_obra(ex["obra_id"])["dados_json"])
    assert ativos(snap["outros"]["data"]) == ["1-U4"]     # a linha ficou como estava


def test_ramais_entram_no_orcamento_autonomo(ambiente, leitor):
    base = BASE + [_linha_base("MAC", "C-MAC"), _linha_base("MAA", "C-MAA")]
    arq = dxf(ambiente / "ramais.dxf", ["1-U4", "TROCAR 3 RS M AC"])
    ctx = contexto(leitor)
    ctx.base_orcamento = base
    ex = pl.processar_arquivo(arq, "P1", "u1", str(ambiente / "saida"), ctx=ctx)
    assert ex["status"] == "ok", ex["mensagem"]
    etapa = next(e for e in ex["relatorio"]["etapas"] if e["etapa"] == "ramais")
    assert etapa["detalhe"] == {"itens_ramais": 1, "linhas_geradas": 2}
    snap = json.loads(_obra(ex["obra_id"])["dados_json"])
    geradas = [r for r in snap["outros"]["data"] if r.get("texto") == "RAMAIS (GERADO)"]
    assert [(r["operacao"], r["ativo"]) for r in geradas] == [("I", "60-MAC"), ("R", "45-MAC")]
    codigos = {(l["codigo"], l["operacao"]) for l in json.load(open(os.path.join(ex["pasta_saida"], "orcamento.json"), encoding="utf-8"))["resultado"]}
    assert ("C-MAC", "I") in codigos and ("C-MAC", "R") in codigos
