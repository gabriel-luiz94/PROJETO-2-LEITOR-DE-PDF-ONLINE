"""TASK-191 — assistente de IA com ferramentas e skills.

Sem rede: os SDKs (Gemini/OpenAI/Claude) são trocados por fakes; as ferramentas rodam sobre uma base pequena em memória.
"""
import asyncio
import json
import types

import pytest
from fastapi.testclient import TestClient

import app as appmod
import config
import database
import routers.ai_chat as ai_chat
from middleware.auth_middleware import create_jwt_token
from services.assistente import agente, catalogo, ferramentas, skills
from services.autonomo import pipeline as pl


def _linha(ativo, codigo, fi=1.0, fr=0.5, projeto="", desc=None):
    return {"ativo": ativo, "desc_ativo": desc or f"Desc {ativo}", "projeto": projeto, "codigo": codigo, "desc_codigo": f"Cod {codigo}", "fator_i": fi, "fator_r": fr,
            "mdo": "M1", "componente": ativo, "filtro": "", "origem": ""}


BASE = [_linha("DT11/300", "C-POSTE", desc="POSTE DT 11/300"), _linha("TR125", "C-TR", desc="TRAFO MONO 25"), _linha("CFU", "C-CFU", desc="CHAVE FUSIVEL"),
        _linha("PR220", "C-PR"), _linha("CAA2", "C-CAA2", desc="CABO CAA 2"), _linha("CAA2", "C-CAA2B", fi=2.0), _linha("X9", "C-X9", projeto="RONDONIA")]


def pipeline_fake(base=None):
    return pl.Contexto(projeto_codigo="081", projeto_nome="PERNAMBUCO", user_id="u1", regras_conversao=[], base_orcamento=base if base is not None else BASE)


def ctx_com(cabos=(), outros=(), base=None, **kw):
    p = pipeline_fake(base)
    return ferramentas.Contexto(projeto_codigo="081", projeto_nome="PERNAMBUCO", user={"user_id": "u1", "role": "admin"}, cabos=list(cabos), outros=list(outros),
                                carregar_pipeline=lambda: p, **kw)


@pytest.fixture
def pastas(tmp_path, monkeypatch):
    pacote, usuario = tmp_path / "pacote", tmp_path / "usuario"
    (pacote / "_geral").mkdir(parents=True)
    (pacote / "pernambuco").mkdir()
    (pacote / "_geral" / "orcamento.md").write_text("descricao: como orçar\n# Orçar\npasso 1\n", encoding="utf-8")
    (pacote / "_geral" / "formato.md").write_text("# Formato\nTexto sem descricao explicita.\n", encoding="utf-8")
    (pacote / "pernambuco" / "orcamento.md").write_text("descricao: orçar em PE\nregra de PE\n", encoding="utf-8")
    monkeypatch.setattr(config, "SKILLS_DIR_PACOTE", str(pacote))
    monkeypatch.setattr(config, "SKILLS_DIR_USUARIO", str(usuario))
    return pacote, usuario


# ── skills ──────────────────────────────────────────────────────────────────────────────────────────
def test_slug_do_projeto():
    assert skills.slug("PERNAMBUCO") == "pernambuco"
    assert skills.slug("Rondônia 2") == "rondonia_2"
    assert skills.slug("") == "" and skills.slug(None) == ""


def test_listar_projeto_vence_geral_e_descricao(pastas):
    r = {s["nome"]: s for s in skills.listar("pernambuco")}
    assert set(r) == {"orcamento", "formato"}
    assert r["orcamento"]["escopo"] == "projeto" and r["orcamento"]["descricao"] == "orçar em PE"
    assert r["formato"]["descricao"] == "Texto sem descricao explicita."        # sem `descricao:` usa a 1ª linha de texto
    assert {s["nome"]: s["escopo"] for s in skills.listar("")} == {"orcamento": "geral", "formato": "geral"}


def test_carregar_tira_a_linha_de_descricao_e_respeita_precedencia(pastas):
    assert "descricao" not in skills.carregar("pernambuco", "orcamento").lower()
    assert skills.carregar("pernambuco", "orcamento").endswith("regra de PE")
    assert "passo 1" in skills.carregar("outro", "orcamento")                     # projeto sem skill própria cai na geral
    assert skills.carregar("pernambuco", "ORCAMENTO.md")                          # tolera caixa e extensão
    with pytest.raises(skills.ErroSkill):
        skills.carregar("pernambuco", "nao_existe")


def test_salvar_do_usuario_vence_o_pacote_e_apagar_volta(pastas):
    pacote, usuario = pastas
    skills.salvar("_geral", "orcamento", "descricao: meu jeito\nmeu texto")
    assert skills.carregar("", "orcamento") == "meu texto"
    assert next(s for s in skills.listar("") if s["nome"] == "orcamento")["origem"] == "usuario"
    assert (pacote / "_geral" / "orcamento.md").read_text(encoding="utf-8").startswith("descricao: como orçar")      # o pacote não foi tocado
    assert skills.apagar("_geral", "orcamento") is True
    assert "passo 1" in skills.carregar("", "orcamento")
    assert skills.apagar("_geral", "orcamento") is False


@pytest.mark.parametrize("escopo,nome", [("_geral", "../fora"), ("_geral", "A B"), ("_geral", ""), ("..", "x"), ("a/b", "x"), ("_geral", "x" * 61), ("_geral", ".oculta")])
def test_nomes_perigosos_sao_recusados(pastas, escopo, nome):
    with pytest.raises(skills.ErroSkill):
        skills.salvar(escopo, nome, "texto")
    with pytest.raises(skills.ErroSkill):
        skills.texto_bruto(escopo, nome)


def test_salvar_recusa_vazio_e_grande(pastas):
    with pytest.raises(skills.ErroSkill):
        skills.salvar("_geral", "x", "  \n")
    with pytest.raises(skills.ErroSkill):
        skills.salvar("_geral", "x", "a" * (skills.MAX_CHARS + 1))


def test_indice_para_prompt_lista_nome_e_descricao(pastas):
    t = skills.indice_para_prompt("pernambuco")
    assert "- orcamento: orçar em PE" in t and "carregar_skill" in t
    assert skills.indice_para_prompt("") != "" and "orçar em PE" not in skills.indice_para_prompt("")


def test_skills_do_pacote_do_repositorio_existem_e_tem_descricao():
    # as skills que acompanham o programa (data/skills) precisam estar lá e legíveis
    nomes = {s["nome"]: s for s in skills.listar("pernambuco")}
    assert {"formato_ativos", "montar_projeto", "orcamento", "projeto"} <= set(nomes)
    assert all(s["descricao"] for s in nomes.values())
    assert "CAA2 3 35 m" in skills.carregar("pernambuco", "formato_ativos")


# ── catálogo ────────────────────────────────────────────────────────────────────────────────────────
def test_indexar_filtra_pelo_projeto():
    i = catalogo.indexar(BASE, "PERNAMBUCO", "081")
    assert "X9" not in i and {"DT11/300", "TR125", "CFU", "PR220", "CAA2"} <= set(i)
    assert i["CAA2"]["componentes"] == 2 and i["CAA2"]["codigos"] == 2
    assert "X9" in catalogo.indexar(BASE, "RONDONIA", "229")


def test_indexar_usa_o_nome_base_de_ativos_com_parametros():
    base = [_linha("N1=N1(cabo=XLPE, poste=DT)", "C-A"), _linha("N1(cabo=XLPE)", "C-B"), _linha("N1", "C-C")]
    i = catalogo.indexar(base, "PERNAMBUCO", "081")
    assert list(i) == ["N1"] and i["N1"]["componentes"] == 3 and len(i["N1"]["formas"]) == 2


def test_buscar_ordena_do_mais_exato_ao_mais_frouxo():
    i = catalogo.indexar(BASE, "PERNAMBUCO", "081")
    assert [e["nome"] for e in catalogo.buscar(i, "cfu")] == ["CFU"]
    assert catalogo.buscar(i, "poste 11 300")[0]["nome"] == "DT11/300"           # várias palavras, nome + descrição
    assert [e["nome"] for e in catalogo.buscar(i, "TR")][0] == "TR125"
    assert catalogo.buscar(i, "inexistente") == [] and catalogo.buscar(i, "  ") == []


def test_gerar_markdown_e_deterministico_e_descreve():
    i = catalogo.indexar(BASE, "PERNAMBUCO", "081")
    a, b = catalogo.gerar_markdown(i, "PERNAMBUCO"), catalogo.gerar_markdown(i, "PERNAMBUCO")
    assert a == b and a.startswith("descricao: Catálogo dos 5 ativos") and "`DT11/300` — POSTE DT 11/300" in a and "**DT**" in a


def test_familia():
    assert [catalogo.familia(x) for x in ("DT11/300", "2ESTR", "C16", "123")] == ["DT", "ESTR", "C", "#"]


# ── ferramentas ─────────────────────────────────────────────────────────────────────────────────────
def test_buscar_e_detalhar_ativo():
    c = ctx_com()
    r = ferramentas.executar(c, "buscar_ativo", {"texto": "chave"})
    assert r["ativos"][0]["ativo"] == "CFU"
    assert "Nenhum ativo" in ferramentas.executar(c, "buscar_ativo", {"texto": "zzz"})["aviso"]
    assert "erro" in ferramentas.executar(c, "buscar_ativo", {"texto": ""})
    d = ferramentas.executar(c, "detalhar_ativo", {"ativo": "caa2"})
    assert len(d["componentes"]) == 2 and d["componentes"][0]["codigo"] == "C-CAA2"
    e = ferramentas.executar(c, "detalhar_ativo", {"ativo": "TR12"})
    assert "não está na base" in e["erro"]


def test_ler_tabela_ignora_congeladas_e_corta():
    c = ctx_com(cabos=[{"id": 0, "operacao": "I", "ativo": "CAA2 3 35 m"}, {"id": 1, "operacao": "I", "ativo": "x", "congelada": True}],
                outros=[{"id": i, "operacao": "I", "ativo": f"{i}-CFU"} for i in range(ferramentas.LIMITE_LINHAS + 5)])
    r = ferramentas.executar(c, "ler_tabela", {})
    assert r["cabos"]["total"] == 1 and r["outros"]["cortado"] is True
    assert sum(len(t["linhas"]) for t in r.values()) == ferramentas.LIMITE_LINHAS
    assert "erro" in ferramentas.executar(c, "ler_tabela", {"tabela": "x"})


def test_propor_alteracoes_valida_ids_e_formato_e_confere_a_base():
    c = ctx_com(cabos=[{"id": 0, "operacao": "I", "ativo": "CAA2 3 35 m"}], outros=[{"id": 0, "operacao": "I", "ativo": "DT11/300, 1 CFU"}, {"id": 1, "operacao": "I", "ativo": "1-PR220"}])
    r = ferramentas.executar(c, "propor_alteracoes", {"resumo": "ajuste", "alteracoes": [
        {"acao": "adicionar", "tabela": "outros", "operacao": "i", "ativo": "DT11/300, 1 TR125"},
        {"acao": "editar", "tabela": "outros", "id": 0, "operacao": "I", "ativo": "DT11/300, 2 CFU"},
        {"acao": "remover", "tabela": "outros", "id": 1},
        {"acao": "remover", "tabela": "outros", "id": 99},
        {"acao": "editar", "tabela": "cabos", "id": 0, "ativo": ""},
        {"acao": "adicionar", "tabela": "outros", "ativo": "1-ZZZ"},
        {"acao": "adicionar", "tabela": "outros", "ativo": "CAA2 35 m"},
        {"acao": "voar", "tabela": "outros"},
    ]})
    assert r["aceitas"] == 5 and len(r["rejeitadas"]) == 3 and r["total_no_plano"] == 5
    assert any("não existe" in x["motivo"] for x in r["rejeitadas"]) and any("ativo vazio" in x["motivo"] for x in r["rejeitadas"])
    plano = {p["n"]: p for p in c.plano}
    assert plano[1]["operacao"] == "I" and plano[3]["antes"] == {"operacao": "I", "ativo": "1-PR220"} and plano[3]["ativo"] == "1-PR220"
    assert any("1-ZZZ" in a and "não encontrado" in a for a in r["avisos"])
    assert any("formato de cabo" in a for a in r["avisos"])
    assert c.resumo_plano == "ajuste"


def test_propor_alteracoes_nao_repete_a_mesma_linha_e_substitui():
    c = ctx_com(outros=[{"id": 0, "operacao": "I", "ativo": "1-CFU"}])
    ferramentas.executar(c, "propor_alteracoes", {"alteracoes": [{"acao": "remover", "tabela": "outros", "id": 0}]})
    r = ferramentas.executar(c, "propor_alteracoes", {"alteracoes": [{"acao": "editar", "tabela": "outros", "id": 0, "ativo": "2-CFU"}]})
    assert r["aceitas"] == 0 and "já está no plano" in r["rejeitadas"][0]["motivo"]
    r = ferramentas.executar(c, "propor_alteracoes", {"substituir": True, "alteracoes": [{"acao": "editar", "tabela": "outros", "id": 0, "ativo": "2-CFU"}]})
    assert r["aceitas"] == 1 and len(c.plano) == 1 and c.plano[0]["acao"] == "editar"


def test_propor_alteracoes_sem_lista_e_limite():
    c = ctx_com()
    assert "erro" in ferramentas.executar(c, "propor_alteracoes", {})
    r = ferramentas.executar(c, "propor_alteracoes", {"alteracoes": [{"acao": "adicionar", "tabela": "outros", "ativo": "1-CFU"}] * (ferramentas.LIMITE_PLANO + 10)})
    assert r["aceitas"] == ferramentas.LIMITE_PLANO and any("até" in x["motivo"] for x in r["rejeitadas"])


def test_validar_linhas_aponta_o_que_a_base_nao_conhece():
    c = ctx_com()
    r = ferramentas.executar(c, "validar_linhas", {"linhas": [{"tabela": "outros", "ativo": "DT11/300, 1 CFU"}, {"tabela": "outros", "ativo": "1-ZZZ"}]})
    assert r["tudo_certo"] is False and any("ZZZ" in p["ativo"] for p in r["problemas"]) and all("CFU" not in p["ativo"] for p in r["problemas"])
    assert ferramentas.executar(c, "validar_linhas", {"linhas": [{"tabela": "outros", "ativo": "2-CFU"}]})["tudo_certo"] is True
    assert "erro" in ferramentas.executar(c, "validar_linhas", {})
    # formatos misturados (vírgula + «1-CFU») são apontados com o motivo da própria Totalizadora
    mix = ferramentas.executar(c, "validar_linhas", {"linhas": [{"tabela": "outros", "ativo": "DT11/300, 1-CFU"}]})
    assert mix["tudo_certo"] is False and "formato antigo" in mix["problemas"][0]["motivo"]
    plano = ferramentas.executar(c, "propor_alteracoes", {"alteracoes": [{"acao": "adicionar", "tabela": "outros", "ativo": "DT11/300, 1-CFU"}]})
    assert any("formato antigo" in a for a in plano["avisos"])


def test_totalizar_soma_e_simula_o_plano():
    c = ctx_com(outros=[{"id": 0, "operacao": "I", "ativo": "2-CFU"}, {"id": 1, "operacao": "I", "ativo": "1-CFU 1-ZZZ"}])
    r = ferramentas.executar(c, "totalizar", {})
    cfu = next(i for i in r["itens"] if i["ativo"] == "CFU")
    assert cfu["qtd"] == 3 and r["nao_encontrados"] == ["ZZZ"] and r["com_plano"] is False
    ferramentas.executar(c, "propor_alteracoes", {"alteracoes": [{"acao": "adicionar", "tabela": "outros", "ativo": "4-CFU"}, {"acao": "remover", "tabela": "outros", "id": 0}]})
    r2 = ferramentas.executar(c, "totalizar", {"com_plano": True})
    assert next(i for i in r2["itens"] if i["ativo"] == "CFU")["qtd"] == 5        # 3 − 2 + 4
    assert next(i for i in ferramentas.executar(c, "totalizar", {})["itens"] if i["ativo"] == "CFU")["qtd"] == 3      # sem o plano nada mudou


def test_totalizar_expande_cabo_em_notacao_de_label():
    c = ctx_com(cabos=[{"id": 0, "entidade": "CABO", "operacao": "I", "ativo": "3xCAA2(2)+2 35 M"}], base=BASE)
    r = ferramentas.executar(c, "totalizar", {})
    item = next(i for i in r["itens"] if i["ativo"] == "CAA2")
    assert item["linhas"] == 5 and item["qtd"] == 175


def test_calcular_orcamento_devolve_linhas_e_nao_encontrados(monkeypatch):
    monkeypatch.setattr("routers.vinculacao_config.postes_do_projeto", lambda x: None)
    monkeypatch.setattr("routers.vinculacao_config.perfil_do_projeto", lambda x: None)
    c = ctx_com(outros=[{"id": 0, "operacao": "I", "ativo": "2-CFU 1-ZZZ"}])
    r = ferramentas.executar(c, "calcular_orcamento", {})
    assert r["linhas_do_orcamento"] >= 1 and any(i["codigo"] == "C-CFU" and i["total"] == 2 for i in r["itens"])
    assert "ZZZ" in r["nao_encontrados"]
    ferramentas.executar(c, "propor_alteracoes", {"alteracoes": [{"acao": "adicionar", "tabela": "outros", "ativo": "3-CFU"}]})
    r2 = ferramentas.executar(c, "calcular_orcamento", {"com_plano": True})
    assert any(i["codigo"] == "C-CFU" and i["total"] == 5 for i in r2["itens"])


def test_erro_de_ferramenta_vira_resultado_nao_excecao():
    c = ferramentas.Contexto(projeto_codigo="081", projeto_nome="PERNAMBUCO", carregar_pipeline=lambda: (_ for _ in ()).throw(RuntimeError("banco fora")))
    assert "banco fora" in ferramentas.executar(c, "buscar_ativo", {"texto": "cfu"})["erro"]
    assert "não existe" in ferramentas.executar(c, "apagar_tudo", {})["erro"]
    assert "JSON" in ferramentas.executar(c, "ler_tabela", "{quebrado")["erro"]
    assert ferramentas.executar(c, "ler_tabela", '{"tabela": "cabos"}')["cabos"]["total"] == 0           # argumentos em texto JSON também servem


def test_obras_listar_consultar_propor():
    obra = {"id": "o1", "nome": "Rede Norte", "tipo": "obra", "dados_json": json.dumps({"cabos": {"data": []}, "outros": {"data": [{"operacao": "I", "ativo": "1-CFU"}]}})}
    modelo = {"id": "m1", "nome": "Modelo", "tipo": "modelo", "dados_json": "{}"}
    todas = {"o1": obra, "m1": modelo}
    c = ctx_com(carregar_obras=lambda: [dict(o, data="d", publica=False, minha=True, dono="eu") for o in todas.values()], buscar_obra=lambda i: todas.get(i))
    assert [o["nome"] for o in ferramentas.executar(c, "listar_obras", {})["obras"]] == ["Rede Norte", "Modelo"]
    assert "1-CFU" in ferramentas.executar(c, "consultar_obra", {"obra_id": "o1"})["conteudo"]
    assert "erro" in ferramentas.executar(c, "consultar_obra", {"obra_id": "zz"})
    assert "erro" in ferramentas.executar(c, "propor_obra", {"obra_id": "o1", "modo": "apagar"})
    assert "erro" in ferramentas.executar(c, "propor_obra", {"obra_id": "zz", "modo": "adicionar"})
    assert ferramentas.executar(c, "propor_obra", {"obra_id": "m1", "modo": "adicionar"})["aviso"].startswith("É um MODELO")
    assert c.plano[0]["acao"] == "obra" and c.plano[0]["obra_id"] == "m1"
    # não mistura obra com alterações (nos dois sentidos)
    assert "erro" in ferramentas.executar(c, "propor_alteracoes", {"alteracoes": [{"acao": "adicionar", "tabela": "outros", "ativo": "1-CFU"}]})
    c2 = ctx_com(buscar_obra=lambda i: todas.get(i))
    ferramentas.executar(c2, "propor_alteracoes", {"alteracoes": [{"acao": "adicionar", "tabela": "outros", "ativo": "1-CFU"}]})
    assert "erro" in ferramentas.executar(c2, "propor_obra", {"obra_id": "o1", "modo": "adicionar"})


def test_ajustar_tela():
    c = ctx_com()
    assert ferramentas.executar(c, "ajustar_tela", {"acao": "ordenar", "alvo": "cabos", "coluna": "ativo", "ordem": "desc"})["ok"]
    assert ferramentas.executar(c, "ajustar_tela", {"acao": "filtrar", "alvo": "outros", "coluna": "operacao", "valor": "I"})["ok"]
    assert ferramentas.executar(c, "ajustar_tela", {"acao": "limpar_filtros"})["ok"]
    assert c.acoes_ui[0] == {"acao_ui": "ordenar", "alvo": "cabos", "coluna": "ativo", "ordem": "desc"} and c.acoes_ui[2] == {"acao_ui": "limpar_filtros", "alvo": "ambos"}
    for ruim in ({"acao": "ordenar", "coluna": "senha"}, {"acao": "filtrar", "coluna": "ativo"}, {"acao": "x"}, {"acao": "limpar_filtros", "alvo": "z"}):
        assert "erro" in ferramentas.executar(c, "ajustar_tela", ruim)
    assert len(c.acoes_ui) == 3


def test_skills_pelas_ferramentas(pastas):
    c = ctx_com()
    assert {s["nome"] for s in ferramentas.executar(c, "listar_skills", {})["skills"]} == {"orcamento", "formato"}
    assert ferramentas.executar(c, "carregar_skill", {"nome": "orcamento"})["conteudo"].endswith("regra de PE")
    r = ferramentas.executar(c, "carregar_skill", {"nome": "nada"})
    assert "erro" in r and "orcamento" in r["disponiveis"]


def test_especificacoes_sao_json_schema_validos():
    specs = ferramentas.especificacoes()
    assert {s["name"] for s in specs} == set(ferramentas.FERRAMENTAS) and len(specs) == 13
    for s in specs:
        assert s["description"] and s["parameters"]["type"] == "object"
        assert set(s["parameters"].get("required", [])) <= set(s["parameters"]["properties"])
        assert s["parameters"].get("required") != []
        json.dumps(s)


# ── adaptadores (conversão para cada SDK) ───────────────────────────────────────────────────────────
MSGS = [{"role": "user", "texto": "oi"},
        {"role": "assistant", "texto": "vou ver", "chamadas": [{"id": "c1", "nome": "ler_tabela", "args": {"tabela": "cabos"}}], "bruto": None},
        {"role": "tool", "resultados": [{"id": "c1", "nome": "ler_tabela", "resultado": {"cabos": {"total": 0}}}]}]


def test_converter_openai():
    m = agente.AdaptadorOpenAI.converter("SYS", MSGS)
    assert m[0] == {"role": "system", "content": "SYS"} and m[1] == {"role": "user", "content": "oi"}
    tc = m[2]["tool_calls"][0]
    assert tc["id"] == "c1" and tc["function"]["name"] == "ler_tabela" and json.loads(tc["function"]["arguments"]) == {"tabela": "cabos"}
    assert m[3]["role"] == "tool" and m[3]["tool_call_id"] == "c1" and json.loads(m[3]["content"])["cabos"]["total"] == 0


def test_converter_claude():
    m = agente.AdaptadorClaude.converter(MSGS)
    assert m[1]["content"][0] == {"type": "text", "text": "vou ver"} and m[1]["content"][1] == {"type": "tool_use", "id": "c1", "name": "ler_tabela", "input": {"tabela": "cabos"}}
    assert m[2]["role"] == "user" and m[2]["content"][0]["type"] == "tool_result" and m[2]["content"][0]["tool_use_id"] == "c1"


def test_converter_gemini_usa_o_bruto_quando_existe():
    from google.genai import types
    bruto = types.Content(role="model", parts=[types.Part.from_function_call(name="ler_tabela", args={"tabela": "cabos"})])
    c = agente.AdaptadorGemini.converter([MSGS[0], dict(MSGS[1], bruto=bruto), MSGS[2]])
    assert c[1] is bruto and c[0].role == "user"
    assert c[2].parts[0].function_response.name == "ler_tabela" and c[2].parts[0].function_response.response["resultado"]["cabos"]["total"] == 0
    sem = agente.AdaptadorGemini.converter(MSGS)                                   # sem bruto reconstrói a chamada
    assert sem[1].role == "model" and sem[1].parts[1].function_call.name == "ler_tabela"


def correr(coro):
    return asyncio.new_event_loop().run_until_complete(coro)


def test_adaptador_openai_le_chamadas_e_texto():
    chamadas = []

    async def criar(**kw):
        chamadas.append(kw)
        tc = types.SimpleNamespace(id="t1", function=types.SimpleNamespace(name="buscar_ativo", arguments='{"texto": "cfu"}'))
        return types.SimpleNamespace(choices=[types.SimpleNamespace(message=types.SimpleNamespace(content="", tool_calls=[tc]))])
    cli = types.SimpleNamespace(chat=types.SimpleNamespace(completions=types.SimpleNamespace(create=criar)))
    t = correr(agente.AdaptadorOpenAI(cli, ["m1"]).turno("SYS", [MSGS[0]], ferramentas.especificacoes()))
    assert t.chamadas[0].nome == "buscar_ativo" and t.chamadas[0].args == {"texto": "cfu"} and t.chamadas[0].id == "t1"
    assert chamadas[0]["tools"][0]["type"] == "function" and chamadas[0]["model"] == "m1"


def test_adaptador_claude_le_blocos():
    async def criar(**kw):
        assert kw["tools"][0]["input_schema"]["type"] == "object" and kw["system"] == "SYS"
        return types.SimpleNamespace(content=[types.SimpleNamespace(type="text", text="ok"), types.SimpleNamespace(type="tool_use", id="u1", name="totalizar", input={"com_plano": True})])
    cli = types.SimpleNamespace(messages=types.SimpleNamespace(create=criar))
    t = correr(agente.AdaptadorClaude(cli, ["m1"]).turno("SYS", [MSGS[0]], ferramentas.especificacoes()))
    assert t.texto == "ok" and t.chamadas[0].nome == "totalizar" and t.chamadas[0].args == {"com_plano": True}


def test_adaptador_gemini_le_function_call_e_guarda_o_bruto():
    from google.genai import types as gt
    conteudo = gt.Content(role="model", parts=[gt.Part.from_function_call(name="ler_tabela", args={"tabela": "outros"}), gt.Part(text="pensando")])

    async def gerar(model, contents, config):
        assert config.tools[0].function_declarations[0].name in ferramentas.FERRAMENTAS and config.automatic_function_calling.disable
        return types.SimpleNamespace(candidates=[types.SimpleNamespace(content=conteudo)])
    cli = types.SimpleNamespace(aio=types.SimpleNamespace(models=types.SimpleNamespace(generate_content=gerar)))
    t = correr(agente.AdaptadorGemini(cli, ["m1"]).turno("SYS", [MSGS[0]], ferramentas.especificacoes()))
    assert t.chamadas[0].nome == "ler_tabela" and t.chamadas[0].args == {"tabela": "outros"} and t.texto == "pensando" and t.bruto is conteudo


def test_fallback_de_modelo_em_erro_de_cota_e_erro_real_sobe():
    usados = []

    async def chamar(m):
        usados.append(m)
        if m == "a":
            raise RuntimeError("429 quota exhausted")
        return m
    assert correr(agente.com_fallback(["a", "b"], chamar)) == "b" and usados == ["a", "b"]

    async def quebra(m):
        raise ValueError("chave inválida")
    with pytest.raises(ValueError):
        correr(agente.com_fallback(["a", "b"], quebra))
    with pytest.raises(RuntimeError, match="quota"):
        correr(agente.com_fallback(["a"], _sempre_cota))


async def _sempre_cota(m):
    raise RuntimeError("quota")


# ── o laço ──────────────────────────────────────────────────────────────────────────────────────────
class Roteiro(agente.AdaptadorBase):
    """Responde por roteiro: cada item é um Turno; guarda o que recebeu."""
    def __init__(self, turnos):
        self.turnos, self.recebido = list(turnos), []

    async def turno(self, system, mensagens, specs):
        self.recebido.append([dict(m) for m in mensagens])
        t = self.turnos.pop(0)
        if isinstance(t, Exception):
            raise t
        return t


def eventos(adaptador, ctx, prompt="oi", historico=()):
    async def tudo():
        return [e async for e in agente.rodar(adaptador, "SYS", list(historico), prompt, ctx)]
    return correr(tudo())


def test_laco_executa_ferramenta_devolve_resultado_e_emite_plano():
    ctx = ctx_com(outros=[{"id": 0, "operacao": "I", "ativo": "1-CFU"}])
    ad = Roteiro([
        agente.Turno("vou olhar", [agente.Chamada("a", "ler_tabela", {})]),
        agente.Turno("", [agente.Chamada("b", "propor_alteracoes", {"resumo": "mais um", "alteracoes": [{"acao": "editar", "tabela": "outros", "id": 0, "ativo": "2-CFU"}]}),
                          agente.Chamada("c", "ajustar_tela", {"acao": "limpar_filtros"})]),
        agente.Turno("Proposto: dobrar a chave."),
    ])
    ev = eventos(ad, ctx, "dobre a chave")
    assert [e["tipo"] for e in ev] == ["texto", "ferramenta", "ferramenta", "ferramenta", "texto", "acao_ui", "plano", "fim"]
    assert ev[0]["texto"] == "vou olhar" and ev[2]["nome"] == "propor_alteracoes" and ev[2]["detalhe"] == "1 alteração(ões)"
    plano = next(e for e in ev if e["tipo"] == "plano")
    assert plano["resumo"] == "mais um" and plano["itens"][0]["acao"] == "editar" and plano["itens"][0]["antes"]["ativo"] == "1-CFU"
    # o 2º turno viu o resultado da ferramenta (ler_tabela) e o 3º viu o do propor_alteracoes
    assert ad.recebido[1][-1]["role"] == "tool" and ad.recebido[1][-1]["resultados"][0]["resultado"]["outros"]["total"] == 1
    assert ad.recebido[2][-1]["resultados"][0]["resultado"]["aceitas"] == 1
    assert ad.recebido[0][-1] == {"role": "user", "texto": "dobre a chave"}


def test_laco_sem_ferramenta_responde_texto_e_nao_emite_plano():
    ev = eventos(Roteiro([agente.Turno("Olá!")]), ctx_com())
    assert ev == [{"tipo": "texto", "texto": "Olá!"}, {"tipo": "fim"}]


def test_laco_erro_do_provedor_vira_evento_e_termina():
    ev = eventos(Roteiro([RuntimeError("chave inválida")]), ctx_com())
    assert ev[0]["tipo"] == "erro" and "chave inválida" in ev[0]["mensagem"] and ev[-1] == {"tipo": "fim"}


def test_laco_corta_excesso_de_passos():
    ad = Roteiro([agente.Turno("", [agente.Chamada(str(i), "listar_skills", {})]) for i in range(agente.MAX_PASSOS + 2)])
    ev = eventos(ad, ctx_com())
    assert ev[-2]["tipo"] == "texto" and "excesso de passos" in ev[-2]["texto"] and len(ad.turnos) == 2


def test_historico_neutro_alterna_junta_e_comeca_em_user():
    h = [{"role": "model", "parts": [{"text": "sobra"}]}, {"role": "user", "parts": [{"text": "a"}]}, {"role": "user", "parts": [{"text": "b"}]},
         {"role": "model", "parts": [{"text": "*(Ação de UI executada)*"}]}, {"role": "model", "parts": [{"text": ""}]}]
    n = agente.historico_neutro(h)
    assert [m["role"] for m in n] == ["user", "assistant"] and n[0]["texto"] == "a\nb"


# ── a rota ──────────────────────────────────────────────────────────────────────────────────────────
@pytest.fixture
def client(tmp_path, monkeypatch, pastas):
    monkeypatch.setattr(database, "DB_PATH", str(tmp_path / "teste.db"))
    for v in ("GEMINI_API_KEY", "GOOGLE_API_KEY", "OPENAI_API_KEY", "ANTHROPIC_API_KEY"):
        monkeypatch.delenv(v, raising=False)
    database.init_db()
    conn = database.get_connection()
    conn.execute("INSERT OR IGNORE INTO projetos (nome, codigo) VALUES ('PERNAMBUCO', '081')")
    conn.commit()
    conn.close()
    ai_chat._rate_hits.clear()
    return TestClient(appmod.app)


def cab(role="admin"):
    return {"Authorization": f"Bearer {create_jwt_token('u1', f'{role}@x.com', role)}"}


def linhas(resp):
    return [json.loads(l) for l in resp.text.splitlines() if l.strip()]


def test_rota_sem_chave_da_401(client):
    assert client.post("/api/assistente/chat", headers=cab(), json={"prompt": "oi"}).status_code == 401


def test_rota_chat_streama_ndjson_com_plano(client, monkeypatch):
    usados = {}

    def criar(provider, api_key, modelos, base_url=""):
        usados.update(provider=provider, api_key=api_key, modelos=modelos)
        return Roteiro([agente.Turno("", [agente.Chamada("a", "ler_tabela", {"tabela": "cabos"})]), agente.Turno("Pronto.")])
    monkeypatch.setattr(agente, "criar_adaptador", criar)
    r = client.post("/api/assistente/chat", headers={**cab(), "X-Gemini-Key": "k1"},
                    json={"prompt": "o que tem na tela?", "projeto_codigo": "081", "cabos": [{"id": 0, "operacao": "I", "ativo": "CAA2 3 35 m"}], "outros": [], "history": []})
    assert r.status_code == 200 and r.headers["content-type"].startswith("application/x-ndjson")
    ev = linhas(r)
    assert [e["tipo"] for e in ev] == ["ferramenta", "texto", "fim"] and ev[0]["nome"] == "ler_tabela"
    assert usados["provider"] == "gemini" and usados["api_key"] == "k1" and usados["modelos"][0] == "gemini-3.1-flash-lite"


def test_rota_monta_o_prompt_com_skills_regras_e_estado(client, monkeypatch):
    visto = {}

    def criar(provider, api_key, modelos, base_url=""):
        class Espia(agente.AdaptadorBase):
            async def turno(self, system, mensagens, specs):
                visto["system"], visto["n_specs"] = system, len(specs)
                return agente.Turno("ok")
        return Espia()
    monkeypatch.setattr(agente, "criar_adaptador", criar)
    client.post("/api/regras", headers=cab(), json={"conteudo": "regra: sempre TERRA3", "projeto_codigo": "081"})
    client.post("/api/assistente/chat", headers={**cab(), "X-Anthropic-Key": "k"}, json={"prompt": "oi", "provider": "claude", "projeto_codigo": "081", "cabos": [{"id": 0, "ativo": "x"}]})
    s = visto["system"]
    assert "OrçaCore" in s and "- orcamento: orçar em PE" in s and "sempre TERRA3" in s and "PERNAMBUCO" in s and "1 linha(s) em Cabos" in s and visto["n_specs"] == 13


def test_rota_atalho_rapido_nao_chama_ia(client, monkeypatch):
    monkeypatch.setattr(agente, "criar_adaptador", lambda *a, **k: (_ for _ in ()).throw(AssertionError("não devia chamar a IA")))
    ev = linhas(client.post("/api/assistente/chat", headers={**cab(), "X-Gemini-Key": "k"}, json={"prompt": "adicionar 2 con", "projeto_codigo": "081"}))
    assert ev[0]["tipo"] == "plano" and ev[0]["auto"] is True and ev[0]["itens"][0]["ativo"] == "2-CON" and ev[0]["itens"][0]["operacao"] == "I"
    ev = linhas(client.post("/api/assistente/chat", headers={**cab(), "X-Gemini-Key": "k"}, json={"prompt": "remover 1,5 IT", "projeto_codigo": "081"}))
    assert ev[0]["itens"][0]["ativo"] == "1.5-IT" and ev[0]["itens"][0]["operacao"] == "R"


def test_rota_skills_listar_ler_e_permissoes(client):
    j = client.get("/api/assistente/skills?projeto=081", headers=cab("operador")).json()
    assert j["projeto"] == "PERNAMBUCO" and j["escopo_projeto"] == "pernambuco" and j["pode_editar"] is False
    assert {s["nome"] for s in j["skills"]} == {"orcamento", "formato"}
    assert client.get("/api/assistente/skills/_geral/orcamento", headers=cab("operador")).json()["texto"].startswith("descricao: como orçar")
    assert client.get("/api/assistente/skills/_geral/naoha", headers=cab("operador")).status_code == 404
    assert client.put("/api/assistente/skills/_geral/novo", headers=cab("operador"), json={"texto": "x"}).status_code == 403
    assert client.delete("/api/assistente/skills/_geral/orcamento", headers=cab("operador")).status_code == 403


def test_rota_skills_admin_grava_apaga_e_recusa_nomes_ruins(client):
    assert client.put("/api/assistente/skills/pernambuco/regras_minhas", headers=cab(), json={"texto": "descricao: minhas\nsempre X"}).status_code == 200
    j = client.get("/api/assistente/skills?projeto=081", headers=cab()).json()
    assert j["pode_editar"] is True and any(s["nome"] == "regras_minhas" and s["origem"] == "usuario" for s in j["skills"])
    assert client.put("/api/assistente/skills/_geral/Nome Ruim", headers=cab(), json={"texto": "x"}).status_code == 400
    assert client.put("/api/assistente/skills/_geral/vazia", headers=cab(), json={"texto": " "}).status_code == 400
    assert client.delete("/api/assistente/skills/pernambuco/regras_minhas", headers=cab()).status_code == 200
    assert client.delete("/api/assistente/skills/pernambuco/regras_minhas", headers=cab()).status_code == 404


def test_rota_gerar_catalogo_grava_a_skill_ativos(client, monkeypatch):
    monkeypatch.setattr("services.sync_service.get_merged_orcamento", lambda uid=None: BASE)
    r = client.post("/api/assistente/skills/gerar-catalogo?projeto=081", headers=cab())
    assert r.status_code == 200 and r.json()["ativos"] == 5 and r.json()["escopo"] == "pernambuco"
    assert "`TR125`" in skills.carregar("pernambuco", "ativos")
    assert client.post("/api/assistente/skills/gerar-catalogo?projeto=081", headers=cab("operador")).status_code == 403
    assert client.post("/api/assistente/skills/gerar-catalogo?projeto=999", headers=cab()).status_code == 404
    monkeypatch.setattr("services.sync_service.get_merged_orcamento", lambda uid=None: [])
    assert client.post("/api/assistente/skills/gerar-catalogo?projeto=081", headers=cab()).status_code == 400


def test_chat_antigo_continua_resolvendo_provedor_igual(client):
    """A extração de resolver_provedor não muda o /api/gemini/chat: sem chave, 401 como antes."""
    assert client.post("/api/gemini/chat", headers=cab(), json={"prompt": "oi", "history": []}).status_code == 401
