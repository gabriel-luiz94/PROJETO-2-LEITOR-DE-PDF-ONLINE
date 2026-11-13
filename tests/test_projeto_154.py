"""TASK-154 — projeto de trabalho (.obra numa pasta): arquivo, ponte da janela desktop, módulo JS e integração nas telas."""
import base64
import json
import os
import shutil
import subprocess
import sys
import types

import pytest

import config
from services import obras_arquivo, projeto_desktop as pd, projeto_obra as po

NODE = shutil.which("node")
DADOS = {"cabos": {"data": [{"entidade": "1", "operacao": "I", "ativo": "C35", "qtdAtivos": 16, "texto": "3xC35 16 m"}]},
         "outros": {"data": [{"entidade": "2", "operacao": "I", "ativo": "11300D"}]}, "totalizadora": {"data": [{"id": "t1", "operacao": "I", "ativo": "RETENSCN(C35)", "qtd": 16}]}}
SAP = {"sap": "4500", "bairro": "Centro", "municipio": "Recife", "utd": "UTD1", "campo_estranho": "x"}
RESULTADO = {"assinatura": "abc-1", "linhas": [{"operacao": "I", "mdo": "MATERIAL", "codigo": "100", "desc_codigo": "Poste", "total": 2.5, "filtro": "", "lixo": 1}]}


def parcial(**kw):
    base = {"nome": "Obra Pirapama 2026", "numero_obra": "4500123", "sap": SAP, "dados": DADOS, "resultado": RESULTADO}
    base.update(kw)
    return base


# ───────── arquivo (funções puras) ─────────
def test_montar_e_ler_ida_e_volta():
    texto = po.montar(parcial(), None, "081")
    carga = json.loads(texto)
    assert carga["formato"] == "leitor-obra" and carga["versao"] == 1 and carga["tipo"] == "obra" and carga["projeto"] == "081"
    assert carga["projeto_trabalho"]["nome"] == "Obra Pirapama 2026" and carga["projeto_trabalho"]["numero_obra"] == "4500123"
    r = po.ler(texto, "081")
    assert r["nome"] == "Obra Pirapama 2026" and r["numero_obra"] == "4500123" and r["legado"] is False and r["projeto"] == "081"
    assert r["sap"] == {"sap": "4500", "bairro": "Centro", "municipio": "Recife", "utd": "UTD1"}                    # campo fora da lista é descartado
    assert r["dados"]["cabos"][0]["ativo"] == "C35" and r["dados"]["totalizadora"][0]["ativo"] == "RETENSCN(C35)"
    assert r["resultado"] == {"assinatura": "abc-1", "linhas": [{"operacao": "I", "mdo": "MATERIAL", "codigo": "100", "desc_codigo": "Poste", "filtro": "", "total": 2.5}]}


def test_gravacao_parcial_mantem_o_resto():
    antigo = json.loads(po.montar(parcial(), None, "081"))
    novo = json.loads(po.montar({"numero_obra": "999"}, antigo))
    assert novo["projeto_trabalho"]["numero_obra"] == "999" and novo["projeto_trabalho"]["nome"] == "Obra Pirapama 2026"
    assert novo["projeto_trabalho"]["sap"]["bairro"] == "Centro" and novo["projeto_trabalho"]["resultado"]["assinatura"] == "abc-1"
    assert novo["dados"]["cabos"][0]["ativo"] == "C35" and novo["projeto"] == "081"
    sem_dados = json.loads(po.montar({"dados": {"cabos": [], "outros": []}}, antigo))
    assert sem_dados["dados"]["cabos"] == [] and sem_dados["projeto_trabalho"]["resultado"]["assinatura"] == "abc-1"


def test_nome_obrigatorio_e_limites():
    with pytest.raises(po.ErroProjeto):
        po.montar({"nome": "  "}, None, "081")
    assert len(json.loads(po.montar({"nome": "x" * 500}, None, "081"))["projeto_trabalho"]["nome"]) == po.LIMITE_NOME
    grande = {"assinatura": "a", "linhas": [{"codigo": str(i)} for i in range(po.LIMITE_LINHAS_RESULTADO + 50)]}
    assert len(po.resultado_limpo(grande)["linhas"]) == po.LIMITE_LINHAS_RESULTADO
    assert po.resultado_limpo("x") is None and po.resultado_limpo({"linhas": "x"}) is None


def test_nome_de_arquivo_e_assinatura():
    assert po.nome_arquivo("Obra: A/B?") == "Obra_ A_B.obra" or po.nome_arquivo("Obra: A/B?").endswith(".obra")
    assert po.nome_arquivo("...") == "projeto.obra" and "/" not in po.nome_arquivo("a/b") and "\\" not in po.nome_arquivo("a\\b")
    assert po.assinatura("abc") == po.assinatura(b"abc") and po.assinatura("abc") != po.assinatura("abd")


def test_abre_obra_json_antigo_sem_bloco_de_projeto():
    antigo = json.dumps(obras_arquivo.montar_exportacao({"nome": "Obra velha", "projeto": "081", "tipo": "obra", "dados_json": json.dumps(DADOS)}))
    r = po.ler(antigo, "081")
    assert r["legado"] is True and r["nome"] == "Obra velha" and r["numero_obra"] == "" and r["sap"] == {} and r["resultado"] is None
    assert r["dados"]["cabos"][0]["ativo"] == "C35"


def test_ler_recusa_o_que_nao_presta():
    texto = po.montar(parcial(), None, "081")
    for ruim, trecho in ((texto, "projeto"), ("{", "JSON"), ("[]", "esperado um objeto"), (b"\xff\xfe\x00", "ilegível")):
        with pytest.raises(po.ErroProjeto) as e:
            po.ler(ruim, "999" if ruim == texto else "081")
        assert trecho in " ".join(e.value.mensagens)
    modelo = json.dumps(obras_arquivo.montar_exportacao({"nome": "M", "projeto": "081", "tipo": "modelo",
                                                         "dados_json": json.dumps({"cabos": [{"operacao": "I", "ativo": "V", "qtdAtivos": 1}], "outros": []})}))
    with pytest.raises(po.ErroProjeto):
        po.ler(modelo, "081")


def test_programa_antigo_abre_o_obra_novo_como_obra_comum():
    carga = json.loads(po.montar(parcial(), None, "081"))
    pronta = obras_arquivo.validar_importacao(carga, "081")                    # o bloco extra é ignorado
    assert pronta["nome"] == "Obra Pirapama 2026" and pronta["dados"]["cabos"][0]["ativo"] == "C35"


# ───────── pasta e arquivo (criar / gravar / abrir) ─────────
def test_criar_gravar_abrir(tmp_path):
    pasta = tmp_path / "Obras" / "Pirapama"
    r = pd.criar(str(pasta), parcial(), "081")
    assert r["ok"] and os.path.basename(r["caminho"]) == "Obra Pirapama 2026.obra" and os.path.isfile(r["caminho"]) and not os.path.exists(r["caminho"] + ".tmp")
    assert pd.criar(str(pasta), parcial(), "081")["existe"] is True                          # não sobrescreve um projeto existente
    g = pd.gravar(r["caminho"], {"numero_obra": "777"}, r["assinatura"], False, "081")
    assert g["ok"] and g["assinatura"] != r["assinatura"]
    a = pd.abrir(r["caminho"], "081")
    assert a["ok"] and a["projeto"]["numero_obra"] == "777" and a["projeto"]["dados"]["outros"][0]["ativo"] == "11300D" and a["assinatura"] == g["assinatura"]
    assert pd.assinatura_atual(r["caminho"]) == {"ok": True, "existe": True, "assinatura": g["assinatura"]}


def test_alteracao_por_fora_pede_decisao(tmp_path):
    r = pd.criar(str(tmp_path), parcial(), "081")
    conteudo = json.loads(open(r["caminho"], encoding="utf-8").read()); conteudo["projeto_trabalho"]["numero_obra"] = "FORA"
    with open(r["caminho"], "w", encoding="utf-8") as f:
        json.dump(conteudo, f)                                                                  # outro computador mexeu
    g = pd.gravar(r["caminho"], {"numero_obra": "MEU"}, r["assinatura"], False, "081")
    assert g["ok"] is False and g["conflito"] is True and "outro" in g["erro"]
    assert pd.abrir(r["caminho"], "081")["projeto"]["numero_obra"] == "FORA"                  # nada foi gravado
    forca = pd.gravar(r["caminho"], {"numero_obra": "MEU"}, r["assinatura"], True, "081")
    assert forca["ok"] and pd.abrir(r["caminho"], "081")["projeto"]["numero_obra"] == "MEU"


def test_so_grava_arquivos_obra_e_nao_levanta_excecao(tmp_path):
    assert pd.gravar(str(tmp_path / "x.txt"), parcial(), "", False)["ok"] is False
    assert pd.gravar(str(tmp_path / "naoexiste" / "x.obra"), parcial(), "", False, "081")["ok"] is False
    assert pd.criar("", parcial())["ok"] is False and pd.criar(str(tmp_path), {"nome": ""})["ok"] is False
    assert pd.abrir(str(tmp_path / "nao.obra"), "081")["ok"] is False
    ruim = tmp_path / "ruim.obra"; ruim.write_text("não é json", encoding="utf-8")
    assert pd.abrir(str(ruim), "081")["ok"] is False
    assert pd.assinatura_atual(str(tmp_path / "nao.obra")) == {"ok": True, "existe": False}


def test_gravacao_sem_assinatura_esperada_nao_acusa_conflito(tmp_path):
    r = pd.criar(str(tmp_path), parcial(), "081")
    assert pd.gravar(r["caminho"], {"numero_obra": "1"}, "", False, "081")["ok"]


# ───────── ponte da janela ─────────
class Janela:
    def __init__(self, saida):
        self.saida, self.chamadas = saida, []

    def create_file_dialog(self, tipo, **kw):
        self.chamadas.append((tipo, kw))
        return self.saida


@pytest.fixture
def webview_falso(monkeypatch):
    m = types.ModuleType("webview")
    m.SAVE_DIALOG, m.OPEN_DIALOG, m.FOLDER_DIALOG = 30, 10, 20
    monkeypatch.setitem(sys.modules, "webview", m)
    return m


def test_api_projeto_pasta_e_arquivo(tmp_path, webview_falso):
    api = pd.ApiProjeto(); api.vincular(Janela((str(tmp_path),)))
    assert api.projeto_escolher_pasta(str(tmp_path)) == {"ok": True, "pasta": str(tmp_path)} and api._janela.chamadas[0][0] == 20
    c = api.projeto_criar(str(tmp_path), json.dumps(parcial()), "081")
    assert c["ok"]
    api2 = pd.ApiProjeto(); api2.vincular(Janela((c["caminho"],)))
    a = api2.projeto_escolher_arquivo("081")
    assert a["ok"] and a["projeto"]["nome"] == "Obra Pirapama 2026" and api2._janela.chamadas[0][0] == 10
    api3 = pd.ApiProjeto(); api3.vincular(Janela(None))
    assert api3.projeto_escolher_arquivo("081") == {"cancelado": True} and api3.projeto_escolher_pasta() == {"cancelado": True}
    assert pd.ApiProjeto().projeto_escolher_pasta()["ok"] is False                                # sem janela
    assert api.projeto_criar("x", "{") ["ok"] is False and api.projeto_gravar(c["caminho"], "[]")["ok"] is False


def test_salvar_arquivo_abre_na_pasta_do_projeto(tmp_path, webview_falso):
    alvo = tmp_path / "REL.pdf"
    api = pd.ApiProjeto(); api.vincular(Janela(str(alvo)))
    assert api.salvar_arquivo("REL.pdf", base64.b64encode(b"x").decode(), "", str(tmp_path))["ok"]
    assert api._janela.chamadas[0][1]["directory"] == str(tmp_path)
    api.salvar_arquivo("REL.pdf", base64.b64encode(b"x").decode(), "", str(tmp_path / "nao_existe"))
    assert api._janela.chamadas[1][1]["directory"] == ""                                        # pasta que sumiu: abre onde o sistema quiser


def test_api_projeto_so_expoe_metodos_publicos_esperados():
    publicos = sorted(n for n in dir(pd.ApiProjeto()) if not n.startswith("_"))
    assert publicos == ["abrir_arquivos", "projeto_abrir", "projeto_assinatura", "projeto_criar", "projeto_escolher_arquivo", "projeto_escolher_pasta",
                        "projeto_gravar", "projeto_rascunho_gravar", "projeto_rascunho_info", "projeto_rascunho_ler", "salvar_arquivo", "vincular"]


def test_app_usa_a_ponte_do_projeto():
    codigo = open(os.path.join(config.BASE_DIR, "app.py"), encoding="utf-8").read()
    assert "ApiProjeto()" in codigo and '"/static/projeto.js"' in codigo


# ───────── módulo JS e telas ─────────
@pytest.mark.skipif(not NODE, reason="node não instalado")
def test_funcoes_puras_do_projeto_js():
    mod = os.path.join(config.STATIC_DIR, "projeto.js")
    codigo = f"""const P=require({json.dumps(mod)});
const r=[P.hashTexto('abc')===P.hashTexto('abc'), P.hashTexto('abc')!==P.hashTexto('abd'), P.hashTexto('')==='811c9dc5-0',
 P.pastaDoCaminho('D:\\\\Obras\\\\A\\\\x.obra')==='D:\\\\Obras\\\\A', P.pastaDoCaminho('/home/u/x.obra')==='/home/u', P.pastaDoCaminho('x.obra')==='',
 P.temDados({{cabos:{{data:[{{ativo:'C35'}}]}}}}), P.temDados({{outros:[{{texto:'abc'}}]}}), !P.temDados({{cabos:{{data:[{{ativo:' '}},null]}},outros:[],totalizadora:{{data:[]}}}}), !P.temDados(null)];
console.log(JSON.stringify(r));"""
    saida = subprocess.run([NODE, "-e", codigo], capture_output=True, text=True, timeout=30)
    assert json.loads(saida.stdout.strip()) == [True] * 10, saida.stderr


def ler(nome):
    return open(os.path.join(config.STATIC_DIR, nome), encoding="utf-8").read()


def test_telas_ligadas_ao_projeto():
    resumo, resultado, js = ler("resumo.html"), ler("resultado_orcamento.html"), ler("resumo.js")
    assert "/static/projeto.js" in resumo and 'id="projeto-menu-host"' in resumo and "GridRProjeto.numeroObra()" in resumo
    assert "/static/projeto.js" in resultado and "confirmarNumeroObra" in resultado and "salvarSap" in resultado and "acompanharResultado" in resultado and "carregarResultado" in resultado
    assert "registrarResumo" in js and "GridRProjeto.marcarSujo(true)" in js
    assert "pastaDoProjeto()" in ler("desktop_arquivos.js")


def test_projeto_so_aparece_no_aplicativo_desktop():
    js = ler("projeto.js")
    assert "if (!host || host.dataset.pronto || !ponte()) return;" in js                    # sem a ponte (navegador/online) o menu nem é criado
    assert "if (!ativo()) return;" in js and "pywebview.api.projeto_criar" in js


# ───────── Colar SAP (JSON do Atlas) ─────────
ATLAS = {"source": "atlas-projeto-rede", "version": 1, "titulo": "OLCF -INST SECCION  U B088237 RECIFE", "nota": "9101578720", "sap": "E0483118",
         "bairro": "BREJO DA GUABIRABA", "municipio": "RECIFE", "utd": "CASA FORTE", "utep": "METROPOLITANO", "projetista": "GABRIEL SALES",
         "data": "11/11/2026", "odi": "8.2 Melhoramento", "odiCodigo": "8.2", "odd": "", "oddCodigo": ""}


def interpretar(texto):
    mod = os.path.join(config.STATIC_DIR, "projeto.js")
    codigo = f"const P=require({json.dumps(mod)});console.log(JSON.stringify(P.interpretarSap({json.dumps(texto)})));"
    saida = subprocess.run([NODE, "-e", codigo], capture_output=True, text=True, timeout=30)
    assert saida.returncode == 0, saida.stderr
    return json.loads(saida.stdout)


@pytest.mark.skipif(not NODE, reason="node não instalado")
def test_colar_sap_mapeia_os_campos_do_atlas():
    r = interpretar(json.dumps(ATLAS))
    assert r["ok"] and r["campos"] == {"titulo": "OLCF -INST SECCION  U B088237 RECIFE", "nota": "9101578720", "sap": "E0483118", "bairro": "BREJO DA GUABIRABA",
                                       "municipio": "RECIFE", "utd": "CASA FORTE", "utep": "METROPOLITANO", "projetista": "GABRIEL SALES", "data": "11/11/2026",
                                       "id_investimento": "8.2"}            # odiCodigo vira o ID de investimento; odd/oddCodigo vazios não aparecem (não apagam nada)
    assert "id_investimento_odd" not in r["campos"]


@pytest.mark.skipif(not NODE, reason="node não instalado")
def test_colar_sap_casos_de_borda():
    com_odd = interpretar(json.dumps({**ATLAS, "oddCodigo": "9.1"}))["campos"]
    assert com_odd["id_investimento_odd"] == "9.1"
    sem_codigo = interpretar(json.dumps({**ATLAS, "odiCodigo": "", "odi": "8.2 Melhoramento"}))["campos"]
    assert sem_codigo["id_investimento"] == "8.2"                                               # tira o número do começo da descrição
    assert interpretar(json.dumps({**ATLAS, "data": "2026-11-11"}))["campos"]["data"] == "11/11/2026"
    cercado = interpretar("```json\n" + json.dumps(ATLAS) + "\n```")
    assert cercado["ok"] and cercado["campos"]["nota"] == "9101578720"
    assert interpretar("\ufeff " + json.dumps(ATLAS))["ok"]
    longo = interpretar(json.dumps({**ATLAS, "titulo": "x" * 500, "bairro": "a\nb\u0000c"}))["campos"]
    assert len(longo["titulo"]) == 120 and "\n" not in longo["bairro"] and "\x00" not in longo["bairro"]


@pytest.mark.skipif(not NODE, reason="node não instalado")
def test_colar_sap_recusa_o_que_nao_e_do_atlas():
    for texto, trecho in (("", "nada copiado"), ("abc", "não é um JSON"), ("[1]", "não é um JSON"), (json.dumps({"titulo": "x"}), "não é do Atlas"),
                          (json.dumps({**ATLAS, "version": 2}), "não é suportada"), (json.dumps({"source": "atlas-projeto-rede", "version": 1}), "nenhum dado"),
                          (json.dumps({**ATLAS, "titulo": 5, "nota": {"a": 1}, "sap": None, "bairro": "", "municipio": "", "utd": "", "utep": "", "projetista": "", "data": "",
                                       "odiCodigo": "", "odi": ""}), "")):
        r = interpretar(texto)
        if trecho:
            assert r["ok"] is False and trecho in r["erro"], (texto, r)
    assert interpretar(json.dumps({**ATLAS, "titulo": 5, "nota": {"a": 1}}))["campos"]["titulo"] == "5"      # número vira texto; objeto é ignorado
    assert "nota" not in interpretar(json.dumps({**ATLAS, "nota": {"a": 1}}))["campos"]


def test_botoes_colar_sap_nas_telas():
    js, resultado = ler("projeto.js"), ler("resultado_orcamento.html")
    assert js.count("📋 Colar SAP") == 2 and "colarSap" in js                                   # janela do projeto e janela das informações do SAP
    assert 'id="rel-colar-sap"' in resultado and "colarSapNoRelatorio" in resultado and "é diferente do número do projeto" in resultado
