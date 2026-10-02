"""TASK-031 fase C — pasta monitorada: estabilidade, fila (um por vez), erros/, processados/, duplicados e a thread da vigia."""
import json
import os
import threading
import time

import ezdxf
import pytest

import database
from services.autonomo import config_autonomo as ca
from services.autonomo import execucoes, pasta, pipeline as pl
from services.autonomo.leitor_js import LeitorJS
from tests.test_autonomo_pipeline import RECEITA_REMOVE, contexto

pytestmark = pytest.mark.filterwarnings("ignore")


@pytest.fixture(scope="module")
def leitor():
    lt = LeitorJS()
    yield lt
    lt.fechar()


@pytest.fixture
def amb(tmp_path, monkeypatch, leitor):
    monkeypatch.setattr(database, "DB_PATH", str(tmp_path / "t.db"))
    database.init_db()
    conn = database.get_connection()
    conn.execute("INSERT OR IGNORE INTO projetos (nome, codigo) VALUES ('P1', 'P1')")
    conn.commit()
    conn.close()
    receitas = []
    monkeypatch.setattr(pl, "carregar_contexto", lambda proj, user: contexto(leitor, receitas))
    amb_ = type("Amb", (), {})()
    amb_.receitas = receitas
    amb_.base = tmp_path / "auto"
    amb_.cfg = ca.salvar({"user_id": "u1", "pasta_base": str(amb_.base), "intervalo_s": 1, "estabilizacao_s": 0})
    amb_.p = ca.garantir_pastas(amb_.cfg)
    amb_.vigia = pasta.Vigia()
    return amb_


def dxf(caminho, textos):
    os.makedirs(os.path.dirname(str(caminho)), exist_ok=True)
    doc = ezdxf.new()
    msp = doc.modelspace()
    for i, t in enumerate(textos):
        msp.add_text(t, dxfattribs={"color": 1, "insert": (i * 10, 0), "height": 2})
    doc.saveas(str(caminho))
    return str(caminho)


def varrer2(a):
    """1ª varredura só observa; a 2ª (arquivo estável) processa."""
    assert a.vigia.varrer(a.cfg) == []
    return a.vigia.varrer(a.cfg)


def test_arquivo_estavel_e_processado_e_movido(amb):
    dxf(os.path.join(amb.p["entrada"], "P1", "obra1.dxf"), ["1-U4", "2-CFU"])
    r = varrer2(amb)
    assert [x["status"] for x in r] == ["ok"] and r[0]["projeto"] == "P1"
    assert not os.path.exists(os.path.join(amb.p["entrada"], "P1", "obra1.dxf"))
    assert os.listdir(os.path.join(amb.p["processados"], "P1")) == ["obra1.dxf"]
    ex = execucoes.buscar(r[0]["execucao"])
    assert ex["arquivo_caminho"] == os.path.join(amb.p["processados"], "P1", "obra1.dxf") and ex["obra_id"]
    assert os.path.isdir(ex["pasta_saida"]) and amb.p["saida"] in ex["pasta_saida"]
    st = amb.vigia.status()
    assert st["processados"] == 1 and st["na_fila"] == 1 and st["processando"] is None


def test_arquivo_ainda_sendo_copiado_espera_estabilizar(amb):
    alvo = os.path.join(amb.p["entrada"], "P1", "grande.dxf")
    dxf(alvo, ["1-U4"])
    assert amb.vigia.varrer(amb.cfg) == []
    with open(alvo, "ab") as f:                      # cresceu entre as varreduras => ainda copiando
        f.write(b"\n")
    assert amb.vigia.varrer(amb.cfg) == []           # reinicia a contagem
    assert [x["status"] for x in amb.vigia.varrer(amb.cfg)] in (["ok"], ["erro"])   # estável: agora é tratado


def test_tempo_de_estabilizacao_usa_o_relogio(amb):
    t = [1000.0]
    vigia = pasta.Vigia(relogio=lambda: t[0])
    cfg = {**amb.cfg, "estabilizacao_s": 10}
    dxf(os.path.join(amb.p["entrada"], "P1", "a.dxf"), ["1-U4"])
    assert vigia.varrer(cfg) == []
    t[0] += 5
    assert vigia.varrer(cfg) == []                   # só 5 s
    t[0] += 6
    assert [x["status"] for x in vigia.varrer(cfg)] == ["ok"]


def test_arquivos_temporarios_sao_ignorados(amb):
    os.makedirs(os.path.join(amb.p["entrada"], "P1"))
    for nome in ("~$obra.dxf", ".oculto.dxf", "obra.dxf.part", "x.tmp", "y.crdownload"):
        open(os.path.join(amb.p["entrada"], "P1", nome), "wb").close()
    assert varrer2(amb) == []
    assert len(os.listdir(os.path.join(amb.p["entrada"], "P1"))) == 5     # nada foi tocado


def test_arquivo_solto_na_raiz_vai_para_erros_com_motivo(amb):
    dxf(os.path.join(amb.p["entrada"], "solto.dxf"), ["1-U4"])
    r = varrer2(amb)
    assert r[0]["status"] == "erro" and r[0]["mensagem"].startswith("Arquivo solto")
    pasta_erro = os.path.join(amb.p["erros"], "sem_projeto")
    assert sorted(os.listdir(pasta_erro)) == ["solto.dxf", "solto.dxf.erro.txt"]
    assert "subpasta" in open(os.path.join(pasta_erro, "solto.dxf.erro.txt"), encoding="utf-8").read()
    assert amb.vigia.status()["erros"] == 1 and "solto.dxf" in amb.vigia.status()["ultimo_erro"]


def test_projeto_inexistente_e_extensao_nao_suportada_vao_para_erros(amb):
    dxf(os.path.join(amb.p["entrada"], "NAOEXISTE", "a.dxf"), ["1-U4"])
    os.makedirs(os.path.join(amb.p["entrada"], "P1"), exist_ok=True)
    with open(os.path.join(amb.p["entrada"], "P1", "nota.txt"), "w") as f:
        f.write("x")
    r = varrer2(amb)
    msgs = {x["arquivo"]: x["mensagem"] for x in r}
    assert "não está cadastrado" in msgs["a.dxf"] and "não suportado" in msgs["nota.txt"]
    assert os.path.exists(os.path.join(amb.p["erros"], "NAOEXISTE", "a.dxf"))
    hist = {x["arquivo"]: x for x in execucoes.listar("erro")}          # as falhas antes do pipeline também aparecem no histórico
    assert set(hist) == {"a.dxf", "nota.txt"} and "não está cadastrado" in hist["a.dxf"]["mensagem"]
    assert hist["a.dxf"]["arquivo_caminho"] == os.path.join(amb.p["erros"], "NAOEXISTE", "a.dxf") and hist["a.dxf"]["projeto_codigo"] == "NAOEXISTE"
    assert os.path.exists(os.path.join(amb.p["erros"], "P1", "nota.txt"))


def test_arquivo_ruim_nao_derruba_a_fila(amb):
    ruim = os.path.join(amb.p["entrada"], "P1", "a_quebrado.dxf")
    os.makedirs(os.path.dirname(ruim), exist_ok=True)
    open(ruim, "w").write("não é DXF")
    dxf(os.path.join(amb.p["entrada"], "P1", "b_bom.dxf"), ["1-U4"])
    r = varrer2(amb)
    assert [(x["arquivo"], x["status"]) for x in r] == [("a_quebrado.dxf", "erro"), ("b_bom.dxf", "ok")]
    assert os.path.exists(os.path.join(amb.p["erros"], "P1", "a_quebrado.dxf"))          # original preservado
    ex = execucoes.buscar(r[0]["execucao"])
    assert ex["status"] == "erro" and ex["arquivo_caminho"].startswith(amb.p["erros"])


def test_erro_inesperado_no_pipeline_nao_derruba_a_fila(amb, monkeypatch):
    original = pl.processar_arquivo

    def quebra_o_primeiro(caminho, *a, **k):
        if os.path.basename(caminho) == "a.dxf":
            raise RuntimeError("boom")
        return original(caminho, *a, **k)
    monkeypatch.setattr(pl, "processar_arquivo", quebra_o_primeiro)
    dxf(os.path.join(amb.p["entrada"], "P1", "a.dxf"), ["1-U4"])
    dxf(os.path.join(amb.p["entrada"], "P1", "b.dxf"), ["1-U4"])
    r = varrer2(amb)
    assert [(x["arquivo"], x["status"]) for x in r] == [("a.dxf", "erro"), ("b.dxf", "ok")]
    assert os.path.exists(os.path.join(amb.p["erros"], "P1", "a.dxf"))


def test_mesmo_conteudo_vai_para_duplicados(amb):
    a = dxf(os.path.join(amb.p["entrada"], "P1", "x.dxf"), ["1-U4"])
    assert varrer2(amb)[0]["status"] == "ok"
    import shutil
    shutil.copy(os.path.join(amb.p["processados"], "P1", "x.dxf"), os.path.join(amb.p["entrada"], "P1", "x_copia.dxf"))
    r = varrer2(amb)
    assert r[0]["status"] == "duplicado"
    assert os.path.exists(os.path.join(amb.p["processados"], "P1", "duplicados", "x_copia.dxf"))


def test_exclusao_pendente_move_o_arquivo_e_fica_aguardando(amb):
    amb.receitas.append(RECEITA_REMOVE)
    dxf(os.path.join(amb.p["entrada"], "P1", "p.dxf"), ["1-U3 1-U4"])
    r = varrer2(amb)
    assert r[0]["status"] == "aguardando_confirmacao"
    assert os.path.exists(os.path.join(amb.p["processados"], "P1", "p.dxf"))
    assert execucoes.buscar(r[0]["execucao"])["obra_id"] is None


def test_sem_usuario_dono_o_arquivo_fica_na_pasta(amb):
    cfg = {**amb.cfg, "user_id": ""}
    dxf(os.path.join(amb.p["entrada"], "P1", "p.dxf"), ["1-U4"])
    amb.vigia.varrer(cfg)
    r = amb.vigia.varrer(cfg)
    assert r[0]["status"] == "parado" and "usuário dono" in r[0]["mensagem"]
    assert os.path.exists(os.path.join(amb.p["entrada"], "P1", "p.dxf"))      # não foi movido nem perdido


def test_nome_repetido_no_destino_ganha_sufixo(amb):
    for _ in range(2):
        dxf(os.path.join(amb.p["entrada"], "P1", "igual.dxf"), ["1-U4"] if _ == 0 else ["1-U3 1-U4"])
        varrer2(amb)
    nomes = os.listdir(os.path.join(amb.p["processados"], "P1"))
    assert len(nomes) == 2 and "igual.dxf" in nomes


def test_um_arquivo_por_vez(amb, monkeypatch):
    original = pl.processar_arquivo
    ativos, maximo = [0], [0]

    def espiao(*a, **k):
        ativos[0] += 1
        maximo[0] = max(maximo[0], ativos[0])
        time.sleep(0.05)
        try:
            return original(*a, **k)
        finally:
            ativos[0] -= 1
    monkeypatch.setattr(pl, "processar_arquivo", espiao)
    for i in range(3):
        dxf(os.path.join(amb.p["entrada"], "P1", f"o{i}.dxf"), [f"{i + 1}-U4"])
    amb.vigia.varrer(amb.cfg)
    threads = [threading.Thread(target=amb.vigia.varrer, args=(amb.cfg,)) for _ in range(2)]
    for t in threads:
        t.start()
    for t in threads:
        t.join()
    assert maximo[0] == 1                       # nunca dois ao mesmo tempo (TRAVA)


def test_reprocessar_cria_nova_execucao_a_partir_do_original(amb):
    dxf(os.path.join(amb.p["entrada"], "P1", "r.dxf"), ["1-U4"])
    ex = varrer2(amb)[0]["execucao"]
    novo = pasta.reprocessar(ex)
    assert novo["id"] != ex and novo["status"] in ("ok", "com_pendencias") and not novo.get("duplicado")
    assert os.listdir(os.path.join(amb.p["processados"], "P1")) == ["r.dxf"]         # o original continua onde está
    with pytest.raises(pl.ErroPipeline):
        pasta.reprocessar("naoexiste")
    os.remove(os.path.join(amb.p["processados"], "P1", "r.dxf"))
    with pytest.raises(pl.ErroPipeline, match="não está mais disponível"):
        pasta.reprocessar(ex)


def test_thread_da_vigia_processa_sozinha_e_para(amb):
    dxf(os.path.join(amb.p["entrada"], "P1", "t.dxf"), ["1-U4"])
    amb.vigia.iniciar()
    try:
        assert amb.vigia.ativo()
        limite = time.time() + 15
        while time.time() < limite and not os.path.exists(os.path.join(amb.p["processados"], "P1", "t.dxf")):
            time.sleep(0.2)
        assert os.path.exists(os.path.join(amb.p["processados"], "P1", "t.dxf"))
        assert amb.vigia.status()["ligado"] is True
    finally:
        amb.vigia.parar()
    assert not amb.vigia.ativo() and amb.vigia.status()["ligado"] is False
    amb.vigia.parar()          # idempotente


def test_a_vigia_sobrevive_a_erro_de_varredura(amb, monkeypatch):
    chamadas = []

    def varre_mal(self, cfg=None):
        chamadas.append(1)
        if len(chamadas) == 1:
            raise RuntimeError("falha")
        return []
    monkeypatch.setattr(pasta.Vigia, "varrer", varre_mal)
    amb.vigia.iniciar()
    try:
        limite = time.time() + 6
        while time.time() < limite and len(chamadas) < 2:
            time.sleep(0.2)
        assert len(chamadas) >= 2 and amb.vigia.ativo()
        assert "Falha inesperada" in amb.vigia.status()["ultimo_erro"]
    finally:
        amb.vigia.parar()


# ── configuração ────────────────────────────────────────────────────────────
def test_config_padrao_e_ida_e_volta(tmp_path, monkeypatch):
    monkeypatch.setattr(database, "DB_PATH", str(tmp_path / "c.db"))
    database.init_db()
    assert ca.carregar()["ligado"] is False and ca.carregar()["intervalo_s"] == 5
    salvo = ca.salvar({"pasta_base": str(tmp_path / "b"), "user_id": " 7 ", "ligado": 1})
    assert salvo["user_id"] == "7" and salvo["ligado"] is True and ca.carregar() == salvo
    p = ca.pastas(salvo)
    assert p["entrada"] == str(tmp_path / "b" / "entrada") and set(p) == {"entrada", "processados", "erros", "saida"}
    assert not os.path.exists(p["entrada"]) and ca.garantir_pastas(salvo) and os.path.isdir(p["saida"])


@pytest.mark.parametrize("cfg,trecho", [
    ({"pasta_base": "relativa/x"}, "absoluto"),
    ({"intervalo_s": 0}, "intervalo"), ({"intervalo_s": "abc"}, "números"), ({"estabilizacao_s": -1}, "estabilização"),
    ({"pasta_base": "/tmp/zz", "pasta_saida": "/tmp/zz/entrada/saida"}, "dentro da pasta de entrada"),
    ({"pasta_base": "/tmp/zz", "pasta_erros": "/tmp/zz/entrada"}, "dentro da pasta de entrada"),
    ({"pasta_base": "/tmp/zz", "pasta_processados": "/tmp/zz/x", "pasta_erros": "/tmp/zz/x"}, "diferentes"),
])
def test_config_invalida_e_recusada(cfg, trecho, tmp_path, monkeypatch):
    monkeypatch.setattr(database, "DB_PATH", str(tmp_path / "c.db"))
    database.init_db()
    with pytest.raises(ca.ErroConfig, match=trecho):
        ca.salvar(cfg)


def test_arquivo_que_nao_abre_ainda_espera(amb, monkeypatch):
    monkeypatch.setattr(pasta, "_abre", lambda caminho: False)       # p.ex. travado por outro programa (Windows)
    dxf(os.path.join(amb.p["entrada"], "P1", "preso.dxf"), ["1-U4"])
    for _ in range(3):
        assert amb.vigia.varrer(amb.cfg) == []
    assert os.path.exists(os.path.join(amb.p["entrada"], "P1", "preso.dxf"))


def test_parar_interrompe_a_fila_entre_arquivos(amb):
    for i in range(2):
        dxf(os.path.join(amb.p["entrada"], "P1", f"f{i}.dxf"), [f"{i + 1}-U4"])
    amb.vigia.varrer(amb.cfg)
    amb.vigia._parar.set()
    assert amb.vigia.varrer(amb.cfg) == []
    assert sorted(os.listdir(os.path.join(amb.p["entrada"], "P1"))) == ["f0.dxf", "f1.dxf"]


def test_a_vigia_respeita_o_intervalo(amb, monkeypatch):
    chamadas = []
    monkeypatch.setattr(pasta.Vigia, "varrer", lambda self, cfg=None: chamadas.append(time.time()) or [])
    amb.vigia.iniciar()
    time.sleep(2.6)
    amb.vigia.parar()
    assert 2 <= len(chamadas) <= 4          # intervalo_s = 1 (e não uma varredura sem pausa)


def test_varreduras_simultaneas_nao_processam_o_mesmo_arquivo_duas_vezes(amb):
    """A thread da vigia e `POST /varrer` podem coincidir: uma varredura por vez, cada arquivo tratado uma única vez."""
    for i in range(3):
        dxf(os.path.join(amb.p["entrada"], "P1", f"s{i}.dxf"), [f"{i + 1}-U4"])
    amb.vigia.varrer(amb.cfg)                      # observa
    resultados = []
    ts = [threading.Thread(target=lambda: resultados.append(amb.vigia.varrer(amb.cfg))) for _ in range(4)]
    for t in ts:
        t.start()
    for t in ts:
        t.join()
    todos = [x for r in resultados for x in r]
    assert sorted(x["arquivo"] for x in todos) == ["s0.dxf", "s1.dxf", "s2.dxf"] and all(x["status"] != "erro" for x in todos)
    assert len(execucoes.listar()) == 3


# ── arquivo EM USO (Windows: dá para ler, mas não para mover) ──────────────────────────────────────────────────
def _em_uso(monkeypatch, bloqueados):
    """Simula o Windows: enquanto o nome estiver em `bloqueados`, mover (os.rename) dá PermissionError."""
    original = os.rename

    def rename(origem, destino, *a, **k):
        if os.path.basename(origem) in bloqueados:
            raise PermissionError(13, "O arquivo está sendo usado por outro processo", origem)
        return original(origem, destino, *a, **k)
    monkeypatch.setattr(pasta.os, "rename", rename)


def test_arquivo_em_uso_e_processado_uma_vez_e_movido_quando_liberar(amb, monkeypatch):
    bloqueados = {"preso.dxf"}
    _em_uso(monkeypatch, bloqueados)
    alvo = dxf(os.path.join(amb.p["entrada"], "P1", "preso.dxf"), ["1-U4"])
    r = varrer2(amb)
    assert [x["status"] for x in r] == ["ok"]
    assert os.path.exists(alvo)                                         # não conseguiu mover: continua na entrada
    for _ in range(3):
        assert amb.vigia.varrer(amb.cfg) == []                          # NÃO é reprocessado nem vira "duplicado" nem "erro"
    assert len(execucoes.listar()) == 1 and execucoes.buscar(r[0]["execucao"])["obra_id"]
    assert execucoes.buscar(r[0]["execucao"])["arquivo_caminho"] == os.path.abspath(alvo)       # ainda aponta para a entrada
    bloqueados.clear()                                                  # o programa que segurava o arquivo soltou
    assert amb.vigia.varrer(amb.cfg) == []
    destino = os.path.join(amb.p["processados"], "P1", "preso.dxf")
    assert os.path.exists(destino) and not os.path.exists(alvo)
    assert execucoes.buscar(r[0]["execucao"])["arquivo_caminho"] == destino                      # histórico atualizado
    assert len(execucoes.listar()) == 1 and amb.vigia._adiados == {}


def test_arquivo_com_erro_em_uso_vai_para_erros_quando_liberar(amb, monkeypatch):
    bloqueados = {"nota.txt"}
    _em_uso(monkeypatch, bloqueados)
    os.makedirs(os.path.join(amb.p["entrada"], "P1"))
    alvo = os.path.join(amb.p["entrada"], "P1", "nota.txt")
    open(alvo, "w").write("x")
    r = varrer2(amb)
    assert r[0]["status"] == "erro" and os.path.exists(alvo)
    amb.vigia.varrer(amb.cfg)
    assert len(execucoes.listar("erro")) == 1                           # um registro só, apesar de várias varreduras
    bloqueados.clear()
    amb.vigia.varrer(amb.cfg)
    destino = os.path.join(amb.p["erros"], "P1", "nota.txt")
    assert os.path.exists(destino) and os.path.exists(destino + ".erro.txt") and not os.path.exists(alvo)
    assert execucoes.listar("erro")[0]["arquivo_caminho"] == destino


def test_mover_entre_volumes_nao_deixa_copia_orfa(tmp_path, monkeypatch):
    """Cross-device: copia e apaga a origem; se não der para apagar (em uso), a CÓPIA é removida e o erro sobe."""
    origem = tmp_path / "a" / "x.dxf"
    origem.parent.mkdir()
    origem.write_text("dados")
    monkeypatch.setattr(pasta.os, "rename", lambda *a, **k: (_ for _ in ()).throw(OSError(18, "Invalid cross-device link")))
    destino = pasta._mover(str(origem), str(tmp_path / "b"))
    assert open(destino).read() == "dados" and not origem.exists()
    origem.write_text("dados2")
    real_remove = os.remove
    monkeypatch.setattr(pasta.os, "remove", lambda p: (_ for _ in ()).throw(PermissionError(13, "em uso")) if p == str(origem) else real_remove(p))
    with pytest.raises(PermissionError):
        pasta._mover(str(origem), str(tmp_path / "b"))
    assert origem.exists() and sorted(os.listdir(tmp_path / "b")) == ["x.dxf"]      # nenhuma cópia órfã
