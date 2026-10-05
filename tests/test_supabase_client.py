"""tests/test_supabase_client.py — TASK-039/041: TTL proativo do cliente Supabase e reset reativo
em falha de rede (nunca mais silenciosa). tests/conftest.py já garante que nenhum teste use um
Supabase real; aqui simulamos um cliente FALSO via monkeypatch de create_client."""
import httpx
import pytest

from services import supabase_client as sc


class _ClienteFalso:
    """Sentinela: cada instância é um objeto distinto, para comparar identidade entre chamadas."""


@pytest.fixture
def credenciais_falsas(monkeypatch):
    monkeypatch.setenv("SUPABASE_URL", "https://exemplo.supabase.co")
    monkeypatch.setenv("SUPABASE_KEY", "chave-falsa")
    chamadas = []

    def _create_client_falso(url, key):
        chamadas.append((url, key))
        return _ClienteFalso()

    monkeypatch.setattr(sc, "create_client", _create_client_falso)
    # O autouse de conftest.py já chamou get_supabase() sem credenciais (fixou "não configurado"
    # no singleton); força reavaliar agora que as credenciais falsas acima estão no ambiente.
    sc.reset_supabase_client()
    return chamadas


def test_primeira_chamada_cria_cliente(credenciais_falsas):
    cliente = sc.get_supabase()
    assert isinstance(cliente, _ClienteFalso)
    assert len(credenciais_falsas) == 1


def test_chamadas_seguintes_dentro_do_ttl_reaproveitam_o_mesmo_cliente(credenciais_falsas):
    c1 = sc.get_supabase()
    c2 = sc.get_supabase()
    assert c1 is c2
    assert len(credenciais_falsas) == 1


def test_ttl_expirado_recria_o_cliente(credenciais_falsas, monkeypatch):
    c1 = sc.get_supabase()
    # Simula o tempo passando sem precisar esperar de verdade: volta o relógio-base do cliente.
    monkeypatch.setattr(sc, "_client_criado_em", sc._client_criado_em - sc._TTL_SEGUNDOS - 1)
    c2 = sc.get_supabase()
    assert c1 is not c2
    assert len(credenciais_falsas) == 2


def test_registrar_falha_de_rede_reseta_o_cliente(credenciais_falsas):
    c1 = sc.get_supabase()
    sc.registrar_falha("teste.rede", httpx.ConnectError("conexão recusada"))
    c2 = sc.get_supabase()
    assert c1 is not c2
    assert len(credenciais_falsas) == 2


def test_registrar_falha_de_dados_nao_reseta_o_cliente(credenciais_falsas):
    c1 = sc.get_supabase()
    sc.registrar_falha("teste.dados", ValueError("coluna inexistente"))
    c2 = sc.get_supabase()
    assert c1 is c2
    assert len(credenciais_falsas) == 1


def test_registrar_falha_sempre_loga_mesmo_sem_resetar(credenciais_falsas, caplog):
    with caplog.at_level("WARNING"):
        sc.registrar_falha("teste.contexto-x", ValueError("algo"))
    assert any("teste.contexto-x" in r.message for r in caplog.records)


@pytest.mark.parametrize("exc,esperado", [
    (httpx.ConnectError("x"), True),
    (httpx.ConnectTimeout("x"), True),
    (httpx.ReadTimeout("x"), True),
    (ConnectionError("x"), True),
    (TimeoutError("x"), True),
    (ValueError("x"), False),
    (KeyError("x"), False),
])
def test_eh_erro_de_rede(exc, esperado):
    assert sc._eh_erro_de_rede(exc) is esperado


def test_sem_credenciais_nao_cria_cliente_nem_expira(monkeypatch):
    monkeypatch.setenv("SUPABASE_URL", "")
    monkeypatch.setenv("SUPABASE_KEY", "")
    sc.reset_supabase_client()
    assert sc.get_supabase() is None
    assert sc.get_supabase() is None  # segunda chamada não tenta recriar nem falha
