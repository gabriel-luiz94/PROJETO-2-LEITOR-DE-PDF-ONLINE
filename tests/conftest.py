"""
tests/conftest.py — ISOLAMENTO TOTAL dos testes (nuvem, chaves de IA e banco local).

Por que existe: várias rotas gravam no Supabase (regras, ajustes, prompts, obras…). Sem credenciais no ambiente, o código usa só o SQLite
temporário e os testes são inofensivos — foi assim em Linux. Num computador com `.env` (SUPABASE_URL/SUPABASE_KEY), os testes liam e GRAVAVAM
na nuvem REAL (aconteceu em 2026-10-02: ver .ai/CHANGELOG.md). Este arquivo roda ANTES de qualquer teste e:

  1. zera as credenciais da nuvem e das IAs no ambiente (o `load_dotenv` não sobrescreve variáveis já definidas);
  2. aponta o banco local para uma pasta temporária (o `banco_resumo.db` real nunca é aberto pelos testes);
  3. recusa rodar (`pytest.exit`) se, ainda assim, existir um cliente Supabase.
Um teste que precise de nuvem usa um cliente FALSO (monkeypatch), como tests/test_ajustes_cadastro.py.
"""
import os
import tempfile

import pytest

for _chave in ("SUPABASE_URL", "SUPABASE_KEY", "SUPABASE_SERVICE_KEY", "SUPABASE_ANON_KEY", "GEMINI_API_KEY", "GOOGLE_API_KEY", "OPENAI_API_KEY", "CLOUDCONVERT_API_KEY"):
    os.environ[_chave] = ""

_PASTA_TESTES = tempfile.mkdtemp(prefix="leitor_testes_")

import config  # noqa: E402  (depois de zerar o ambiente)

config.DB_PATH = os.path.join(_PASTA_TESTES, "banco_isolado.db")      # database.py importa este valor por nome


def pytest_sessionstart(session):
    from services import supabase_client
    supabase_client.reset_supabase_client()
    if supabase_client.get_supabase() is not None:
        pytest.exit("RECUSADO: os testes encontraram um cliente Supabase configurado e poderiam gravar na nuvem real. "
                    "Remova SUPABASE_URL/SUPABASE_KEY do ambiente de teste.", returncode=3)


@pytest.fixture(autouse=True)
def _sem_nuvem_nem_banco_real(monkeypatch, tmp_path):
    """Cada teste começa SEM nuvem e com um banco temporário próprio (os testes que querem outro banco o definem depois)."""
    import database
    from services import supabase_client
    supabase_client.reset_supabase_client()
    monkeypatch.setattr(database, "DB_PATH", str(tmp_path / "isolado.db"))
    assert supabase_client.get_supabase() is None, "teste com Supabase real configurado"
    yield
    supabase_client.reset_supabase_client()
