"""Garante o isolamento (tests/conftest.py): nenhum teste enxerga a nuvem, chaves de IA ou o banco real."""
import os

import config
import database
from services import supabase_client


def test_sem_credenciais_de_nuvem_nem_de_ia():
    for chave in ("SUPABASE_URL", "SUPABASE_KEY", "GEMINI_API_KEY", "GOOGLE_API_KEY", "OPENAI_API_KEY"):
        assert os.environ.get(chave, "") == "", chave
    assert supabase_client.get_supabase() is None


def test_banco_dos_testes_nao_e_o_real(tmp_path):
    real = os.path.join(config.BASE_DIR, "banco_resumo.db")
    assert os.path.normcase(os.path.abspath(database.DB_PATH)) != os.path.normcase(os.path.abspath(real))
    assert os.path.normcase(os.path.abspath(config.DB_PATH)) != os.path.normcase(os.path.abspath(real))


def test_credenciais_do_dotenv_nao_vazam_para_a_sessao(monkeypatch):
    """Mesmo que o .env tenha chaves, o ambiente zerado vence (load_dotenv não sobrescreve) e o cliente continua None."""
    monkeypatch.setenv("SUPABASE_URL", "")
    supabase_client.reset_supabase_client()
    assert supabase_client.get_supabase() is None
