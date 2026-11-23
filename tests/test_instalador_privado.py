import sqlite3
from pathlib import Path

import pytest

from scripts import compilar_instalador_privado as privado


def test_compila_com_executavel_env_e_copia_do_banco_sem_sessao(tmp_path, monkeypatch):
    executavel = tmp_path / "portatil.exe"
    executavel.write_bytes(b"MZ")
    env = tmp_path / ".env"
    env.write_bytes(b"SUPABASE_URL=https://exemplo.supabase.co\r\nSUPABASE_KEY=chave-de-teste\r\n")
    db = tmp_path / "origem.db"
    iscc = tmp_path / "ISCC.exe"
    iscc.touch()
    with sqlite3.connect(db) as banco:
        banco.executescript("""
            PRAGMA journal_mode=WAL;
            CREATE TABLE usuarios_locais (email);
            CREATE TABLE sessoes (access_token);
            CREATE TABLE tabela_orcamento (ativo);
            INSERT INTO usuarios_locais VALUES ('operador@example.com');
            INSERT INTO sessoes VALUES ('sessao-da-origem');
            INSERT INTO tabela_orcamento VALUES ('N3');
        """)

        def simular_inno(entradas, check):
            assert check
            assert entradas[-1] == str(privado.MODELO)
            assert "/DAppVersion=2.8.2" in entradas
            dados = Path(next(arg.split("=", 1)[1] for arg in entradas if arg.startswith("/DEntregaDir=")))
            assert sorted(item.name for item in dados.iterdir()) == [".env", "OrcaCore.exe", "banco_resumo.db"]
            assert (dados / ".env").read_bytes() == env.read_bytes()
            assert (dados / "OrcaCore.exe").read_bytes() == b"MZ"
            with sqlite3.connect(dados / "banco_resumo.db") as copia:
                assert copia.execute("SELECT ativo FROM tabela_orcamento").fetchall() == [("N3",)]
                assert copia.execute("SELECT email FROM usuarios_locais").fetchall() == [("operador@example.com",)]
                assert copia.execute("SELECT * FROM sessoes").fetchall() == []
            saida = Path(next(arg[2:] for arg in entradas if arg.startswith("/O")))
            (saida / privado.NOME_SETUP).write_bytes(b"MZ-setup-simulado")

        monkeypatch.setattr(privado.subprocess, "run", simular_inno)
        arquivo = privado.compilar(executavel, env, db, "2.8.2", tmp_path / "entrega", iscc)
        assert arquivo.read_bytes() == b"MZ-setup-simulado"
        assert banco.execute("SELECT access_token FROM sessoes").fetchall() == [("sessao-da-origem",)]
        with pytest.raises(privado.ErroPacote, match="Já existe"):
            privado.compilar(executavel, env, db, "2.8.2", tmp_path / "entrega", iscc)
    banco.close()


@pytest.mark.parametrize("conteudo", [
    "SUPABASE_URL=https://outro.onrender.com\nSUPABASE_KEY=chave",
    "SUPABASE_URL=http://exemplo.supabase.co\nSUPABASE_KEY=chave",
    "SUPABASE_URL=https://exemplo.supabase.co\n",
    "SUPABASE_URL=https://exemplo.supabase.co\nSUPABASE_KEY=chave\nAPP_MODE=server",
])
def test_env_precisa_configurar_o_supabase_direto_e_o_modo_desktop(tmp_path, conteudo):
    env = tmp_path / ".env"
    env.write_text(conteudo, encoding="utf-8")
    with pytest.raises(privado.ErroPacote):
        privado.validar_env(env)


def test_instalador_so_inclui_dados_na_compilacao_privada():
    modelo = privado.MODELO.read_text(encoding="utf-8-sig")
    arquivos = modelo.split("[Files]", 1)[1].split("[Icons]", 1)[0]
    privado_iss, publico_iss = arquivos.split("#else", 1)
    assert "#ifdef EntregaDir" in privado_iss
    linhas = [linha for linha in privado_iss.splitlines() if ".env" in linha or "banco_resumo.db" in linha]
    assert len(linhas) == 2
    assert all('DestDir: "{app}"' in linha and "onlyifdoesntexist uninsneveruninstall" in linha for linha in linhas)
    assert ".env" not in publico_iss and "banco_resumo.db" not in publico_iss
    assert "--online" not in modelo
