"""TASK-230 — scripts/montar_pacote_entrega.py: pacote de entrega com banco higienizado, .env limpo e conferência de que o usuário entra e trabalha de cara."""
import hashlib
import importlib.util
import json
import sqlite3
import zipfile
from pathlib import Path

import pytest

import config
import database

RAIZ = Path(config.BASE_DIR)


def _carregar():
    spec = importlib.util.spec_from_file_location("montar_pacote_entrega", RAIZ / "scripts" / "montar_pacote_entrega.py")
    m = importlib.util.module_from_spec(spec)
    spec.loader.exec_module(m)
    return m


@pytest.fixture()
def origem(tmp_path, monkeypatch):
    """Banco «do desenvolvedor»: usuário dono com sessão, obra, base pessoal, chaves de IA salvas e histórico — tudo o que NÃO pode ir junto."""
    db = tmp_path / "origem.db"
    monkeypatch.setattr(database, "DB_PATH", str(db))
    database.init_db()
    con = sqlite3.connect(db)
    con.execute("INSERT INTO usuarios_locais (email, senha_hash, is_admin, role) VALUES ('dono@empresa.com', ?, 1, 'admin')", (database.hash_password("SenhaDoDono#1"),))
    dono = str(con.execute("SELECT id FROM usuarios_locais WHERE email='dono@empresa.com'").fetchone()[0])
    con.execute("INSERT INTO sessoes (user_id, email, access_token, refresh_token, is_admin, role) VALUES (?, 'dono@empresa.com', 'TOKEN-SECRETO-DO-DONO', 'x', 1, 'admin')", (dono,))
    con.execute("INSERT INTO configuracoes (chave, valor) VALUES ('ai_reserva_api_key', 'CHAVE-IA-SECRETA-123'), ('cloudconvert_api_key', 'CHAVE-CLOUD-456'), ('tema', 'escuro')")
    con.execute("INSERT INTO ativo_versoes (projeto, ativo, linhas_json, usuario) VALUES ('081', 'N3', '[]', 'dono@empresa.com')")
    con.execute("DELETE FROM tabela_orcamento")
    for i in range(3):
        con.execute("INSERT INTO tabela_orcamento (user_id, ativo, desc_ativo, projeto) VALUES (?, ?, ?, 'PERNAMBUCO')", (dono, f"BASE-DO-DONO-{i}", "descrição pessoal"))
    cols = [r[1] for r in con.execute("PRAGMA table_info(obras)")]
    valores = {"user_id": dono, "nome": "OBRA-PARTICULAR-DO-DONO", "numero_obra": "9", "data": "2026-11-20", "dados_json": "{}"}
    usar = {c: v for c, v in valores.items() if c in cols}
    con.execute(f"INSERT INTO obras ({', '.join(usar)}) VALUES ({', '.join('?' for _ in usar)})", tuple(usar.values()))
    con.commit()
    con.close()
    env = tmp_path / "origem.env"
    env.write_text("SUPABASE_URL=https://x.supabase.co\nSUPABASE_KEY=chave-supabase\nGEMINI_API_KEY=chave-gemini\nJWT_SECRET=nao-leva\nADMIN_PASSWORD=admin-do-dono\n", encoding="utf-8")
    return {"db": db, "env": env, "dono": dono, "tmp": tmp_path}


def _rodar(m, origem, *extra, email="novo@cliente.com", senha="SenhaNova12345"):
    saida = origem["tmp"] / "saida"
    codigo = m.main(["--saida", str(saida), "--email", email, "--senha", senha, "--banco", str(origem["db"]), "--env", str(origem["env"]), *extra])
    return codigo, saida


def test_pacote_completo_sem_nada_do_dono_e_com_login_funcionando(origem, capsys):
    m = _carregar()
    instalador = origem["tmp"] / "OrcaCore-Setup.exe"
    instalador.write_bytes(b"MZ-falso")
    codigo, saida = _rodar(m, origem, "--instalador", str(instalador), "--zip", "--base-de", "dono@empresa.com")
    texto = capsys.readouterr().out
    assert codigo == 0, texto
    pasta = next(p for p in saida.iterdir() if p.is_dir())
    assert {p.name for p in pasta.iterdir()} >= {"Instalar.bat", "LEIA-ME.txt", "OrcaCore-Setup.exe", "SHA256SUMS.txt", "dados"}
    assert "OK   login com o e-mail e a senha entregues" in texto and "FALHA" not in texto                    # a conferência abriu o programa sobre o pacote
    # banco: só a conta nova, base do dono como base padrão, nada pessoal
    con = sqlite3.connect(pasta / "dados" / "banco_resumo.db")
    assert [r[0] for r in con.execute("SELECT email FROM usuarios_locais")] == ["novo@cliente.com"]
    assert con.execute("SELECT COUNT(*) FROM sessoes").fetchone()[0] == 0 and con.execute("SELECT COUNT(*) FROM obras").fetchone()[0] == 0
    assert sorted(r[0] for r in con.execute("SELECT ativo FROM tabela_orcamento")) == ["BASE-DO-DONO-0", "BASE-DO-DONO-1", "BASE-DO-DONO-2"]
    assert con.execute("SELECT COUNT(*) FROM tabela_orcamento WHERE user_id IS NOT NULL").fetchone()[0] == 0
    assert {r[0] for r in con.execute("SELECT chave FROM configuracoes")} == {"tema"}                          # chaves salvas saíram; preferências ficam
    assert con.execute("SELECT usuario FROM ativo_versoes").fetchone()[0] == ""
    assert con.execute("PRAGMA integrity_check").fetchone()[0] == "ok"
    con.close()
    # nada do dono dentro do arquivo (VACUUM apagou de verdade) nem em lugar nenhum do pacote
    proibidos = [b"CHAVE-IA-SECRETA-123", b"CHAVE-CLOUD-456", b"TOKEN-SECRETO-DO-DONO", b"dono@empresa.com", b"OBRA-PARTICULAR-DO-DONO", b"chave-supabase", b"chave-gemini", b"nao-leva", b"admin-do-dono", b"SenhaDoDono#1"]
    for arquivo in pasta.rglob("*"):
        if arquivo.is_file():
            dados = arquivo.read_bytes()
            assert not any(p in dados for p in proibidos), arquivo
    assert not list((pasta / "dados").glob("*-wal")) and not list((pasta / "dados").glob("*-shm"))
    # .env: sem a senha do usuário (ela vai por outro canal), sem nuvem nem IA
    env = (pasta / "dados" / ".env").read_text(encoding="utf-8")
    assert "ADMIN_EMAIL=novo@cliente.com" in env and "SenhaNova12345" not in env and "SUPABASE" not in env and "GEMINI" not in env and "JWT_SECRET" not in env
    # senha só no arquivo ACESSO, fora da pasta e fora do zip
    acesso = next(p for p in saida.glob("ACESSO_*.txt"))
    assert "SenhaNova12345" in acesso.read_text(encoding="utf-8") and "novo@cliente.com" in acesso.read_text(encoding="utf-8")
    with zipfile.ZipFile(next(saida.glob("*.zip"))) as z:
        assert not any("ACESSO" in n for n in z.namelist())
        assert not any(b"SenhaNova12345" in z.read(n) for n in z.namelist())
    # somas de verificação batem
    for linha in (pasta / "SHA256SUMS.txt").read_text(encoding="utf-8").splitlines():
        h, nome = linha.split("  ", 1)
        assert hashlib.sha256((pasta / nome).read_bytes()).hexdigest() == h
    bat = (pasta / "Instalar.bat").read_bytes().decode("utf-8")
    assert "OrcaCore-Setup.exe\" /SILENT" in bat and "banco_resumo.db.anterior" in bat and "LOCALAPPDATA%\\Programs\\OrcaCore" in bat and "\r\n" in bat


def test_so_leva_nuvem_e_ia_quando_pedido(origem):
    m = _carregar()
    codigo, saida = _rodar(m, origem, "--base-de", "dono@empresa.com", "--com-nuvem", "--com-chaves-ia")
    assert codigo == 0
    env = (next(p for p in saida.iterdir() if p.is_dir()) / "dados" / ".env").read_text(encoding="utf-8")
    assert "SUPABASE_URL=https://x.supabase.co" in env and "SUPABASE_KEY=chave-supabase" in env and "GEMINI_API_KEY=chave-gemini" in env
    assert "JWT_SECRET" not in env and "admin-do-dono" not in env


def test_obras_so_vao_com_a_opcao_e_passam_para_o_novo_usuario(origem):
    m = _carregar()
    codigo, saida = _rodar(m, origem, "--base-de", "dono@empresa.com", "--com-obras")
    assert codigo == 0
    con = sqlite3.connect(next(p for p in saida.iterdir() if p.is_dir()) / "dados" / "banco_resumo.db")
    novo = str(con.execute("SELECT id FROM usuarios_locais WHERE email='novo@cliente.com'").fetchone()[0])
    assert [r[0] for r in con.execute("SELECT user_id FROM obras")] == [novo]


def test_recusa_o_que_deixaria_o_usuario_sem_conseguir_trabalhar(origem, capsys):
    m = _carregar()
    con = sqlite3.connect(origem["db"])                                                           # base de mais de um dono sem dizer de quem
    con.execute("INSERT INTO usuarios_locais (email, senha_hash, is_admin, role) VALUES ('outro@empresa.com', 'x', 0, 'operador')")
    con.execute("INSERT INTO tabela_orcamento (user_id, ativo, projeto) VALUES ('999', 'OUTRA-BASE', 'X')")
    con.commit()
    con.close()
    codigo, saida = _rodar(m, origem)
    assert codigo == 2 and "--base-de" in capsys.readouterr().err
    assert not saida.exists() or not any(saida.iterdir())                                         # não deixa meia entrega
    assert _rodar(m, origem, "--base-de", "dono@empresa.com", email="sem-arroba")[0] == 2
    assert _rodar(m, origem, "--base-de", "dono@empresa.com", senha="curta")[0] == 2
    vazio = sqlite3.connect(origem["db"])
    vazio.execute("DELETE FROM tabela_orcamento")
    vazio.commit()
    vazio.close()
    assert _rodar(m, origem, "--base-de", "dono@empresa.com")[0] == 2                            # sem base de orçamento não há o que entregar
    assert _rodar(m, {**origem, "db": origem["tmp"] / "naoexiste.db"})[0] == 2
