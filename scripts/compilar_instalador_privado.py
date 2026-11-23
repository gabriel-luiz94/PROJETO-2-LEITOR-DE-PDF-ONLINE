"""Compila um Setup privado com os arquivos de uma instalação local do OrçaCore."""

import argparse
from contextlib import closing
import os
from pathlib import Path
import re
import shutil
import sqlite3
import subprocess
import sys
import tempfile
from urllib.parse import urlsplit

from dotenv import dotenv_values

if __package__:
    from .montar_pacote_entrega import ErroPacote, copiar_banco_consistente
else:
    from montar_pacote_entrega import ErroPacote, copiar_banco_consistente


MODELO = Path(__file__).resolve().parents[1] / "installer" / "OrcaCore.iss"
NOME_SETUP = "OrcaCore-Setup-Privado.exe"


def validar_env(caminho):
    dados = dotenv_values(caminho, interpolate=False)
    try:
        url = urlsplit(dados.get("SUPABASE_URL") or "")
        if (url.scheme != "https" or not url.hostname or url.username or url.password
                or url.hostname == "onrender.com" or url.hostname.endswith(".onrender.com")):
            raise ValueError
    except ValueError:
        raise ErroPacote("Informe no .env o endereço HTTPS direto do Supabase.") from None
    if not (dados.get("SUPABASE_KEY") or "").strip():
        raise ErroPacote("O .env precisa conter SUPABASE_KEY para o login.")
    if (dados.get("APP_MODE") or "desktop").strip() != "desktop" or (dados.get("HEADLESS") or "").strip() == "1":
        raise ErroPacote("Use o .env da instalação desktop do OrçaCore.")


def preparar_banco(origem, destino):
    copiar_banco_consistente(origem, destino)
    with closing(sqlite3.connect(destino)) as banco:
        tabelas = {linha[0] for linha in banco.execute("SELECT name FROM sqlite_master WHERE type = 'table'")}
        if not {"usuarios_locais", "tabela_orcamento", "sessoes"} <= tabelas:
            raise ErroPacote("O arquivo selecionado não parece ser um banco local do OrçaCore.")
        if banco.execute("PRAGMA integrity_check").fetchone()[0] != "ok":
            raise ErroPacote("O banco de origem falhou na verificação de integridade.")
        # A cópia leva os dados locais. A pessoa autentica a própria sessão ao abrir.
        banco.execute("DELETE FROM sessoes")
        banco.commit()
        banco.execute("PRAGMA journal_mode=DELETE")
        banco.execute("VACUUM")


def achar_iscc(caminho):
    if caminho:
        return Path(caminho)
    no_path = shutil.which("ISCC.exe")
    if no_path:
        return Path(no_path)
    return Path(os.environ.get("ProgramFiles(x86)", r"C:\Program Files (x86)")) / "Inno Setup 6" / "ISCC.exe"


def compilar(executavel, env, banco, versao, saida, iscc):
    for caminho, rotulo in ((executavel, "Executável"), (env, ".env"), (banco, "Banco"), (iscc, "Inno Setup")):
        if not caminho.is_file():
            raise ErroPacote(f"{rotulo} não encontrado em {caminho}.")
    if not re.fullmatch(r"\d+\.\d+\.\d+", versao):
        raise ErroPacote("Informe a versão do executável no formato X.Y.Z.")
    validar_env(env)
    destino = saida / NOME_SETUP
    if destino.exists():
        raise ErroPacote(f"Já existe um instalador privado em {saida}. Escolha outra pasta de saída.")

    with tempfile.TemporaryDirectory(prefix="orcacore_instalador_privado_") as temporario:
        pasta = Path(temporario)
        dados = pasta / "dados"
        compilado = pasta / "compilado"
        dados.mkdir()
        compilado.mkdir()
        shutil.copyfile(executavel, dados / "OrcaCore.exe")
        shutil.copyfile(env, dados / ".env")
        preparar_banco(banco, dados / "banco_resumo.db")
        subprocess.run(
            [str(iscc.resolve()), f"/DAppVersion={versao}", f"/DEntregaDir={dados}", f"/O{compilado}", str(MODELO)],
            check=True,
        )
        resultado = compilado / NOME_SETUP
        if not resultado.is_file():
            raise ErroPacote("O Inno Setup terminou sem gerar o instalador privado.")
        saida.mkdir(parents=True, exist_ok=True)
        try:
            arquivo = destino.open("xb")
        except FileExistsError:
            raise ErroPacote(f"Já existe um instalador privado em {saida}. Escolha outra pasta de saída.") from None
        try:
            with arquivo, resultado.open("rb") as origem:
                shutil.copyfileobj(origem, arquivo)
        except BaseException:
            destino.unlink(missing_ok=True)
            raise
    return destino


def main(argv=None):
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument("--executavel", required=True, type=Path, help="OrcaCore.exe portátil da release")
    parser.add_argument("--versao", required=True, help="versão X.Y.Z do executável")
    parser.add_argument("--env", required=True, type=Path, help="arquivo .env que será entregue")
    parser.add_argument("--banco", required=True, type=Path, help="banco_resumo.db de origem")
    parser.add_argument("--saida", required=True, type=Path, help="pasta privada para o instalador")
    parser.add_argument("--iscc", type=Path, help="caminho do compilador Inno Setup 6")
    args = parser.parse_args(argv)
    try:
        instalador = compilar(args.executavel, args.env, args.banco, args.versao, args.saida, achar_iscc(args.iscc))
    except (ErroPacote, OSError, sqlite3.Error, subprocess.CalledProcessError) as erro:
        print(f"Erro ao compilar o instalador em modo privado. {erro}", file=sys.stderr)
        return 1
    print(f"Instalador privado gerado em {instalador}")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
