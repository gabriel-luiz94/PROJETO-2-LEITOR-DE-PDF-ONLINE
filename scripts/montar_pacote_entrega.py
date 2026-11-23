#!/usr/bin/env python3
"""
Monta o PACOTE DE ENTREGA do OrçaCore para outra pessoa (TASK-230): instalador + dados (banco higienizado, .env limpo, skills) + instruções.

O que o pacote garante (e o script confere ANTES de terminar, rodando o programa de verdade sobre uma cópia):
  - o banco abre, passa na verificação de integridade e tem a base de orçamento e as regras;
  - o usuário recebe uma conta de administrador NOVA (e-mail/senha que você escolhe ou que o script gera) e o login funciona;
  - nada seu vai junto: sem usuários/senhas/sessões, sem chaves de IA salvas, sem obras e históricos pessoais (use as opções --com-* para levar algo de propósito);
  - o .env só leva o que você mandar (--com-nuvem, --com-chaves-ia).

Uso (a partir da raiz do projeto):
    python scripts/montar_pacote_entrega.py --saida entrega --instalador dist/OrcaCore-Setup.exe --email fulano@empresa.com
    python scripts/montar_pacote_entrega.py --saida entrega --instalador dist/OrcaCore-Setup.exe --banco "C:\\...\\banco_resumo.db" --base-de voce@empresa.com --zip

Resultado em <saida>/OrcaCore-entrega-AAAAMMDD-HHMM/:
    OrcaCore-Setup.exe, Instalar.bat (instala o programa e copia os dados), dados/ (banco_resumo.db, .env, skills/), LEIA-ME.txt, SHA256SUMS.txt
e, AO LADO da pasta, ACESSO_<pasta>.txt com e-mail e senha — mande por OUTRO canal (não vai dentro do zip).
"""
import argparse
import datetime
import hashlib
import json
import os
import re
import secrets
import shutil
import sqlite3
import subprocess
import sys
import tempfile
import zipfile
from pathlib import Path

RAIZ = Path(__file__).resolve().parents[1]
if str(RAIZ) not in sys.path:
    sys.path.insert(0, str(RAIZ))

# Tabelas que identificam pessoas ou o estado de sincronização do computador de origem: nunca vão.
TABELAS_PESSOAIS = ["usuarios_locais", "sessoes", "audit_log", "sync_log", "conflitos_locais", "copia_local_meta", "copia_local_base"]
# Trabalho do usuário de origem: só vai com --com-obras (e passa a ser do novo usuário).
TABELAS_DE_OBRAS = ["obras", "historico_rec"]
PADRAO_SEGREDO = re.compile(r"(api_?key|token|secret|senha|password)", re.I)
VARS_NUVEM = ["SUPABASE_URL", "SUPABASE_KEY"]
VARS_IA = ["GEMINI_API_KEY", "GOOGLE_API_KEY", "OPENAI_API_KEY", "ANTHROPIC_API_KEY", "AI_OPENAI_ALLOWED_BASE_URLS", "AI_RATE_LIMIT_POR_MINUTO"]
TABELAS_COM_AUTOR = {"ativo_versoes": "usuario"}          # coluna com o e-mail de quem publicou (apaga o texto, mantém a linha)


class ErroPacote(Exception):
    """Algo impede de entregar com segurança (a mensagem diz o que fazer)."""


def _tabelas(con):
    return [r[0] for r in con.execute("SELECT name FROM sqlite_master WHERE type='table' AND name NOT LIKE 'sqlite_%'")]


def _colunas(con, tabela):
    return [r[1] for r in con.execute(f'PRAGMA table_info("{tabela}")')]


def senha_aleatoria(n=14):
    alfabeto = "abcdefghjkmnpqrstuvwxyzABCDEFGHJKLMNPQRSTUVWXYZ23456789"            # sem caracteres que se confundem (0/O, 1/l/I)
    return "".join(secrets.choice(alfabeto) for _ in range(n))


# ── banco ─────────────────────────────────────────────────────────────────────────────────────────
def copiar_banco_consistente(origem, destino):
    """Cópia consistente (usa a API de backup do SQLite: vale mesmo com o programa aberto e o banco em modo WAL)."""
    origem = Path(origem)
    if not origem.is_file():
        raise ErroPacote(f"Banco não encontrado: {origem}")
    src = sqlite3.connect(origem.resolve().as_uri() + "?mode=ro", uri=True)
    try:
        dst = sqlite3.connect(str(destino))
        try:
            src.backup(dst)
        finally:
            dst.close()
    finally:
        src.close()


def base_efetiva(con, usuario_ids):
    """A base de orçamento que o usuário de origem enxerga (mesma ordem do programa: edições pessoais > mestre > seed), como lista de dicts sem id/user_id."""
    con.row_factory = sqlite3.Row
    colunas = [c for c in _colunas(con, "tabela_orcamento") if c not in ("id", "user_id")]
    for uid in usuario_ids:
        linhas = con.execute("SELECT * FROM tabela_orcamento WHERE user_id = ? ORDER BY id", (uid,)).fetchall()
        if linhas:
            return [{c: l[c] for c in colunas} for l in linhas], "edições pessoais de " + str(uid)
    if "tabela_orcamento_master" in _tabelas(con):
        cm = [c for c in colunas if c in _colunas(con, "tabela_orcamento_master")]
        linhas = con.execute("SELECT * FROM tabela_orcamento_master ORDER BY id").fetchall()
        if linhas:
            return [{c: l[c] for c in cm} for l in linhas], "base mestre"
    linhas = con.execute("SELECT * FROM tabela_orcamento WHERE user_id IS NULL ORDER BY id").fetchall()
    return [{c: l[c] for c in colunas} for l in linhas], "base padrão do programa"


def higienizar_banco(caminho, email, senha, base_de=None, com_obras=False, com_historico=False):
    """Deixa o banco copiado pronto para entregar. Devolve um relatório (dict). Levanta ErroPacote se algo impedir de entregar."""
    from database import hash_password
    con = sqlite3.connect(str(caminho))
    rel = {"removido": {}, "base": "", "linhas_base": 0, "segredos_removidos": [], "colunas_com_usuario": []}
    try:
        if con.execute("PRAGMA integrity_check").fetchone()[0] != "ok":
            raise ErroPacote("O banco de origem falhou na verificação de integridade (PRAGMA integrity_check). Restaure um backup antes de entregar.")
        tabelas = _tabelas(con)
        if "tabela_orcamento" not in tabelas or "usuarios_locais" not in tabelas:
            raise ErroPacote("O arquivo não parece um banco do OrçaCore (faltam tabelas). Use o banco_resumo.db da pasta do programa.")
        # 1) quem é o «usuário de origem» (dono da base pessoal): o que o --base-de aponta, ou o único com edições
        ids_por_email = {r[0].lower(): [str(r[1]), str(r[2])] for r in con.execute("SELECT email, id, COALESCE(nuvem_id, id) FROM usuarios_locais")}
        candidatos = []
        if base_de:
            ids = ids_por_email.get(base_de.lower())
            if not ids:
                raise ErroPacote(f"--base-de {base_de}: não há esse usuário no banco. Usuários: {', '.join(sorted(ids_por_email)) or '(nenhum)'}.")
            candidatos = ids
        else:
            donos = [str(r[0]) for r in con.execute("SELECT DISTINCT user_id FROM tabela_orcamento WHERE user_id IS NOT NULL")]
            if len(donos) > 1:
                raise ErroPacote("Há edições pessoais da base de mais de um usuário; diga de quem é a base a entregar com --base-de EMAIL.")
            candidatos = donos
        linhas, origem = base_efetiva(con, candidatos)
        if not linhas:
            raise ErroPacote("A base de orçamento do banco está vazia: o usuário não conseguiria orçar. Abra o programa uma vez para importar a base ou escolha outro banco.")
        rel["base"], rel["linhas_base"] = origem, len(linhas)
        # 2) a base entregue vira a base padrão (user_id NULL); personais e mestre saem (a nuvem, se houver, refaz a mestre)
        con.execute("DELETE FROM tabela_orcamento")
        colunas = list(linhas[0].keys())
        con.executemany(f"INSERT INTO tabela_orcamento (user_id, {', '.join(colunas)}) VALUES (NULL, {', '.join('?' for _ in colunas)})", [tuple(l[c] for c in colunas) for l in linhas])
        if "tabela_orcamento_master" in tabelas:
            con.execute("DELETE FROM tabela_orcamento_master")
        # 3) o que é pessoal sai
        for t in TABELAS_PESSOAIS + ([] if com_historico else [t for t in tabelas if t.endswith("_historico")]):
            if t in tabelas:
                rel["removido"][t] = con.execute(f'DELETE FROM "{t}"').rowcount
        if "sqlite_sequence" in [r[0] for r in con.execute("SELECT name FROM sqlite_master")]:
            con.execute("DELETE FROM sqlite_sequence WHERE name = 'usuarios_locais'")
        for t, col in TABELAS_COM_AUTOR.items():
            if t in tabelas and col in _colunas(con, t):
                con.execute(f'UPDATE "{t}" SET "{col}" = \'\'')
        for r in con.execute("SELECT chave FROM configuracoes").fetchall() if "configuracoes" in tabelas else []:
            if PADRAO_SEGREDO.search(r[0]):
                rel["segredos_removidos"].append(r[0])
        for chave in rel["segredos_removidos"]:
            con.execute("DELETE FROM configuracoes WHERE chave = ?", (chave,))
        # 4) a conta do usuário
        con.execute("INSERT INTO usuarios_locais (email, senha_hash, is_admin, role, ativo) VALUES (?, ?, 1, 'admin', 1)", (email, hash_password(senha)))
        novo_id = str(con.execute("SELECT id FROM usuarios_locais WHERE email = ?", (email,)).fetchone()[0])
        # 5) obras e históricos de origem: só com --com-obras (viram do novo usuário)
        for t in TABELAS_DE_OBRAS:
            if t not in tabelas:
                continue
            if com_obras:
                con.execute(f'UPDATE "{t}" SET user_id = ?', (novo_id,))
                rel["removido"][t] = 0
            else:
                rel["removido"][t] = con.execute(f'DELETE FROM "{t}"').rowcount
        for t in tabelas:                                                          # o que ainda carrega identificação de usuário (só informa)
            cols = [c for c in _colunas(con, t) if c in ("email", "usuario", "user_id", "criado_por", "autor")]
            if cols and t not in TABELAS_PESSOAIS and t not in TABELAS_DE_OBRAS and t not in TABELAS_COM_AUTOR and t != "tabela_orcamento":
                n = con.execute(f'SELECT COUNT(*) FROM "{t}"').fetchone()[0]
                if n:
                    rel["colunas_com_usuario"].append(f"{t}.{'/'.join(cols)} ({n} linhas)")
        con.commit()
        con.execute("PRAGMA journal_mode=DELETE")                                  # sem arquivos -wal/-shm (o programa volta ao modo WAL sozinho)
        con.execute("VACUUM")                                                      # apaga de verdade o que foi removido (senão os dados continuam dentro do arquivo)
        con.commit()
    finally:
        con.close()
    rel["usuario_id"] = novo_id
    return rel


# ── .env ──────────────────────────────────────────────────────────────────────────────────────────
def montar_env(origem_env, email, com_nuvem=False, com_ia=False):
    """Texto do .env da entrega. Só leva da origem o que for pedido; devolve (texto, lista_do_que_foi_copiado).
    A senha do usuário NÃO entra aqui (o .env viaja junto do zip; a senha vai por outro canal). O ADMIN_PASSWORD é só uma senha aleatória de reserva: o programa só o usa se o banco
    for recriado do zero, e assim nunca cai na senha padrão conhecida."""
    from dotenv import dotenv_values
    fonte = dotenv_values(origem_env) if origem_env and Path(origem_env).is_file() else {}
    copiados = []
    linhas = ["# OrçaCore — configuração desta instalação (não compartilhe este arquivo).", f"ADMIN_EMAIL={email}", f"ADMIN_PASSWORD={senha_aleatoria(24)}", ""]
    for rotulo, grupo, pedido in (("Nuvem (Supabase)", VARS_NUVEM, com_nuvem), ("IA (chave padrão do sistema)", VARS_IA, com_ia)):
        if not pedido:
            continue
        itens = [(k, fonte.get(k)) for k in grupo if fonte.get(k)]
        if itens:
            linhas.append(f"# {rotulo}")
            for k, v in itens:
                linhas.append(f"{k}={v}")
                copiados.append(k)
            linhas.append("")
    return "\n".join(linhas).rstrip() + "\n", copiados


# ── conferência: roda o programa de verdade sobre uma cópia ────────────────────────────────────────
_CONFERENCIA = r'''
import json, os, sys
os.environ["HEADLESS"] = "1"
os.environ["JWT_SECRET"] = "conferencia-" + "x" * 40
sys.path.insert(0, __RAIZ__)
import dotenv
from dotenv import dotenv_values
dotenv.load_dotenv = lambda *a, **k: False                  # não deixa o .env do desenvolvedor entrar na conferência
for k, v in dotenv_values(__ENV__).items():
    if not k.startswith("SUPABASE_"):                       # a conferência é local: não depende de rede
        os.environ[k] = v or ""
for k in list(os.environ):
    if k.startswith("SUPABASE_"):
        del os.environ[k]
import config
config.DB_PATH = __DB__
import database
database.DB_PATH = __DB__
database.init_db()                                          # o que o programa faz ao abrir
import app as appmod
from fastapi.testclient import TestClient
c = TestClient(appmod.app)
saida = {}
h = c.get("/api/health"); j = h.json() if h.status_code == 200 else {}
saida["health"] = [h.status_code, j.get("db_ok"), j.get("tabela_orcamento_count")]
ok = c.post("/api/auth/login", json={"email": __EMAIL__, "password": __SENHA__})
saida["login"] = ok.status_code
ruim = c.post("/api/auth/login", json={"email": __EMAIL__, "password": __SENHA__ + "x"})
saida["login_errado"] = ruim.status_code
dados = []
if ok.status_code == 200:
    tk = ok.json()["access_token"]
    r = c.get("/api/orcamento/dados", headers={"Authorization": "Bearer " + tk})
    def tamanho(resp, chave):
        if resp.status_code != 200:
            return None
        j = resp.json()
        j = j.get(chave) if isinstance(j, dict) else j
        return len(j) if isinstance(j, list) else None
    r = c.get("/api/orcamento/dados", headers={"Authorization": "Bearer " + tk})
    saida["base"] = [r.status_code, tamanho(r, "dados")]
    pr = c.get("/api/projetos", headers={"Authorization": "Bearer " + tk})
    saida["projetos"] = [pr.status_code, tamanho(pr, "projetos")]
print("RESULTADO=" + json.dumps(saida))
'''


def conferir_pacote(pasta_dados, email, senha):
    """Abre o programa sobre uma CÓPIA dos dados e confere saúde, login e leitura da base. Devolve (ok, [mensagens])."""
    msgs, ok = [], True
    with tempfile.TemporaryDirectory(prefix="orcacore_conferencia_") as tmp:
        tmp = Path(tmp)
        shutil.copy2(Path(pasta_dados) / "banco_resumo.db", tmp / "banco_resumo.db")
        shutil.copy2(Path(pasta_dados) / ".env", tmp / ".env")
        codigo = _CONFERENCIA.replace("__RAIZ__", repr(str(RAIZ))).replace("__ENV__", repr(str(tmp / ".env"))).replace("__DB__", repr(str(tmp / "banco_resumo.db"))) \
            .replace("__EMAIL__", repr(email)).replace("__SENHA__", repr(senha))
        r = subprocess.run([sys.executable, "-c", codigo], capture_output=True, text=True, timeout=180, cwd=str(tmp))
        linha = next((x for x in r.stdout.splitlines() if x.startswith("RESULTADO=")), None)
        if r.returncode != 0 or not linha:
            return False, ["O programa não abriu sobre o pacote: " + (r.stderr.strip().splitlines() or ["(sem mensagem)"])[-1]]
        res = json.loads(linha[len("RESULTADO="):])
    saude = res.get("health") or [0, None, None]
    checagens = [
        (saude[0] == 200 and saude[1] is True, f"programa abre e o banco responde (/api/health → {saude[0]})"),
        ((saude[2] or 0) > 0, f"base de orçamento presente ({saude[2]} linhas)"),
        (res.get("login") == 200, "login com o e-mail e a senha entregues"),
        (res.get("login_errado") == 401, "senha errada é recusada"),
        ((res.get("base") or [0, 0])[0] == 200 and ((res.get("base") or [0, 0])[1] or 0) > 0, "a base chega ao usuário logado (/api/orcamento/dados)"),
        ((res.get("projetos") or [0, 0])[0] == 200 and ((res.get("projetos") or [0, 0])[1] or 0) > 0, "projetos cadastrados"),
    ]
    for passou, texto in checagens:
        msgs.append(("OK   " if passou else "FALHA ") + texto)
        ok = ok and passou
    return ok, msgs


# ── arquivos da entrega ────────────────────────────────────────────────────────────────────────────
INSTALAR_BAT = r"""@echo off
chcp 65001 >nul
setlocal
title Instalar o OrçaCore
set "DESTINO=%LOCALAPPDATA%\Programs\OrcaCore"
set "AQUI=%~dp0"
echo.
echo  Instalando o OrçaCore para o seu usuário do Windows...
echo.
tasklist /FI "IMAGENAME eq OrcaCore.exe" 2>nul | find /I "OrcaCore.exe" >nul
if not errorlevel 1 (
  echo  O OrçaCore está aberto. Feche-o e execute este arquivo de novo.
  pause
  exit /b 1
)
if not exist "%AQUI%OrcaCore-Setup.exe" (
  echo  Não achei OrcaCore-Setup.exe ao lado deste arquivo.
  pause
  exit /b 1
)
if not exist "%AQUI%dados\banco_resumo.db" (
  echo  Não achei a pasta "dados" ao lado deste arquivo. Extraia o zip INTEIRO antes de executar.
  pause
  exit /b 1
)
"%AQUI%OrcaCore-Setup.exe" /SILENT /NORESTART /SP-
if not exist "%DESTINO%\OrcaCore.exe" (
  echo  A instalação não terminou. Execute OrcaCore-Setup.exe manualmente e depois este arquivo de novo.
  pause
  exit /b 1
)
rem  o instalador pode ter aberto o programa; fecha antes de copiar os dados
taskkill /IM OrcaCore.exe /F >nul 2>nul
timeout /t 2 /nobreak >nul
if exist "%DESTINO%\banco_resumo.db" (
  copy /Y "%DESTINO%\banco_resumo.db" "%DESTINO%\banco_resumo.db.anterior" >nul
  echo  Já havia um banco nesta instalação: ficou guardado como banco_resumo.db.anterior
)
copy /Y "%AQUI%dados\banco_resumo.db" "%DESTINO%\banco_resumo.db" >nul || goto erro
copy /Y "%AQUI%dados\.env" "%DESTINO%\.env" >nul || goto erro
if exist "%AQUI%dados\skills" xcopy /E /I /Y "%AQUI%dados\skills" "%DESTINO%\skills" >nul
del /Q "%DESTINO%\banco_resumo.db-wal" "%DESTINO%\banco_resumo.db-shm" >nul 2>nul
echo.
echo  Pronto! Abrindo o OrçaCore. Entre com o e-mail e a senha que você recebeu (arquivo ACESSO).
start "" "%DESTINO%\OrcaCore.exe"
exit /b 0
:erro
echo  Não foi possível copiar os dados para %DESTINO%. Feche o OrçaCore e tente de novo.
pause
exit /b 1
"""


def texto_leia_me(nome, email_dono_nota=""):
    return f"""OrçaCore — pacote de entrega ({nome})

COMO INSTALAR (leva 1 minuto)
1. Extraia o zip INTEIRO numa pasta (não execute de dentro do zip).
2. Dê dois cliques em Instalar.bat. Ele instala o OrçaCore na sua conta do Windows (sem pedir administrador),
   copia o banco e a configuração que acompanham este pacote e abre o programa.
3. Entre com o e-mail e a senha do arquivo ACESSO (enviado por outro canal) e troque a senha quando puder.

SE O WINDOWS AVISAR «O Windows protegeu o computador»: clique em «Mais informações» → «Executar assim mesmo»
(o instalador não é assinado digitalmente).

O QUE VEM PRONTO
- A base de orçamento e as regras do projeto já estão no banco.
- Para trabalhar sem internet não é preciso mais nada. Se este pacote incluiu a configuração da nuvem ou da IA, ela já está ativa.

SE ALGO NÃO FUNCIONAR
- «Credenciais inválidas»: confira o e-mail e a senha do arquivo ACESSO (a senha diferencia maiúsculas de minúsculas).
- O programa não abre: execute OrcaCore.exe em %LOCALAPPDATA%\\Programs\\OrcaCore e mande a mensagem de erro a quem enviou o pacote.
- Para reaproveitar um banco que você já tinha: o Instalar.bat guarda o anterior como banco_resumo.db.anterior na mesma pasta.

NÃO COMPARTILHE o arquivo .env (fica em %LOCALAPPDATA%\\Programs\\OrcaCore): ele contém a sua configuração.
"""


def sha256(arquivo):
    h = hashlib.sha256()
    with open(arquivo, "rb") as f:
        for bloco in iter(lambda: f.read(1 << 20), b""):
            h.update(bloco)
    return h.hexdigest()


def montar(args):
    saida = Path(args.saida)
    carimbo = datetime.datetime.now().strftime("%Y%m%d-%H%M")
    nome = f"OrcaCore-entrega-{carimbo}"
    pacote = saida / nome
    senha = args.senha or senha_aleatoria()
    email = args.email.strip()
    if not re.fullmatch(r"[^@\s]+@[^@\s]+\.[^@\s]+", email):
        raise ErroPacote("Informe um e-mail válido para a conta do usuário (--email).")
    if len(senha) < 8:
        raise ErroPacote("A senha precisa de pelo menos 8 caracteres (ou deixe vazio para o script gerar uma).")
    origem_db = Path(args.banco) if args.banco else RAIZ / "banco_resumo.db"
    origem_env = Path(args.env) if args.env else RAIZ / ".env"
    if pacote.exists():
        raise ErroPacote(f"A pasta {pacote} já existe.")
    dados = pacote / "dados"
    dados.mkdir(parents=True)
    relatorio = []
    try:
        # banco
        copiar_banco_consistente(origem_db, dados / "banco_resumo.db")
        rel = higienizar_banco(dados / "banco_resumo.db", email, senha, base_de=args.base_de, com_obras=args.com_obras, com_historico=args.com_historico)
        relatorio.append(f"Banco higienizado a partir de {origem_db}: base de orçamento = {rel['base']} ({rel['linhas_base']} linhas).")
        removidos = {t: n for t, n in rel["removido"].items() if n}
        if removidos:
            relatorio.append("  Removido: " + ", ".join(f"{t} ({n})" for t, n in removidos.items()))
        if rel["segredos_removidos"]:
            relatorio.append("  Chaves salvas removidas da configuração: " + ", ".join(rel["segredos_removidos"]))
        if rel["colunas_com_usuario"]:
            relatorio.append("  ATENÇÃO — ainda há identificação de usuário em: " + "; ".join(rel["colunas_com_usuario"]))
        # .env
        texto_env, copiados = montar_env(origem_env, email, com_nuvem=args.com_nuvem, com_ia=args.com_chaves_ia)
        (dados / ".env").write_text(texto_env, encoding="utf-8")
        relatorio.append(".env limpo criado" + (f" com {', '.join(copiados)}" if copiados else " (sem nuvem e sem chaves de IA)") + ".")
        if (args.com_nuvem or args.com_chaves_ia) and not copiados:
            relatorio.append("  ATENÇÃO — pedi nuvem/IA mas o .env de origem não tinha essas variáveis: " + str(origem_env))
        # skills
        origem_skills = Path(args.skills) if args.skills else None          # só as SUAS skills (as do pacote já vão dentro do programa); aponte a pasta skills ao lado do seu OrcaCore.exe
        if origem_skills and origem_skills.is_dir() and any(origem_skills.iterdir()):
            shutil.copytree(origem_skills, dados / "skills")
            relatorio.append(f"Skills copiadas de {origem_skills}.")
        # conferência — o que garante o uso «de cara»
        ok, msgs = conferir_pacote(dados, email, senha)
        try:
            import config
            relatorio.append(f"Conferido com o código da versão {config.APP_VERSION}: o instalador entregue deve ser DESSA mesma versão (um banco mais novo que o programa pode não abrir).")
        except Exception:  # noqa: BLE001
            pass
        relatorio.append("Conferência (o programa foi aberto sobre uma cópia do pacote):")
        relatorio += ["  " + m for m in msgs]
        if not ok:
            raise ErroPacote("A conferência falhou — o pacote NÃO foi concluído:\n" + "\n".join(msgs))
        # instalador e instruções
        if args.instalador:
            inst = Path(args.instalador)
            if not inst.is_file():
                raise ErroPacote(f"Instalador não encontrado: {inst}")
            shutil.copy2(inst, pacote / "OrcaCore-Setup.exe")
        else:
            relatorio.append("ATENÇÃO — sem --instalador: coloque o OrcaCore-Setup.exe da release na pasta do pacote (o Instalar.bat o procura ali).")
        (pacote / "Instalar.bat").write_text(INSTALAR_BAT.replace("\n", "\r\n"), encoding="utf-8", newline="")
        (pacote / "LEIA-ME.txt").write_text(texto_leia_me(nome), encoding="utf-8")
        somas = [f"{sha256(p)}  {p.relative_to(pacote).as_posix()}" for p in sorted(pacote.rglob("*")) if p.is_file() and p.name != "SHA256SUMS.txt"]
        (pacote / "SHA256SUMS.txt").write_text("\n".join(somas) + "\n", encoding="utf-8")
    except BaseException:
        shutil.rmtree(pacote, ignore_errors=True)                     # não deixa meia entrega
        raise
    acesso = saida / f"ACESSO_{nome}.txt"
    acesso.write_text(f"OrçaCore — acesso\nE-mail: {email}\nSenha: {senha}\n\nMande este arquivo por um canal DIFERENTE do pacote e peça para trocar a senha no primeiro uso.\n", encoding="utf-8")
    zip_path = None
    if args.zip:
        zip_path = saida / f"{nome}.zip"
        with zipfile.ZipFile(zip_path, "w", zipfile.ZIP_DEFLATED) as z:
            for p in sorted(pacote.rglob("*")):
                if p.is_file():
                    z.write(p, Path(nome) / p.relative_to(pacote))
    return {"pacote": pacote, "acesso": acesso, "zip": zip_path, "relatorio": relatorio, "email": email}


def main(argv=None):
    ap = argparse.ArgumentParser(description="Monta o pacote de entrega do OrçaCore (instalador + banco higienizado + .env limpo).")
    ap.add_argument("--saida", required=True, help="pasta onde criar a entrega")
    ap.add_argument("--email", required=True, help="e-mail da conta de administrador do usuário que vai receber")
    ap.add_argument("--senha", help="senha dessa conta (vazio = o script gera uma)")
    ap.add_argument("--instalador", help="caminho do OrcaCore-Setup.exe (da release ou de dist/)")
    ap.add_argument("--banco", help="banco de origem (padrão: banco_resumo.db da raiz do projeto)")
    ap.add_argument("--env", help=".env de origem (padrão: .env da raiz); só é lido com --com-nuvem/--com-chaves-ia")
    ap.add_argument("--base-de", dest="base_de", help="e-mail do usuário de origem cuja base de orçamento será entregue (obrigatório se houver edições pessoais de mais de um)")
    ap.add_argument("--skills", help="pasta de skills suas a levar (a pasta skills ao lado do seu OrcaCore.exe); sem isto, só as skills que já vêm no programa")
    ap.add_argument("--com-nuvem", action="store_true", help="leva SUPABASE_URL/SUPABASE_KEY do .env")
    ap.add_argument("--com-chaves-ia", action="store_true", help="leva as chaves de IA padrão do .env (use uma chave só para esta pessoa)")
    ap.add_argument("--com-obras", action="store_true", help="leva as obras e o histórico REC do banco (passam a ser do novo usuário)")
    ap.add_argument("--com-historico", action="store_true", help="mantém os históricos de regras/ajustes (têm o nome de quem editou)")
    ap.add_argument("--zip", action="store_true", help="também gera o .zip da entrega")
    args = ap.parse_args(argv)
    try:
        r = montar(args)
    except ErroPacote as e:
        print("ERRO:", e, file=sys.stderr)
        return 2
    print("\n".join(r["relatorio"]))
    print(f"\nPacote pronto: {r['pacote']}")
    if r["zip"]:
        print(f"Zip: {r['zip']}")
    print(f"Acesso (envie por outro canal): {r['acesso']}  →  {r['email']}")
    return 0


if __name__ == "__main__":
    sys.exit(main())
