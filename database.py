"""
database.py — Inicialização do banco de dados SQLite e acesso centralizado.
"""
import sqlite3
import csv
import io
import os
from config import DB_PATH, SEED_CSV_PATH, logger


def get_connection() -> sqlite3.Connection:
    """Retorna uma conexão SQLite configurada. Caller é responsável por fechar."""
    return sqlite3.connect(DB_PATH)


def get_row_connection() -> sqlite3.Connection:
    """Retorna uma conexão SQLite com row_factory = sqlite3.Row."""
    conn = sqlite3.connect(DB_PATH)
    conn.row_factory = sqlite3.Row
    return conn


def init_db():
    """Cria todas as tabelas e faz seed automático se necessário."""
    conn = get_connection()
    cursor = conn.cursor()

    # Tabela Obras
    cursor.execute('''
        CREATE TABLE IF NOT EXISTS obras (
            id TEXT PRIMARY KEY,
            nome TEXT NOT NULL,
            data TEXT NOT NULL,
            dados_json TEXT NOT NULL
        )
    ''')

    # Tabela Regras de Aprendizado
    cursor.execute('''
        CREATE TABLE IF NOT EXISTS regras (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            conteudo TEXT NOT NULL,
            embedding TEXT
        )
    ''')
    try:
        cursor.execute("ALTER TABLE regras ADD COLUMN embedding TEXT")
    except sqlite3.OperationalError:
        pass  # Column already exists

    # Tabela Configurações
    cursor.execute('''
        CREATE TABLE IF NOT EXISTS configuracoes (
            chave TEXT PRIMARY KEY,
            valor TEXT NOT NULL
        )
    ''')

    # Tabela Orçamento (Customizado do Usuário)
    cursor.execute('''
        CREATE TABLE IF NOT EXISTS tabela_orcamento (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            user_id TEXT,
            ativo TEXT,
            desc_ativo TEXT,
            componente TEXT,
            projeto TEXT,
            mdo TEXT,
            codigo TEXT,
            desc_codigo TEXT,
            fator_i REAL,
            fator_r REAL,
            filtro TEXT
        )
    ''')
    try:
        cursor.execute("ALTER TABLE tabela_orcamento ADD COLUMN filtro TEXT")
    except sqlite3.OperationalError:
        pass
    try:
        cursor.execute("ALTER TABLE tabela_orcamento ADD COLUMN user_id TEXT")
    except sqlite3.OperationalError:
        pass

    # Tabela Orçamento Mestre (Admin)
    cursor.execute('''
        CREATE TABLE IF NOT EXISTS tabela_orcamento_master (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            ativo TEXT,
            desc_ativo TEXT,
            componente TEXT,
            projeto TEXT,
            mdo TEXT,
            codigo TEXT,
            desc_codigo TEXT,
            fator_i REAL,
            fator_r REAL,
            filtro TEXT
        )
    ''')

    # Tabela de Sessão Local (Login offline)
    cursor.execute('''
        CREATE TABLE IF NOT EXISTS sessoes (
            user_id TEXT PRIMARY KEY,
            email TEXT NOT NULL,
            access_token TEXT NOT NULL,
            refresh_token TEXT,
            is_admin BOOLEAN DEFAULT 0,
            ultimo_login TEXT
        )
    ''')

    # Tabela Histórico REC
    cursor.execute('''
        CREATE TABLE IF NOT EXISTS historico_rec (
            numero_obra TEXT PRIMARY KEY,
            dados_json TEXT NOT NULL,
            data_criacao TEXT NOT NULL
        )
    ''')

    # Tabela Projetos
    cursor.execute('''
        CREATE TABLE IF NOT EXISTS projetos (
            nome TEXT PRIMARY KEY,
            codigo TEXT NOT NULL
        )
    ''')
    cursor.execute("INSERT OR IGNORE INTO projetos (nome, codigo) VALUES ('PARAIBA', '027')")
    cursor.execute("INSERT OR IGNORE INTO projetos (nome, codigo) VALUES ('RONDONIA', '229')")

    conn.commit()

    # ── Seed automático da tabela_orcamento ──
    try:
        cursor.execute("SELECT COUNT(*) FROM tabela_orcamento")
        count = cursor.fetchone()[0]
        if count == 0 and os.path.exists(SEED_CSV_PATH):
            logger.info(f"tabela_orcamento vazia -> importando seed de {SEED_CSV_PATH}")
            with open(SEED_CSV_PATH, "r", encoding="utf-8-sig") as f:
                text = f.read()
                reader = csv.DictReader(io.StringIO(text), delimiter=";")
                if not reader.fieldnames or len(reader.fieldnames) < 2:
                    f.seek(0)
                    reader = csv.DictReader(io.StringIO(f.read()), delimiter=",")
                for row in reader:
                    row_upper = {k.strip().upper(): v for k, v in row.items() if k}
                    ativo = row_upper.get("ATIVO", "")
                    codigo = row_upper.get("CODIGO", "")
                    if not ativo and not codigo:
                        continue
                    try:
                        fator_i = float(row_upper.get("FATOR I", "0").replace(",", "."))
                    except ValueError:
                        fator_i = 0.0
                    try:
                        fator_r = float(row_upper.get("FATOR R", "0").replace(",", "."))
                    except ValueError:
                        fator_r = 0.0
                    cursor.execute('''
                        INSERT INTO tabela_orcamento (ativo, desc_ativo, componente, projeto, mdo, codigo, desc_codigo, fator_i, fator_r, filtro)
                        VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?)
                    ''', (
                        ativo,
                        row_upper.get("DESC ATIVO", ""),
                        row_upper.get("COMPONENTE", ""),
                        row_upper.get("PROJETO", ""),
                        row_upper.get("MDO", ""),
                        row_upper.get("CODIGO", ""),
                        row_upper.get("DESC CODIGO", ""),
                        fator_i,
                        fator_r,
                        row_upper.get("FILTRO", "")
                    ))
            conn.commit()
            cursor.execute("SELECT COUNT(*) FROM tabela_orcamento")
            logger.info(f"Seed concluido: {cursor.fetchone()[0]} linhas importadas.")
    except Exception as e:
        logger.error(f"Erro no seed automatico: {e}")

    conn.close()
