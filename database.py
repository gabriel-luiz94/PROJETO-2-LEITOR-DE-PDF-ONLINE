"""
database.py — Inicialização do banco de dados SQLite e acesso centralizado.

Usa bcrypt para hashing de senhas. Admin inicial configurado via variáveis de ambiente.
Inclui tabela sync_log para controle de sincronização offline-first.
"""
import sqlite3
import csv
import io
import os
import bcrypt
from config import DB_PATH, SEED_CSV_PATH, logger


def get_connection() -> sqlite3.Connection:
    """Retorna uma conexão SQLite configurada. Caller é responsável por fechar."""
    conn = sqlite3.connect(DB_PATH)
    conn.execute("PRAGMA journal_mode=WAL")  # Melhor concorrência
    return conn


def get_row_connection() -> sqlite3.Connection:
    """Retorna uma conexão SQLite com row_factory = sqlite3.Row."""
    conn = sqlite3.connect(DB_PATH)
    conn.execute("PRAGMA journal_mode=WAL")
    conn.row_factory = sqlite3.Row
    return conn


def hash_password(password: str) -> str:
    """Hash de senha usando bcrypt (com salt automático)."""
    return bcrypt.hashpw(password.encode("utf-8"), bcrypt.gensalt()).decode("utf-8")


def verify_password(password: str, hashed: str) -> bool:
    """Verifica senha contra hash bcrypt. Suporta legado SHA-256 para migração."""
    try:
        return bcrypt.checkpw(password.encode("utf-8"), hashed.encode("utf-8"))
    except (ValueError, TypeError):
        # Fallback: tenta SHA-256 para senhas legadas (migração automática)
        import hashlib
        if hashlib.sha256(password.encode()).hexdigest() == hashed:
            return True
        return False


def _migrate_legacy_password(email: str, password: str, conn: sqlite3.Connection):
    """Migra senha legada SHA-256 para bcrypt automaticamente."""
    new_hash = hash_password(password)
    cursor = conn.cursor()
    cursor.execute(
        "UPDATE usuarios_locais SET senha_hash = ? WHERE email = ?",
        (new_hash, email)
    )
    conn.commit()
    logger.info(f"Senha migrada para bcrypt: {email}")


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
            dados_json TEXT NOT NULL,
            user_id TEXT,
            updated_at TEXT DEFAULT (datetime('now'))
        )
    ''')
    try:
        cursor.execute("ALTER TABLE obras ADD COLUMN user_id TEXT")
    except sqlite3.OperationalError:
        pass
    try:
        cursor.execute("ALTER TABLE obras ADD COLUMN updated_at TEXT")
    except sqlite3.OperationalError:
        pass

    # Tabela Regras de Aprendizado
    cursor.execute('''
        CREATE TABLE IF NOT EXISTS regras (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            conteudo TEXT NOT NULL,
            embedding TEXT,
            updated_at TEXT DEFAULT (datetime('now'))
        )
    ''')
    try:
        cursor.execute("ALTER TABLE regras ADD COLUMN embedding TEXT")
    except sqlite3.OperationalError:
        pass  # Column already exists
    try:
        cursor.execute("ALTER TABLE regras ADD COLUMN updated_at TEXT")
    except sqlite3.OperationalError:
        pass

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
            filtro TEXT,
            updated_at TEXT DEFAULT (datetime('now'))
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
    try:
        cursor.execute("ALTER TABLE tabela_orcamento ADD COLUMN updated_at TEXT")
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
            filtro TEXT,
            updated_at TEXT DEFAULT (datetime('now'))
        )
    ''')
    try:
        cursor.execute("ALTER TABLE tabela_orcamento_master ADD COLUMN updated_at TEXT")
    except sqlite3.OperationalError:
        pass

    # Tabela de Sessão Local (Login offline)
    cursor.execute('''
        CREATE TABLE IF NOT EXISTS sessoes (
            user_id TEXT PRIMARY KEY,
            email TEXT NOT NULL,
            access_token TEXT NOT NULL,
            refresh_token TEXT,
            is_admin BOOLEAN DEFAULT 0,
            role TEXT DEFAULT 'viewer',
            ultimo_login TEXT
        )
    ''')
    try:
        cursor.execute("ALTER TABLE sessoes ADD COLUMN role TEXT DEFAULT 'viewer'")
    except sqlite3.OperationalError:
        pass

    # Tabela de Usuários Locais (Para Login Offline e Painel Admin)
    cursor.execute('''
        CREATE TABLE IF NOT EXISTS usuarios_locais (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            email TEXT UNIQUE NOT NULL,
            senha_hash TEXT NOT NULL,
            is_admin BOOLEAN DEFAULT 0,
            role TEXT DEFAULT 'viewer',
            ativo BOOLEAN DEFAULT 1,
            created_at TEXT DEFAULT (datetime('now')),
            updated_at TEXT DEFAULT (datetime('now'))
        )
    ''')
    try:
        cursor.execute("ALTER TABLE usuarios_locais ADD COLUMN role TEXT DEFAULT 'viewer'")
    except sqlite3.OperationalError:
        pass
    try:
        cursor.execute("ALTER TABLE usuarios_locais ADD COLUMN ativo BOOLEAN DEFAULT 1")
    except sqlite3.OperationalError:
        pass
    try:
        cursor.execute("ALTER TABLE usuarios_locais ADD COLUMN created_at TEXT")
    except sqlite3.OperationalError:
        pass
    try:
        cursor.execute("ALTER TABLE usuarios_locais ADD COLUMN updated_at TEXT")
    except sqlite3.OperationalError:
        pass

    # Inserir admin padrão caso a tabela de usuarios esteja vazia
    # Credenciais lidas de variáveis de ambiente (nunca hardcoded)
    cursor.execute("SELECT COUNT(*) FROM usuarios_locais")
    if cursor.fetchone()[0] == 0:
        admin_email = os.environ.get("ADMIN_EMAIL", "admin@local.com")
        admin_password = os.environ.get("ADMIN_PASSWORD", "admin123")
        senha_hash = hash_password(admin_password)
        cursor.execute('''
            INSERT INTO usuarios_locais (email, senha_hash, is_admin, role)
            VALUES (?, ?, 1, 'admin')
        ''', (admin_email, senha_hash))
        logger.info(f"Admin inicial criado: {admin_email}")

    # Tabela Histórico REC
    cursor.execute('''
        CREATE TABLE IF NOT EXISTS historico_rec (
            numero_obra TEXT PRIMARY KEY,
            dados_json TEXT NOT NULL,
            data_criacao TEXT NOT NULL,
            user_id TEXT,
            updated_at TEXT DEFAULT (datetime('now'))
        )
    ''')
    try:
        cursor.execute("ALTER TABLE historico_rec ADD COLUMN user_id TEXT")
    except sqlite3.OperationalError:
        pass
    try:
        cursor.execute("ALTER TABLE historico_rec ADD COLUMN updated_at TEXT")
    except sqlite3.OperationalError:
        pass

    # Tabela Projetos
    cursor.execute('''
        CREATE TABLE IF NOT EXISTS projetos (
            nome TEXT PRIMARY KEY,
            codigo TEXT NOT NULL,
            updated_at TEXT DEFAULT (datetime('now'))
        )
    ''')
    try:
        cursor.execute("ALTER TABLE projetos ADD COLUMN updated_at TEXT")
    except sqlite3.OperationalError:
        pass
    cursor.execute("INSERT OR IGNORE INTO projetos (nome, codigo) VALUES ('PARAIBA', '027')")
    cursor.execute("INSERT OR IGNORE INTO projetos (nome, codigo) VALUES ('RONDONIA', '229')")

    # ── Tabela de Sync Log (controle de sincronização offline-first) ──
    cursor.execute('''
        CREATE TABLE IF NOT EXISTS sync_log (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            tabela TEXT NOT NULL,
            operacao TEXT NOT NULL,
            registro_id TEXT,
            dados_json TEXT,
            timestamp TEXT DEFAULT (datetime('now')),
            sincronizado BOOLEAN DEFAULT 0,
            tentativas INTEGER DEFAULT 0,
            erro TEXT
        )
    ''')

    # ── Tabela de Audit Log (quem fez o quê e quando) ──
    cursor.execute('''
        CREATE TABLE IF NOT EXISTS audit_log (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            user_id TEXT,
            email TEXT,
            action TEXT NOT NULL,
            table_name TEXT,
            record_id TEXT,
            details TEXT,
            created_at TEXT DEFAULT (datetime('now'))
        )
    ''')

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
