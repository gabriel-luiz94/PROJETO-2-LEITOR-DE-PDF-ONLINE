"""
services/copia_local.py — Cópia local do banco da nuvem (TASK-121). SÓ no modo desktop.

A cada início do programa (em segundo plano) e sempre que a nuvem volta depois de uma queda, o banco SQLite local é posto igual ao da nuvem
(Supabase), para o modo offline usar exatamente os mesmos dados. Cobre os dados de uso (projetos, obras, regras, vinculação, modais, contextos,
ajustes, prompts, REC, regras de conversão) e as contas de usuário (para o login offline). A base de orçamento (master) segue em `sync_service`.

Segurança do trabalho feito offline (nada some): para as tabelas "mescladas", guardamos uma "foto" (hash) de cada linha na última cópia.
Na cópia seguinte comparamos local × nuvem × foto:
  • só o local mudou (ou a linha foi criada/apagada offline) → ENVIA para a nuvem e mantém;
  • só a nuvem mudou → aplica no local;
  • os dois mudaram diferente (conflito) → a nuvem vence, e a versão local fica guardada em `conflitos_locais`.
Na primeira cópia de uma tabela (sem foto) a nuvem vence e o que existia só aqui é guardado em `conflitos_locais`.
Tabelas "espelho" (histórico de versões, `regras` da IA) são só espelhadas da nuvem (o que sobrar local é guardado em `conflitos_locais`).
Ficam de fora (só existem no computador): sessões, logs, a cópia pessoal da base e `configuracoes` (tem a chave da API).
"""
import hashlib
import json
import threading
import time

from config import APP_MODE, logger
from database import get_connection, get_row_connection

PAGINA = 1000
LOTE_ENVIO = 100

# (tabela_nuvem, tabela_local, [colunas da chave], modo)  — modo: "mesclar" | "espelho"
TABELAS = [
    ("projetos", "projetos", ["nome"], "mesclar"),
    ("obras", "obras", ["id"], "mesclar"),
    ("historico_rec", "historico_rec", ["numero_obra"], "mesclar"),
    ("atalhos_modais", "atalhos_modais", ["projeto_codigo", "modal"], "mesclar"),
    ("contextos", "contextos", ["projeto_codigo", "contexto"], "mesclar"),
    ("ajustes_contextos", "ajustes_contextos", ["projeto_codigo", "contexto"], "mesclar"),
    ("ajustes_planilhas", "ajustes_planilhas", ["projeto_codigo"], "mesclar"),
    ("prompts_validacao", "prompts_validacao", ["projeto_codigo", "prompt_id"], "mesclar"),
    ("regras_dominio", "regras_dominio", ["projeto_codigo"], "mesclar"),
    ("regras_leitor_classificacao", "regras_leitor_classificacao", ["projeto_codigo"], "mesclar"),
    ("regras_leitor_processamento", "regras_leitor_processamento", ["projeto_codigo"], "mesclar"),
    ("regras_vinculacao", "regras_vinculacao", ["projeto_codigo"], "mesclar"),
    ("vinculacao_config", "vinculacao_config", ["projeto_codigo"], "mesclar"),
    ("regras_conversao", "configuracoes", ["projeto_codigo"], "mesclar"),          # local: configuracoes[regras_conversao_<projeto>] (adaptador abaixo)
    ("ajustes_contextos_historico", "ajustes_contextos_historico", ["id"], "espelho"),
    ("ajustes_planilhas_historico", "ajustes_planilhas_historico", ["id"], "espelho"),
    ("prompts_validacao_historico", "prompts_validacao_historico", ["id"], "espelho"),
    ("regras_dominio_historico", "regras_dominio_historico", ["id"], "espelho"),
    ("regras_leitor_classificacao_historico", "regras_leitor_classificacao_historico", ["id"], "espelho"),
    ("regras_leitor_processamento_historico", "regras_leitor_processamento_historico", ["id"], "espelho"),
    ("regras", "regras", ["id"], "espelho"),
]

_lock = threading.Lock()
_andamento = {"em_andamento": False, "ultima": None, "erro": None}


def habilitada() -> bool:
    return APP_MODE == "desktop"


# ─────────────────────────── utilidades ───────────────────────────
def _norm(v):
    """Valor canônico para comparar linha local × nuvem (o Supabase devolve bool/inteiros/decimais de outro jeito que o SQLite)."""
    if v is None:
        return None
    if isinstance(v, bool):
        return int(v)
    if isinstance(v, (int, float)):
        f = float(v)
        return int(f) if f.is_integer() else f
    return str(v)


def _hash(linha: dict, colunas) -> str:
    return hashlib.sha1(json.dumps([_norm(linha.get(c)) for c in colunas], ensure_ascii=False, sort_keys=False).encode("utf-8")).hexdigest()


def _chave(linha: dict, chaves) -> str:
    return json.dumps([_norm(linha.get(k)) for k in chaves], ensure_ascii=False)


def _meta_get(chave, padrao=None):
    conn = get_row_connection()
    try:
        r = conn.execute("SELECT valor FROM copia_local_meta WHERE chave = ?", (chave,)).fetchone()
        return r[0] if r else padrao
    except Exception:  # noqa: BLE001 — banco ainda sem a tabela (primeira execução): sem cópia anterior
        return padrao
    finally:
        conn.close()


def _meta_set(chave, valor):
    conn = get_connection()
    try:
        conn.execute("INSERT OR REPLACE INTO copia_local_meta (chave, valor) VALUES (?, ?)", (chave, valor))
        conn.commit()
    finally:
        conn.close()


def estado() -> dict:
    return {"em_andamento": _andamento["em_andamento"], "ultima": _meta_get("ultima_copia"), "erro": _andamento["erro"], "pendencias": _meta_get("pendencias_envio", "0")}


def _colunas_locais(conn, tabela):
    return [r[1] for r in conn.execute(f"PRAGMA table_info({tabela})").fetchall()]


# ─────────────────────────── adaptadores de tabela ───────────────────────────
class _Local:
    """Acesso às linhas locais de uma tabela (genérico, ou o adaptador de regras_conversao ⇄ configuracoes)."""
    def __init__(self, tabela, chaves, adaptador=None):
        self.tabela, self.chaves, self.adaptador = tabela, chaves, adaptador

    def colunas(self, conn):
        return ["projeto_codigo", "regras_json"] if self.adaptador else _colunas_locais(conn, self.tabela)

    def ler(self, conn):
        if self.adaptador:
            return [{"projeto_codigo": c[len("regras_conversao_"):], "regras_json": v}
                    for c, v in conn.execute("SELECT chave, valor FROM configuracoes WHERE chave LIKE 'regras_conversao_%'").fetchall()]
        conn.row_factory = None
        cols = _colunas_locais(conn, self.tabela)
        return [dict(zip(cols, r)) for r in conn.execute(f"SELECT * FROM {self.tabela}").fetchall()]

    def gravar(self, conn, linha, colunas):
        if self.adaptador:
            conn.execute("INSERT OR REPLACE INTO configuracoes (chave, valor) VALUES (?, ?)", ("regras_conversao_" + str(linha["projeto_codigo"]), linha.get("regras_json")))
            return
        cols = [c for c in colunas if c in linha]
        marcas = ",".join("?" for _ in cols)
        atualiza = ",".join(f"{c}=excluded.{c}" for c in cols if c not in self.chaves)
        sql = f"INSERT INTO {self.tabela} ({','.join(cols)}) VALUES ({marcas})"
        sql += f" ON CONFLICT({','.join(self.chaves)}) DO UPDATE SET {atualiza}" if atualiza else f" ON CONFLICT({','.join(self.chaves)}) DO NOTHING"
        conn.execute(sql, [linha.get(c) for c in cols])

    def apagar(self, conn, valores):
        if self.adaptador:
            conn.execute("DELETE FROM configuracoes WHERE chave = ?", ("regras_conversao_" + str(valores["projeto_codigo"]),))
            return
        conn.execute(f"DELETE FROM {self.tabela} WHERE " + " AND ".join(f"{k} = ?" for k in self.chaves), [valores[k] for k in self.chaves])


# ─────────────────────────── nuvem ───────────────────────────
def _ler_nuvem(supabase, tabela, chaves):
    """Todas as linhas da tabela na nuvem, ordenadas pela chave (paginação estável). Levanta exceção se a tabela não existe/falha."""
    linhas, ini = [], 0
    while True:
        q = supabase.table(tabela).select("*")
        for k in chaves:
            q = q.order(k)
        dados = q.range(ini, ini + PAGINA - 1).execute().data or []
        linhas.extend(dados)
        if len(dados) < PAGINA:
            return linhas
        ini += PAGINA


def _enviar(supabase, tabela, chaves, linhas, colunas_nuvem):
    """Upsert em lotes; devolve o conjunto de chaves (str) enviadas com sucesso."""
    ok = set()
    for i in range(0, len(linhas), LOTE_ENVIO):
        lote = linhas[i:i + LOTE_ENVIO]
        corpo = [{c: l.get(c) for c in colunas_nuvem if c in l and c != "updated_at"} for l in lote]
        try:
            supabase.table(tabela).upsert(corpo, on_conflict=",".join(chaves)).execute()
            ok.update(_chave(l, chaves) for l in lote)
        except Exception as e:  # noqa: BLE001
            logger.warning(f"Cópia local: não consegui enviar {len(lote)} linha(s) de {tabela} para a nuvem: {e}")
    return ok


def _apagar_na_nuvem(supabase, tabela, chaves, valores):
    q = supabase.table(tabela).delete()
    for k in chaves:
        q = q.eq(k, valores[k])
    q.execute()


def _guardar_conflito(conn, tabela, chave, linha, motivo):
    conn.execute("INSERT INTO conflitos_locais (tabela, chave, dados_json, motivo) VALUES (?, ?, ?, ?)",
                 (tabela, chave, json.dumps(linha, ensure_ascii=False, default=str) if linha is not None else None, motivo))


# ─────────────────────────── uma tabela ───────────────────────────
def _sincronizar_tabela(supabase, tabela_nuvem, tabela_local, chaves, modo) -> dict:
    local = _Local(tabela_local, chaves, adaptador=(tabela_nuvem == "regras_conversao"))
    nuvem = _ler_nuvem(supabase, tabela_nuvem, chaves)
    conn = get_connection()
    resumo = {"aplicadas": 0, "enviadas": 0, "removidas": 0, "conflitos": 0, "falhas": 0}
    try:
        cols_local = local.colunas(conn)
        cols_nuvem = sorted({c for r in nuvem for c in r.keys()}) if nuvem else [c for c in cols_local if c != "updated_at"]
        comuns = [c for c in cols_local if c in cols_nuvem and c != "updated_at"]            # o que se compara (updated_at muda toda hora)
        gravaveis = [c for c in cols_local if c in cols_nuvem]
        L = {_chave(r, chaves): r for r in local.ler(conn)}
        R = {_chave(r, chaves): r for r in nuvem}
        assinatura = ",".join(comuns)
        rotulo = f"{tabela_nuvem}"
        base = {c: h for c, h in conn.execute("SELECT chave, hash FROM copia_local_base WHERE tabela = ?", (rotulo,)).fetchall()}
        primeira = modo == "espelho" or _meta_get("cols:" + rotulo) != assinatura or not base and _meta_get("base:" + rotulo) is None
        nova_base = {}
        a_enviar, a_apagar_nuvem = [], []

        for k in set(L) | set(R) | set(base):
            l, r, b = L.get(k), R.get(k), base.get(k)
            hl = _hash(l, comuns) if l is not None else None
            hr = _hash(r, comuns) if r is not None else None
            if primeira:
                if l is not None and (r is None or hl != hr):
                    _guardar_conflito(conn, rotulo, k, l, "primeira cópia: a nuvem prevaleceu" if r is not None else "linha só local (a nuvem não a tem)")
                    resumo["conflitos"] += 1 if r is not None else 0
                acao = "nuvem"
            else:
                local_mudou, nuvem_mudou = hl != b, hr != b
                if not local_mudou:
                    acao = "nuvem"
                elif not nuvem_mudou:
                    acao = "enviar"
                elif hl == hr:
                    acao = "nuvem"
                else:
                    _guardar_conflito(conn, rotulo, k, l, "conflito: mudou aqui e na nuvem; a nuvem prevaleceu")
                    resumo["conflitos"] += 1
                    acao = "nuvem"
            if acao == "enviar":
                (a_enviar if l is not None else a_apagar_nuvem).append((k, l if l is not None else json.loads(k)))
                continue
            # a nuvem vence
            if r is not None:
                if l is None or hl != hr:
                    local.gravar(conn, r, gravaveis)
                    resumo["aplicadas"] += 1
                nova_base[k] = hr
            elif l is not None:
                local.apagar(conn, {c: l[c] for c in chaves} if not local.adaptador else l)
                resumo["removidas"] += 1

        # envio do que mudou só aqui (depois de aplicar a nuvem, para não travar a transação local)
        conn.commit()
        enviadas = _enviar(supabase, tabela_nuvem, chaves, [l for k, l in a_enviar], cols_nuvem or cols_local) if a_enviar else set()
        for k, l in a_enviar:
            if k in enviadas:
                nova_base[k] = _hash(l, comuns)
                resumo["enviadas"] += 1
            else:
                resumo["falhas"] += 1
                if base.get(k):
                    nova_base[k] = base[k]                       # segue "mudada só aqui": tenta enviar de novo na próxima cópia
        for k, valores in a_apagar_nuvem:
            try:
                _apagar_na_nuvem(supabase, tabela_nuvem, chaves, dict(zip(chaves, valores)))
                resumo["enviadas"] += 1
            except Exception as e:  # noqa: BLE001
                resumo["falhas"] += 1
                nova_base[k] = base.get(k)
                logger.warning(f"Cópia local: não consegui apagar {k} de {tabela_nuvem} na nuvem: {e}")

        # nova "foto"
        conn.execute("DELETE FROM copia_local_base WHERE tabela = ?", (rotulo,))
        conn.executemany("INSERT OR REPLACE INTO copia_local_base (tabela, chave, hash) VALUES (?, ?, ?)", [(rotulo, k, h) for k, h in nova_base.items() if h])
        conn.execute("INSERT OR REPLACE INTO copia_local_meta (chave, valor) VALUES (?, ?)", ("cols:" + rotulo, assinatura))
        conn.execute("INSERT OR REPLACE INTO copia_local_meta (chave, valor) VALUES (?, ?)", ("base:" + rotulo, "1"))
        conn.commit()
    finally:
        conn.close()
    return resumo


# ─────────────────────────── contas de usuário ───────────────────────────
def _copiar_usuarios(supabase) -> int:
    """usuarios_nuvem → usuarios_locais (login offline com a mesma senha e o MESMO user_id do online). Usuários só locais são mantidos."""
    nuvem = _ler_nuvem(supabase, "usuarios_nuvem", ["id"])
    conn = get_connection()
    try:
        for u in nuvem:
            role = u.get("role") or ("admin" if u.get("is_admin") else "operador")
            is_admin = int(bool(u.get("is_admin")) or role == "admin")
            conn.execute(
                "INSERT INTO usuarios_locais (email, senha_hash, is_admin, role, nuvem_id) VALUES (?, ?, ?, ?, ?) "
                "ON CONFLICT(email) DO UPDATE SET senha_hash=excluded.senha_hash, is_admin=excluded.is_admin, role=excluded.role, nuvem_id=excluded.nuvem_id",
                (u.get("email"), u.get("senha_hash"), is_admin, role, str(u.get("id"))))
        conn.commit()
    finally:
        conn.close()
    return len(nuvem)


# ─────────────────────────── execução ───────────────────────────
def sincronizar(motivo: str = "") -> dict | None:
    """Copia/mescla tudo (uma execução por vez; precisa da nuvem). Devolve um resumo ou None se não rodou."""
    from services import rede_adaptativa
    from services.supabase_client import get_supabase, registrar_falha
    if not habilitada() or not _lock.acquire(blocking=False):
        return None
    try:
        supabase = get_supabase()
        if not supabase or rede_adaptativa.offline():
            return None
        _andamento.update(em_andamento=True, erro=None)
        t0 = time.time()
        logger.info(f"Cópia local do banco: iniciando ({motivo or 'início'})…")
        resumo, falhas = {}, 0
        try:
            resumo["usuarios"] = _copiar_usuarios(supabase)
        except Exception as e:  # noqa: BLE001
            falhas += 1
            registrar_falha("copia_local.usuarios", e)
        for tn, tl, chaves, modo in TABELAS:
            if rede_adaptativa.offline():
                break
            try:
                resumo[tn] = _sincronizar_tabela(supabase, tn, tl, chaves, modo)
            except Exception as e:  # noqa: BLE001
                falhas += 1
                registrar_falha(f"copia_local.{tn}", e)
        try:
            from services.sync_service import sync_tabela_master
            sync_tabela_master()                                                         # base de orçamento (nuvem → local, na ordem da planilha)
        except Exception as e:  # noqa: BLE001
            falhas += 1
            logger.warning(f"Cópia local: base de orçamento não sincronizou: {e}")
        pend = sum(r.get("falhas", 0) for r in resumo.values() if isinstance(r, dict))
        _meta_set("pendencias_envio", str(pend))
        if not falhas:
            _meta_set("ultima_copia", time.strftime("%Y-%m-%d %H:%M:%S"))
        _andamento["erro"] = f"{falhas} tabela(s) com falha" if falhas else None
        logger.info(f"Cópia local do banco: concluída em {time.time() - t0:.1f}s" + (f" ({falhas} falha(s))." if falhas else "."))
        return resumo
    finally:
        _andamento["em_andamento"] = False
        _lock.release()


def iniciar_em_segundo_plano(motivo: str = "início") -> bool:
    """Dispara `sincronizar` numa thread (o programa continua abrindo/usando a nuvem enquanto copia)."""
    if not habilitada():
        return False
    threading.Thread(target=lambda: _seguro(motivo), daemon=True, name="CopiaLocal").start()
    return True


def _seguro(motivo):
    try:
        sincronizar(motivo)
    except Exception as e:  # noqa: BLE001 — a cópia nunca derruba o programa
        _andamento.update(em_andamento=False, erro=str(e))
        logger.error(f"Cópia local do banco falhou: {e}")
