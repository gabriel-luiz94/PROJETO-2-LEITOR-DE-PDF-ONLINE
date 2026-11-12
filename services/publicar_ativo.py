"""
services/publicar_ativo.py — Publicar (e desfazer) UM ativo na master da Base de Orçamento, local + nuvem (TASK-145).

Em vez de reenviar a tabela inteira, só as linhas do ativo (no projeto) são trocadas. Antes de gravar: a validação é a da importação da base (`rascunho_ativo.avaliar`), e
uma ASSINATURA do que o usuário viu detecta edição simultânea por outro admin. A nuvem é gravada PRIMEIRO (insere as novas, depois apaga as antigas; se algo falhar nada
fica pela metade e o local não muda); depois o banco local e a cópia pessoal do admin. Guarda UM passo anterior por (projeto, ativo): o que havia antes da última publicação
(vazio = o ativo não existia), usado por `desfazer` (que troca o atual pelo anterior, então um segundo "desfazer" refaz).
"""
import hashlib
import json

from config import logger
from database import get_connection, get_row_connection
from services.ativo_parametros import analisar_celula_base, nome_base
from services.rascunho_ativo import avaliar
from services.supabase_client import get_supabase
from services.sync_service import espelhar_master_no_pessoal

CAMPOS = ("ativo", "desc_ativo", "componente", "projeto", "mdo", "codigo", "desc_codigo", "fator_i", "fator_r", "filtro", "origem", "lv", "arred", "formula_i", "formula_r")
TABELA = "tabela_orcamento_master"


class ErroPublicacao(Exception):
    def __init__(self, mensagem, status=400, extra=None):
        super().__init__(mensagem)
        self.mensagem, self.status, self.extra = mensagem, status, extra or {}


def _u(v) -> str:
    return str(v if v is not None else "").strip().upper()


def _escopo(r) -> str:
    return _u(r.get("projeto"))


def ids_do_ativo(linhas, nome, projeto):
    """`linhas` na ordem da tabela (dicts com 'id'): os ids das linhas do ativo `nome` no projeto (linha com ATIVO em branco herda o ativo de cima)."""
    alvo, proj, atual, ids = _u(nome), _u(projeto), None, []
    for r in linhas:
        cel = str(r.get("ativo") or "").strip()
        if cel:
            atual = nome_base(cel)
        if atual == alvo and _escopo(r) == proj:
            ids.append(r["id"])
    return ids


def _canon(r) -> tuple:
    def n(v):
        try:
            return f"{float(v or 0):.6g}"
        except (TypeError, ValueError):
            return "0"
    return (str(r.get("ativo") or "").strip(), str(r.get("desc_ativo") or "").strip(), str(r.get("componente") or "").strip(), _u(r.get("projeto")), str(r.get("mdo") or "").strip(),
            str(r.get("codigo") or "").strip(), str(r.get("desc_codigo") or "").strip(), n(r.get("fator_i")), n(r.get("fator_r")), str(r.get("filtro") or "").strip(),
            (str(r.get("lv") or "A").strip() or "A"), (str(r.get("arred") or "NORMAL").strip() or "NORMAL"), str(r.get("formula_i") or "").strip(), str(r.get("formula_r") or "").strip())


def assinatura(linhas) -> str:
    """Impressão digital estável das linhas do ativo (sem ids): muda se qualquer campo mudar."""
    return hashlib.sha1(json.dumps([_canon(r) for r in linhas], ensure_ascii=False).encode("utf-8")).hexdigest()[:16]


def _so_campos(r) -> dict:
    return {c: r.get(c, "" if c not in ("fator_i", "fator_r") else 0) for c in CAMPOS}


# ───────────────────────── leitura ─────────────────────────
def _locais():
    conn = get_row_connection()
    try:
        return [dict(x) for x in conn.execute(f"SELECT * FROM {TABELA} ORDER BY id").fetchall()]
    finally:
        conn.close()


def _da_nuvem(sb, projeto):
    linhas, ini = [], 0
    while True:
        res = sb.table(TABELA).select("*").ilike("projeto", _u(projeto)).order("id").range(ini, ini + 999).execute()
        lote = res.data or []
        linhas.extend(lote)
        if len(lote) < 1000:
            break
        ini += 1000
    return [r for r in linhas if _escopo(r) == _u(projeto)]


def linhas_do_ativo(projeto, nome, sb=None):
    """(linhas do ativo no projeto, na fonte da verdade: a nuvem se houver, senão o banco local)."""
    base = _da_nuvem(sb, projeto) if sb else [r for r in _locais() if _escopo(r) == _u(projeto)]
    ids = set(ids_do_ativo(base, nome, projeto))
    return [r for r in base if r["id"] in ids]


def _anterior(projeto, nome):
    conn = get_row_connection()
    try:
        r = conn.execute("SELECT linhas_json, usuario, criado_em FROM ativo_versoes WHERE projeto = ? AND ativo = ?", (_u(projeto), _u(nome))).fetchone()
        return (json.loads(r["linhas_json"]), r["usuario"], r["criado_em"]) if r else None
    finally:
        conn.close()


def _ultimo_editor(projeto, nome):
    conn = get_row_connection()
    try:
        r = conn.execute("SELECT email, created_at FROM audit_log WHERE record_id = ? AND action IN ('PUBLICAR_ATIVO', 'DESFAZER_ATIVO') ORDER BY id DESC LIMIT 1", (f"{_u(projeto)}|{_u(nome)}",)).fetchone()
        return (r["email"], r["created_at"]) if r else (None, None)
    finally:
        conn.close()


def versao(projeto, nome, sb=None) -> dict:
    atuais = linhas_do_ativo(projeto, nome, sb)
    ant = _anterior(projeto, nome)
    return {"assinatura": assinatura(atuais), "linhas": [_so_campos(r) for r in atuais], "tem_anterior": ant is not None, "anterior_em": ant[2] if ant else None, "anterior_por": ant[1] if ant else None}


# ───────────────────────── prévia ─────────────────────────
def diff(antigas, novas) -> dict:
    """Linhas novas / alteradas / removidas de um ativo. Chave: (ATIVO, código, MDO), com ocorrências repetidas casadas na ordem."""
    chave = lambda r: (str(r.get("ativo") or "").strip(), str(r.get("codigo") or "").strip(), _u(r.get("mdo")) if str(r.get("codigo") or "").strip() else "")   # linha sem código (declaração/chamada): o MDO não identifica
    fila = {}
    for r in antigas:
        fila.setdefault(chave(r), []).append(r)
    campos = ("desc_ativo", "componente", "mdo", "desc_codigo", "fator_i", "fator_r", "formula_i", "formula_r", "filtro", "lv", "arred")
    novas_l, alteradas, iguais = [], [], 0
    for r in novas:
        velha = (fila.get(chave(r)) or [None]).pop(0) if fila.get(chave(r)) else None
        if velha is None:
            novas_l.append(_so_campos(r))
            continue
        ca, cb = _canon(velha), _canon(r)
        if ca == cb:
            iguais += 1
        else:
            mudou = [c for c in campos if _campo(velha, c) != _campo(r, c)]
            alteradas.append({"antes": _so_campos(velha), "depois": _so_campos(r), "campos": mudou})
    removidas = [_so_campos(r) for lista in fila.values() for r in lista]
    return {"novas": novas_l, "alteradas": alteradas, "removidas": removidas, "iguais": iguais}


def _campo(r, c):
    v = r.get(c)
    if c in ("fator_i", "fator_r"):
        try:
            return f"{float(v or 0):.6g}"
        except (TypeError, ValueError):
            return "0"
    return str(v if v is not None else "").strip()


def chamadores(base, nome) -> list:
    """Ativos que chamam `nome` (composto `:=` ou grupo que chama): [{ativo, qtd}]."""
    alvo, saida = _u(nome), {}
    for r in base:
        cel = str(r.get("ativo") or "").strip()
        if ":=" not in cel:
            continue
        a = analisar_celula_base(cel)
        for it in (a.get("itens") or []):
            if nome_base(it["ativo"]) == alvo and _u(a.get("base")) != alvo:
                saida[a["base"]] = saida.get(a["base"], 0) + 1
    return [{"ativo": k, "referencias": v} for k, v in sorted(saida.items())]


def previa(projeto, nome, linhas, sb=None) -> dict:
    base = _locais()
    ava = avaliar(base, linhas, projeto)
    atuais = linhas_do_ativo(projeto, nome, sb)
    novas = ava["linhas"] if not ava["erros"] else []
    d = diff(atuais, novas) if novas else {"novas": [], "alteradas": [], "removidas": [], "iguais": 0}
    ant = _anterior(projeto, nome)
    return {"erros": ava["erros"], "avisos": ava["avisos"], "diff": d, "chamadores": chamadores(base, nome), "total_base": len(base), "assinatura": assinatura(atuais),
            "existe": bool(atuais), "tem_anterior": ant is not None}


# ───────────────────────── gravação ─────────────────────────
def _inserir_local(cur, linhas):
    for r in linhas:
        v = _so_campos(r)
        cur.execute(f"INSERT INTO {TABELA} ({', '.join(CAMPOS)}) VALUES ({', '.join('?' * len(CAMPOS))})", [v[c] for c in CAMPOS])


def _aplicar(projeto, nome, novas, sb):
    """Troca as linhas do ativo no projeto: nuvem primeiro (sem deixar nada pela metade), depois o banco local."""
    novas = [dict(_so_campos(r), projeto=r.get("projeto") or projeto) for r in novas]
    nuvem = False
    if sb:
        try:
            na_nuvem = _da_nuvem(sb, projeto)
            ids_velhos = ids_do_ativo(na_nuvem, nome, projeto)
            novos_ids = []
            for i in range(0, len(novas), 500):
                res = sb.table(TABELA).insert(novas[i:i + 500]).execute()
                novos_ids += [x.get("id") for x in (res.data or []) if x.get("id") is not None]
            try:
                for i in range(0, len(ids_velhos), 200):
                    sb.table(TABELA).delete().in_("id", ids_velhos[i:i + 200]).execute()
            except Exception:
                try:                                                   # desfaz as inserções: a nuvem volta ao que era
                    for i in range(0, len(novos_ids), 200):
                        sb.table(TABELA).delete().in_("id", novos_ids[i:i + 200]).execute()
                except Exception as e2:  # noqa: BLE001
                    logger.error(f"Não foi possível desfazer as inserções na nuvem: {e2}")
                raise
            nuvem = True
        except Exception as e:  # noqa: BLE001
            logger.warning(f"Publicar ativo: falha na nuvem: {e}")
            raise ErroPublicacao(f"Não foi possível gravar na nuvem ({e}). Nada foi alterado.", 502)
    conn = get_row_connection()
    try:
        cur = conn.cursor()
        cur.execute("BEGIN")
        locais = [dict(x) for x in conn.execute(f"SELECT * FROM {TABELA} ORDER BY id").fetchall()]
        ids = ids_do_ativo(locais, nome, projeto)
        for i in range(0, len(ids), 500):
            cur.execute(f"DELETE FROM {TABELA} WHERE id IN ({','.join('?' * len(ids[i:i + 500]))})", ids[i:i + 500])
        _inserir_local(cur, novas)
        conn.commit()
    except Exception:
        conn.rollback()
        raise
    finally:
        conn.close()
    return nuvem


def _guardar_anterior(projeto, nome, linhas, usuario):
    conn = get_connection()
    try:
        conn.execute("INSERT OR REPLACE INTO ativo_versoes (projeto, ativo, linhas_json, usuario, criado_em) VALUES (?, ?, ?, ?, datetime('now'))",
                     (_u(projeto), _u(nome), json.dumps([_so_campos(r) for r in linhas], ensure_ascii=False), usuario))
        conn.commit()
    finally:
        conn.close()


def publicar(admin: dict, projeto: str, nome: str, linhas: list, esperada=None, forcar=False, sb="auto") -> dict:
    """Publica as `linhas` como a definição do ativo `nome` no `projeto`. Levanta ErroPublicacao (400 validação / 409 conflito / 502 nuvem)."""
    sb = get_supabase() if sb == "auto" else sb
    projeto, nome = str(projeto or "").strip(), str(nome or "").strip()
    if not projeto or not nome:
        raise ErroPublicacao("Escolha um projeto e informe o nome do ativo.")
    atuais = linhas_do_ativo(projeto, nome, sb)
    if esperada is not None and esperada != assinatura(atuais) and not forcar:
        quem, quando = _ultimo_editor(projeto, nome)
        raise ErroPublicacao("Este ativo mudou desde que você o abriu" + (f" (alterado por {quem} em {quando})" if quem else "") + ".", 409,
                             {"conflito": True, "assinatura_atual": assinatura(atuais), "linhas_atuais": [_so_campos(r) for r in atuais]})
    ava = avaliar(_locais(), linhas, projeto)
    if ava["erros"]:
        raise ErroPublicacao("A publicação foi recusada: " + " | ".join(ava["erros"][:6]), 400, {"erros": ava["erros"]})
    if not ava["linhas"]:
        raise ErroPublicacao("Não há linhas para publicar.")
    if any(nome_base(l["ativo"]) != _u(nome) for l in ava["linhas"] if str(l.get("ativo") or "").strip()):
        raise ErroPublicacao("Todas as linhas precisam ser do ativo " + nome + ".")
    nuvem = _aplicar(projeto, nome, ava["linhas"], sb)
    _guardar_anterior(projeto, nome, atuais, admin.get("email"))
    espelhar_master_no_pessoal(admin.get("user_id"), projeto)
    novas = linhas_do_ativo(projeto, nome, sb)
    return {"status": "ok", "linhas": len(ava["linhas"]), "removidas": len(atuais), "nuvem": nuvem, "assinatura": assinatura(novas), "avisos": ava["avisos"]}


def desfazer(admin: dict, projeto: str, nome: str, sb="auto") -> dict:
    """Volta o ativo ao que havia antes da última publicação (e guarda o atual como o novo "anterior": desfazer de novo refaz)."""
    sb = get_supabase() if sb == "auto" else sb
    ant = _anterior(projeto, nome)
    if ant is None:
        raise ErroPublicacao("Não há versão anterior guardada para este ativo.", 404)
    atuais = linhas_do_ativo(projeto, nome, sb)
    nuvem = _aplicar(projeto, nome, ant[0], sb)
    _guardar_anterior(projeto, nome, atuais, admin.get("email"))
    espelhar_master_no_pessoal(admin.get("user_id"), projeto)
    return {"status": "ok", "linhas": len(ant[0]), "removidas": len(atuais), "nuvem": nuvem, "assinatura": assinatura(linhas_do_ativo(projeto, nome, sb))}
