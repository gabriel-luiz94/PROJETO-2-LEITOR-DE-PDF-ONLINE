"""
routers/obras.py — Rotas para gerenciamento de obras.
"""
import json
import secrets
import time
from datetime import datetime
from urllib.parse import quote

from fastapi import APIRouter, HTTPException, Request
from fastapi.responses import JSONResponse
from pydantic import BaseModel
from database import get_connection
from models import ObraModel
from middleware.auth_middleware import get_current_user_from_state
from services import obras_arquivo

router = APIRouter(prefix="/api/obras", tags=["obras"])


from services.supabase_client import get_supabase, registrar_falha

def _uid(user):
    return (user or {}).get("user_id") or None


def _e_admin(user) -> bool:
    return (user or {}).get("role") == "admin"


def _nome_curto(email) -> str:
    """Só a parte antes do `@` (TASK-057): é tudo o que se guarda e se mostra do dono."""
    return (str(email or "").split("@")[0]).strip()[:60]


def _pode_ver(linha: dict, user, projeto) -> bool:
    """Regra única de acesso (TASK-057): dono; ou pública do MESMO projeto selecionado; ou sem dono, só para administrador."""
    uid = _uid(user)
    if not uid:
        return False
    dono = linha.get("user_id") or None
    if dono == uid:
        return True
    if dono is None:
        return _e_admin(user)
    return bool(linha.get("publica")) and bool(projeto) and str(linha.get("projeto")) == str(projeto)


def _apresentar(linha: dict, user) -> dict:
    """Linha do banco -> resposta da API. Nunca expõe `user_id`/e-mail de outros; só o nome curto do dono."""
    uid = _uid(user)
    dono = linha.get("user_id") or None
    saida = {k: linha[k] for k in ("id", "nome", "data", "projeto", "dados_json") if k in linha}
    saida["publica"] = bool(linha.get("publica"))
    saida["minha"] = dono is not None and dono == uid
    saida["sem_dono"] = dono is None
    if saida["minha"]:
        saida["dono"] = linha.get("dono_nome") or _nome_curto((user or {}).get("email")) or "você"
    elif dono is None:
        saida["dono"] = ""
    else:
        saida["dono"] = linha.get("dono_nome") or "outro usuário"
    return saida


_COLS_LEVES = "id, nome, data, projeto, user_id, publica, dono_nome"
_COLS_LEVES_ANTIGAS = "id, nome, data, projeto, user_id"   # nuvem ainda sem as colunas novas (SQL da TASK-057 não rodado)


def _nuvem(sb, completo, montar):
    cols = "*" if completo else _COLS_LEVES
    try:
        return montar(sb.table("obras").select(cols)).execute().data or []
    except Exception:
        if completo:
            raise
        return montar(sb.table("obras").select(_COLS_LEVES_ANTIGAS)).execute().data or []


def _ler_linhas(user, projeto=None, completo=False, obra_id=None) -> list:
    """Linhas que `user` pode ver (dele; públicas de outros do `projeto`; sem dono se admin). Cada linha ainda traz `user_id`."""
    uid = _uid(user)
    if not uid:
        return []
    admin = _e_admin(user)

    def base(q):
        if projeto:
            q = q.eq("projeto", projeto)
        if obra_id:
            q = q.eq("id", obra_id)
        return q

    sb = get_supabase()
    linhas = None
    if sb:
        try:
            linhas = list(_nuvem(sb, completo, lambda q: base(q).eq("user_id", uid)))
            if projeto:
                try:
                    linhas += _nuvem(sb, completo, lambda q: base(q).eq("publica", True))
                except Exception as e:      # coluna `publica` ainda não existe na nuvem: só as próprias
                    registrar_falha("obras.publicas", e)
            if admin:
                try:
                    linhas += _nuvem(sb, completo, lambda q: base(q).is_("user_id", "null"))
                except Exception as e:
                    registrar_falha("obras.sem_dono", e)
        except Exception as e:
            registrar_falha("obras.ler", e)
            linhas = None
    if linhas is None:
        cols = "id, nome, data, projeto, user_id, publica, dono_nome" + (", dados_json" if completo else "")
        sql = f"SELECT {cols} FROM obras WHERE (user_id = ? OR (publica = 1 AND projeto = ?) OR (? AND (user_id IS NULL OR user_id = '')))"
        args = [uid, projeto, 1 if admin else 0]
        if projeto:
            sql += " AND projeto = ?"
            args.append(projeto)
        if obra_id:
            sql += " AND id = ?"
            args.append(obra_id)
        conn = get_connection()
        try:
            rows = conn.execute(sql + " ORDER BY data DESC", args).fetchall()
        finally:
            conn.close()
        nomes = cols.split(", ")
        linhas = [dict(zip(nomes, r)) for r in rows]
    vistos, saida = set(), []
    for l in linhas:
        if l["id"] in vistos or not _pode_ver(l, user, projeto):
            continue
        vistos.add(l["id"])
        saida.append(l)
    saida.sort(key=lambda l: l.get("data") or "", reverse=True)
    return saida


@router.get("")
def get_obras(request: Request, projeto: str = None, visibilidade: str = "todas", origem: str = "todas"):
    """Obras visíveis ao usuário (TASK-057): as dele + as públicas de outros no `projeto`. Filtros: visibilidade
    (todas|particulares|publicas) e origem (todas|minhas|outros). Sem `projeto` só vêm as dele."""
    user = get_current_user_from_state(request)
    itens = [_apresentar(l, user) for l in _ler_linhas(user, projeto, completo=True)]
    if visibilidade == "publicas":
        itens = [o for o in itens if o["publica"]]
    elif visibilidade == "particulares":
        itens = [o for o in itens if not o["publica"]]
    if origem == "minhas":
        itens = [o for o in itens if o["minha"]]
    elif origem == "outros":
        itens = [o for o in itens if not o["minha"]]
    return itens


def listar_leves(user_id, projeto):
    """Obras PRÓPRIAS do usuário no projeto SEM `dados_json` (id, nome, data, projeto), mais recentes primeiro."""
    return [{k: l[k] for k in ("id", "nome", "data", "projeto")}
            for l in _ler_linhas({"user_id": user_id}, projeto) if l.get("user_id") == user_id]


def listar_visiveis(user, projeto):
    """Índice leve para a IA (TASK-057): as próprias primeiro, depois as públicas de outros (nunca particulares alheias)."""
    itens = [_apresentar(l, user) for l in _ler_linhas(user, projeto) if l.get("user_id")]
    return [o for o in itens if o["minha"]] + [o for o in itens if not o["minha"]]


def buscar_obra(user, obra_id, projeto=None):
    """Uma obra completa que `user` pode ver (dele; ou pública do `projeto`) ou None. `user` = dict com `user_id`/`role`/`email`
    (uma string é aceita como `user_id` por compatibilidade)."""
    if isinstance(user, str):
        user = {"user_id": user}
    achadas = _ler_linhas(user, projeto, completo=True, obra_id=obra_id)
    return _apresentar(achadas[0], user) if achadas else None


@router.get("/indice")
def indice_obras(request: Request, projeto: str):
    """Índice leve (sem dados_json) das obras visíveis ao usuário no projeto — TASK-028/057."""
    return listar_visiveis(get_current_user_from_state(request), projeto)


def _projeto_existe(codigo: str) -> bool:
    conn = get_connection()
    try:
        return conn.execute("SELECT 1 FROM projetos WHERE codigo = ?", (codigo,)).fetchone() is not None
    finally:
        conn.close()


def _nome_livre(user_id, projeto, nome: str) -> str:
    """Nome que não repete outra obra do mesmo usuário no projeto: acrescenta ' (importada)', ' (importada 2)'…"""
    existentes = {o["nome"] for o in listar_leves(user_id, projeto)}
    if nome not in existentes:
        return nome
    candidato, n = f"{nome} (importada)", 2
    while candidato in existentes:
        candidato = f"{nome} (importada {n})"
        n += 1
    return candidato


@router.post("/importar")
async def importar_obra(request: Request, projeto: str = None):
    """Importa um arquivo `.obra.json` como obra NOVA e PARTICULAR do usuário que importa (TASK-056).

    Só aceita obra do MESMO projeto selecionado na tela (`?projeto=`). O corpo é o JSON do arquivo (até 5 MB)."""
    user = get_current_user_from_state(request)
    corpo = await request.body()
    if len(corpo) > obras_arquivo.LIMITE_BYTES:
        raise HTTPException(status_code=413, detail={"erros": [f"Arquivo grande demais (máx. {obras_arquivo.LIMITE_BYTES // (1024 * 1024)} MB)."]})
    try:
        carga = json.loads(corpo.decode("utf-8"))
    except (UnicodeDecodeError, ValueError):
        raise HTTPException(status_code=400, detail={"erros": ["O arquivo não é um JSON válido."]})
    try:
        pronta = obras_arquivo.validar_importacao(carga, projeto)
    except obras_arquivo.ErroArquivoObra as e:
        raise HTTPException(status_code=400, detail={"erros": e.mensagens})
    if not _projeto_existe(pronta["projeto"]):
        raise HTTPException(status_code=400, detail={"erros": [f"O projeto {pronta['projeto']} não está cadastrado."]})
    user_id = user["user_id"]
    obra_id = f"obra_{int(time.time() * 1000)}_{secrets.token_hex(2)}"
    nome = _nome_livre(user_id, pronta["projeto"], pronta["nome"])
    gravar_obra(user, {"id": obra_id, "nome": nome, "data": datetime.now().strftime("%d/%m/%Y, %H:%M:%S"),
                       "dados_json": obras_arquivo.dados_json_para_gravar(pronta["dados"]), "projeto": pronta["projeto"], "publica": False})
    return {"status": "success", "id": obra_id, "nome": nome, "projeto": pronta["projeto"],
            "cabos": len(pronta["dados"]["cabos"]), "outros": len(pronta["dados"]["outros"])}


@router.get("/{obra_id}/exportar")
def exportar_obra(obra_id: str, request: Request, projeto: str = None):
    """Baixa a obra do usuário (ou pública de outro no `projeto`) como `.obra.json` (TASK-056)."""
    obra = buscar_obra(get_current_user_from_state(request), obra_id, projeto)
    if obra is None:
        raise HTTPException(status_code=404, detail="Obra não encontrada.")
    envelope = obras_arquivo.montar_exportacao(obra)
    nome = obras_arquivo.nome_de_arquivo(envelope["nome"])
    return JSONResponse(envelope, headers={"Content-Disposition": f"attachment; filename=\"obra.obra.json\"; filename*=UTF-8''{quote(nome)}"})


@router.get("/{obra_id}")
def get_obra(obra_id: str, request: Request, projeto: str = None):
    """Uma obra completa que o usuário pode ver: dele, ou pública de outro no `projeto` (o frontend valida o id citado
    pela IA contra isto) — TASK-028/057."""
    obra = buscar_obra(get_current_user_from_state(request), obra_id, projeto)
    if obra is None:
        raise HTTPException(status_code=404, detail="Obra não encontrada.")
    return obra


def _existente(obra_id):
    """(existe, user_id, publica) da obra `obra_id` em qualquer dono — para proteger contra sobrescrita (TASK-057)."""
    conn = get_connection()
    try:
        r = conn.execute("SELECT user_id, publica FROM obras WHERE id = ?", (obra_id,)).fetchone()
    finally:
        conn.close()
    if r:
        return True, (r[0] or None), bool(r[1])
    sb = get_supabase()
    if sb:
        try:
            d = sb.table("obras").select("id, user_id").eq("id", obra_id).execute().data
            if d:
                return True, (d[0].get("user_id") or None), False
        except Exception as e:
            registrar_falha("obras.existente", e)
    return False, None, False


def gravar_obra(user, registro: dict) -> None:
    """Grava a obra do `user` (Supabase quando disponível + SQLite). `registro['publica']` (bool, padrão falso).
    Usado por `save_obra`, pela importação e pelo modo autônomo."""
    uid = _uid(user)
    publica = bool(registro.get("publica"))
    dono_nome = _nome_curto((user or {}).get("email")) or None
    base = {"id": registro["id"], "nome": registro["nome"], "data": registro["data"], "dados_json": registro["dados_json"],
            "user_id": uid, "projeto": registro["projeto"]}
    supabase = get_supabase()
    if supabase:
        try:
            try:
                supabase.table("obras").upsert({**base, "publica": publica, "dono_nome": dono_nome}).execute()
            except Exception:   # nuvem ainda sem as colunas novas: grava como antes (a obra segue particular)
                supabase.table("obras").upsert(base).execute()
        except Exception as e:
            registrar_falha("obras.save_obra", e)

    conn = get_connection()
    cursor = conn.cursor()
    cursor.execute("INSERT OR REPLACE INTO obras (id, nome, data, dados_json, user_id, projeto, publica, dono_nome) VALUES (?, ?, ?, ?, ?, ?, ?, ?)",
                   (registro["id"], registro["nome"], registro["data"], registro["dados_json"], uid, registro["projeto"], int(publica), dono_nome))
    conn.commit()
    conn.close()


@router.post("")
def save_obra(obra: ObraModel, request: Request):
    user = get_current_user_from_state(request)
    existe, dono, publica_atual = _existente(obra.id)
    if existe and dono != _uid(user):
        raise HTTPException(status_code=403, detail="Esta obra pertence a outro usuário. Salve como uma nova obra.")
    publica = obra.publica if obra.publica is not None else (publica_atual if existe else False)
    gravar_obra(user, {"id": obra.id, "nome": obra.nome, "data": obra.data, "dados_json": obra.dados_json,
                       "projeto": obra.projeto, "publica": publica})
    return {"status": "success", "publica": bool(publica)}


class VisibilidadeModel(BaseModel):
    publica: bool


@router.put("/{obra_id}/visibilidade")
def alterar_visibilidade(obra_id: str, corpo: VisibilidadeModel, request: Request):
    """Torna a obra pública (visível no projeto) ou particular. Só o dono (TASK-057)."""
    user = get_current_user_from_state(request)
    existe, dono, _ = _existente(obra_id)
    if not existe or dono is None:
        raise HTTPException(status_code=404, detail="Obra não encontrada.")
    if dono != _uid(user):
        # quem não enxerga a obra recebe 404 (não revela que existe); quem a enxerga (pública) recebe 403
        visivel = any(l["id"] == obra_id for l in _ler_linhas(user, _projeto_da(obra_id)))
        raise HTTPException(status_code=403 if visivel else 404, detail="Só o dono da obra altera a visibilidade." if visivel else "Obra não encontrada.")
    dono_nome = _nome_curto(user.get("email")) or None
    aviso = None
    sb = get_supabase()
    if sb:
        try:
            sb.table("obras").update({"publica": corpo.publica, "dono_nome": dono_nome}).eq("id", obra_id).eq("user_id", dono).execute()
        except Exception as e:
            registrar_falha("obras.visibilidade", e)
            aviso = "A nuvem ainda não tem as colunas de visibilidade; rode o SQL da TASK-057 no Supabase."
    conn = get_connection()
    conn.execute("UPDATE obras SET publica = ?, dono_nome = COALESCE(?, dono_nome) WHERE id = ? AND user_id = ?",
                 (int(corpo.publica), dono_nome, obra_id, dono))
    conn.commit()
    conn.close()
    return {"status": "success", "publica": corpo.publica, **({"aviso": aviso} if aviso else {})}


def _projeto_da(obra_id):
    conn = get_connection()
    try:
        r = conn.execute("SELECT projeto FROM obras WHERE id = ?", (obra_id,)).fetchone()
    finally:
        conn.close()
    return r[0] if r else None


@router.post("/{obra_id}/assumir")
def assumir_obra(obra_id: str, request: Request):
    """Administrador assume uma obra antiga SEM dono: vira obra PARTICULAR dele (TASK-057)."""
    user = get_current_user_from_state(request)
    if not _e_admin(user):
        raise HTTPException(status_code=403, detail="Só administradores podem assumir obras sem dono.")
    existe, dono, _ = _existente(obra_id)
    if not existe or dono is not None:
        raise HTTPException(status_code=404, detail="Obra sem dono não encontrada.")
    dono_nome = _nome_curto(user.get("email")) or None
    sb = get_supabase()
    if sb:
        try:
            q = sb.table("obras").update({"user_id": _uid(user), "publica": False, "dono_nome": dono_nome}).eq("id", obra_id).is_("user_id", "null")
            q.execute()
        except Exception as e:
            registrar_falha("obras.assumir", e)
    conn = get_connection()
    conn.execute("UPDATE obras SET user_id = ?, publica = 0, dono_nome = ? WHERE id = ? AND (user_id IS NULL OR user_id = '')",
                 (_uid(user), dono_nome, obra_id))
    conn.commit()
    conn.close()
    return {"status": "success"}


@router.delete("/{obra_id}")
def delete_obra(obra_id: str, request: Request):
    user = get_current_user_from_state(request)
    user_id = user["user_id"]

    supabase = get_supabase()
    if supabase:
        try:
            supabase.table("obras").delete().eq("id", obra_id).eq("user_id", user_id).execute()
        except Exception as e:
            registrar_falha("obras.delete_obra", e)

    conn = get_connection()
    cursor = conn.cursor()
    cursor.execute("DELETE FROM obras WHERE id = ? AND user_id = ?", (obra_id, user_id))
    conn.commit()
    conn.close()
    return {"status": "success"}
