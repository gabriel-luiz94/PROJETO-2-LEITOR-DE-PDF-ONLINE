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
from database import get_connection
from models import ObraModel
from middleware.auth_middleware import get_current_user_from_state
from services import obras_arquivo

router = APIRouter(prefix="/api/obras", tags=["obras"])


from services.supabase_client import get_supabase, registrar_falha

@router.get("")
def get_obras(request: Request, projeto: str = None):
    user = get_current_user_from_state(request)
    user_id = user["user_id"]

    supabase = get_supabase()
    if supabase:
        try:
            query = supabase.table("obras").select("*").eq("user_id", user_id)
            if projeto:
                query = query.eq("projeto", projeto)
            res = query.order("data", desc=True).execute()
            if res.data is not None:
                return res.data
        except Exception as e:
            registrar_falha("obras.get_obras", e)

    conn = get_connection()
    cursor = conn.cursor()
    if projeto:
        cursor.execute(
            "SELECT id, nome, data, dados_json, projeto FROM obras WHERE (user_id = ? OR user_id IS NULL) AND projeto = ? ORDER BY data DESC",
            (user_id, projeto)
        )
    else:
        cursor.execute(
            "SELECT id, nome, data, dados_json, projeto FROM obras WHERE user_id = ? OR user_id IS NULL ORDER BY data DESC",
            (user_id,)
        )
    rows = cursor.fetchall()
    conn.close()
    return [{"id": r[0], "nome": r[1], "data": r[2], "dados_json": r[3], "projeto": r[4]} for r in rows]


def listar_leves(user_id, projeto):
    """Obras do usuário no projeto SEM `dados_json` (id, nome, data, projeto), mais recentes primeiro. Barato: serve ao índice da IA."""
    supabase = get_supabase()
    if supabase:
        try:
            res = (supabase.table("obras").select("id, nome, data, projeto").eq("user_id", user_id).eq("projeto", projeto)
                   .order("data", desc=True).execute())
            if res.data is not None:
                return res.data
        except Exception as e:
            registrar_falha("obras.listar_leves", e)
    conn = get_connection()
    rows = conn.execute(
        "SELECT id, nome, data, projeto FROM obras WHERE (user_id = ? OR user_id IS NULL) AND projeto = ? ORDER BY data DESC",
        (user_id, projeto)).fetchall()
    conn.close()
    return [{"id": r[0], "nome": r[1], "data": r[2], "projeto": r[3]} for r in rows]


def buscar_obra(user_id, obra_id):
    """Uma obra completa do usuário (com `dados_json`) ou None."""
    supabase = get_supabase()
    if supabase:
        try:
            res = supabase.table("obras").select("*").eq("id", obra_id).eq("user_id", user_id).execute()
            if res.data:
                return res.data[0]
        except Exception as e:
            registrar_falha("obras.buscar_obra", e)
    conn = get_connection()
    r = conn.execute("SELECT id, nome, data, dados_json, projeto FROM obras WHERE id = ? AND (user_id = ? OR user_id IS NULL)",
                     (obra_id, user_id)).fetchone()
    conn.close()
    return {"id": r[0], "nome": r[1], "data": r[2], "dados_json": r[3], "projeto": r[4]} if r else None


@router.get("/indice")
def indice_obras(request: Request, projeto: str):
    """Índice leve (sem dados_json) das obras do usuário no projeto — TASK-028."""
    return listar_leves(get_current_user_from_state(request)["user_id"], projeto)


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
    _gravar_obra(user_id, {"id": obra_id, "nome": nome, "data": datetime.now().strftime("%d/%m/%Y, %H:%M:%S"),
                           "dados_json": obras_arquivo.dados_json_para_gravar(pronta["dados"]), "projeto": pronta["projeto"]})
    return {"status": "success", "id": obra_id, "nome": nome, "projeto": pronta["projeto"],
            "cabos": len(pronta["dados"]["cabos"]), "outros": len(pronta["dados"]["outros"])}


@router.get("/{obra_id}/exportar")
def exportar_obra(obra_id: str, request: Request):
    """Baixa a obra do usuário como `.obra.json` (TASK-056)."""
    obra = buscar_obra(get_current_user_from_state(request)["user_id"], obra_id)
    if obra is None:
        raise HTTPException(status_code=404, detail="Obra não encontrada.")
    envelope = obras_arquivo.montar_exportacao(obra)
    nome = obras_arquivo.nome_de_arquivo(envelope["nome"])
    return JSONResponse(envelope, headers={"Content-Disposition": f"attachment; filename=\"obra.obra.json\"; filename*=UTF-8''{quote(nome)}"})


@router.get("/{obra_id}")
def get_obra(obra_id: str, request: Request):
    """Uma obra completa do usuário (o frontend valida o id citado pela IA contra isto) — TASK-028."""
    obra = buscar_obra(get_current_user_from_state(request)["user_id"], obra_id)
    if obra is None:
        raise HTTPException(status_code=404, detail="Obra não encontrada.")
    return obra


def _gravar_obra(user_id, registro: dict) -> None:
    """Grava a obra (Supabase quando disponível + SQLite). Usado por `save_obra` e pela importação."""
    supabase = get_supabase()
    if supabase:
        try:
            supabase.table("obras").upsert({
                "id": registro["id"],
                "nome": registro["nome"],
                "data": registro["data"],
                "dados_json": registro["dados_json"],
                "user_id": user_id,
                "projeto": registro["projeto"]
            }).execute()
        except Exception as e:
            registrar_falha("obras.save_obra", e)

    conn = get_connection()
    cursor = conn.cursor()
    cursor.execute("INSERT OR REPLACE INTO obras (id, nome, data, dados_json, user_id, projeto) VALUES (?, ?, ?, ?, ?, ?)",
                   (registro["id"], registro["nome"], registro["data"], registro["dados_json"], user_id, registro["projeto"]))
    conn.commit()
    conn.close()


@router.post("")
def save_obra(obra: ObraModel, request: Request):
    user = get_current_user_from_state(request)
    _gravar_obra(user["user_id"], {"id": obra.id, "nome": obra.nome, "data": obra.data, "dados_json": obra.dados_json, "projeto": obra.projeto})
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
