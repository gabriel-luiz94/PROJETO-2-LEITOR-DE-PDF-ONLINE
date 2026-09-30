"""
routers/validacao_prompts.py — Prompts de validação das planilhas (TASK-012, ADR-004 camada 3).

Um texto (cabeçalho + corpo) por (projeto, prompt). "DEFAULT" vale para todo projeto sem versão
própria; a semente vem de data/validacoes/*.md (database.py:_seed_prompts_validacao). Mesmo padrão de
routers/regras_leitor.py: nuvem primeiro com fallback local, validação antes de salvar, histórico com
reversão, escrita só para admin e erro de sincronização visível ao usuário (nunca engolido).

Não registrar o conteúdo dos prompts em log.
"""
from fastapi import APIRouter, Depends, HTTPException, Request
from pydantic import BaseModel

from config import VALIDACOES_SEED_DIR, logger
from database import get_connection
from middleware.auth_middleware import require_role
from services.prompts_validacao import ler_sementes, resumo_meta, validar_prompt
from services.supabase_client import get_supabase

router = APIRouter(prefix="/api/validacao/prompts", tags=["validacao-prompts"])

DEFAULT = "DEFAULT"


class SalvarPromptPayload(BaseModel):
    projeto_codigo: str = DEFAULT
    conteudo: str


class ProjetoPayload(BaseModel):
    projeto_codigo: str = DEFAULT


class ReverterPayload(BaseModel):
    projeto_codigo: str = DEFAULT
    historico_id: int


def _extrair_email(request: Request) -> str | None:
    user = getattr(request.state, "user", None)
    return user.get("email") if user else None


# ═══════════════════════════════════════════════════════════════════════════
# LEITURA
# ═══════════════════════════════════════════════════════════════════════════
def _buscar_um(prompt_id: str, projeto_codigo: str):
    supabase = get_supabase()
    if supabase:
        try:
            res = (supabase.table("prompts_validacao").select("conteudo")
                   .eq("projeto_codigo", projeto_codigo).eq("prompt_id", prompt_id).execute())
            if res.data:
                return res.data[0]["conteudo"]
        except Exception as e:
            logger.warning(f"Falha ao buscar prompts_validacao no Supabase: {e}")

    conn = get_connection()
    row = conn.execute(
        "SELECT conteudo FROM prompts_validacao WHERE projeto_codigo = ? AND prompt_id = ?",
        (projeto_codigo, prompt_id)
    ).fetchone()
    conn.close()
    return row[0] if row else None


def buscar_prompt(prompt_id: str, projeto_codigo: str = DEFAULT):
    """(conteudo, projeto_de_origem) — versão do projeto ou, na falta, a DEFAULT. None se não existir.

    Função pública: a TASK-014 usa para montar o prompt a enviar à IA.
    """
    for codigo in dict.fromkeys([projeto_codigo or DEFAULT, DEFAULT]):
        conteudo = _buscar_um(prompt_id, codigo)
        if conteudo is not None:
            return conteudo, codigo
    return None


def _ids_conhecidos() -> list:
    conn = get_connection()
    rows = conn.execute("SELECT DISTINCT prompt_id FROM prompts_validacao").fetchall()
    conn.close()
    return [r[0] for r in rows]


def _item(prompt_id: str, projeto_codigo: str):
    achado = buscar_prompt(prompt_id, projeto_codigo)
    if achado is None:
        return None
    conteudo, origem = achado
    return {
        "prompt_id": prompt_id,
        "projeto_codigo": projeto_codigo,
        "versao_de": origem,  # DEFAULT quando o projeto não tem versão própria
        "personalizado": origem != DEFAULT,
        "meta": resumo_meta(conteudo),
        "conteudo": conteudo,
    }


# ═══════════════════════════════════════════════════════════════════════════
# ESCRITA
# ═══════════════════════════════════════════════════════════════════════════
def _salvar(prompt_id: str, projeto_codigo: str, conteudo: str, usuario_email: str | None):
    if prompt_id not in _ids_conhecidos():
        raise HTTPException(status_code=404, detail=f"Prompt '{prompt_id}' não existe.")
    erros = validar_prompt(conteudo, prompt_id)
    if erros:
        raise HTTPException(status_code=400, detail={"erros": erros})

    conn = get_connection()
    cursor = conn.cursor()
    cursor.execute(
        "SELECT conteudo FROM prompts_validacao WHERE projeto_codigo = ? AND prompt_id = ?",
        (projeto_codigo, prompt_id)
    )
    atual = cursor.fetchone()
    # Preserva a versão atual no histórico ANTES de sobrescrever.
    if atual:
        cursor.execute(
            "INSERT INTO prompts_validacao_historico (projeto_codigo, prompt_id, conteudo, criado_por) VALUES (?, ?, ?, ?)",
            (projeto_codigo, prompt_id, atual[0], usuario_email)
        )
    cursor.execute(
        "INSERT OR REPLACE INTO prompts_validacao (projeto_codigo, prompt_id, conteudo) VALUES (?, ?, ?)",
        (projeto_codigo, prompt_id, conteudo)
    )
    conn.commit()
    conn.close()

    supabase = get_supabase()
    if supabase:
        try:
            if atual:
                supabase.table("prompts_validacao_historico").insert({
                    "projeto_codigo": projeto_codigo, "prompt_id": prompt_id,
                    "conteudo": atual[0], "criado_por": usuario_email,
                }).execute()
            supabase.table("prompts_validacao").upsert({
                "projeto_codigo": projeto_codigo, "prompt_id": prompt_id, "conteudo": conteudo,
            }).execute()
        except Exception as e:
            logger.warning(f"Erro ao salvar prompts_validacao no Supabase: {e}")
            raise HTTPException(
                status_code=500,
                detail=f"Prompt salvo localmente, mas falhou ao sincronizar com a nuvem (Supabase): {e}"
            )
    return {"status": "success"}


def _listar_historico(prompt_id: str, projeto_codigo: str):
    supabase = get_supabase()
    if supabase:
        try:
            res = (supabase.table("prompts_validacao_historico")
                   .select("id, conteudo, criado_em, criado_por")
                   .eq("projeto_codigo", projeto_codigo).eq("prompt_id", prompt_id)
                   .order("criado_em", desc=True).execute())
            if res.data:
                return res.data
        except Exception as e:
            logger.warning(f"Falha ao buscar histórico de prompts_validacao no Supabase: {e}")

    conn = get_connection()
    rows = conn.execute(
        "SELECT id, conteudo, criado_em, criado_por FROM prompts_validacao_historico "
        "WHERE projeto_codigo = ? AND prompt_id = ? ORDER BY criado_em DESC, id DESC",
        (projeto_codigo, prompt_id)
    ).fetchall()
    conn.close()
    return [{"id": r[0], "conteudo": r[1], "criado_em": r[2], "criado_por": r[3]} for r in rows]


# ═══════════════════════════════════════════════════════════════════════════
# ROTAS
# ═══════════════════════════════════════════════════════════════════════════
@router.get("")
def listar(projeto_codigo: str = DEFAULT):
    itens = [i for i in (_item(pid, projeto_codigo) for pid in _ids_conhecidos()) if i]
    itens.sort(key=lambda i: (float(i["meta"].get("ordem") or 0), i["prompt_id"]))
    return {"prompts": itens}


@router.get("/{prompt_id}")
def obter(prompt_id: str, projeto_codigo: str = DEFAULT):
    item = _item(prompt_id, projeto_codigo)
    if not item:
        raise HTTPException(status_code=404, detail=f"Prompt '{prompt_id}' não existe.")
    return item


@router.post("/{prompt_id}", dependencies=[Depends(require_role("admin"))])
def salvar(prompt_id: str, payload: SalvarPromptPayload, request: Request):
    return _salvar(prompt_id, payload.projeto_codigo, payload.conteudo, _extrair_email(request))


@router.get("/{prompt_id}/historico", dependencies=[Depends(require_role("admin"))])
def historico(prompt_id: str, projeto_codigo: str = DEFAULT):
    return {"historico": _listar_historico(prompt_id, projeto_codigo)}


@router.post("/{prompt_id}/reverter", dependencies=[Depends(require_role("admin"))])
def reverter(prompt_id: str, payload: ReverterPayload, request: Request):
    versao = next((h for h in _listar_historico(prompt_id, payload.projeto_codigo)
                   if h["id"] == payload.historico_id), None)
    if versao is None:
        raise HTTPException(status_code=404, detail="Versão de histórico não encontrada para este projeto.")
    return _salvar(prompt_id, payload.projeto_codigo, versao["conteudo"], _extrair_email(request))


@router.post("/{prompt_id}/restaurar-semente", dependencies=[Depends(require_role("admin"))])
def restaurar_semente(prompt_id: str, payload: ProjetoPayload, request: Request):
    sementes, _ = ler_sementes(VALIDACOES_SEED_DIR)
    if prompt_id not in sementes:
        raise HTTPException(status_code=404, detail=f"Não há semente para '{prompt_id}'.")
    return _salvar(prompt_id, payload.projeto_codigo, sementes[prompt_id], _extrair_email(request))
