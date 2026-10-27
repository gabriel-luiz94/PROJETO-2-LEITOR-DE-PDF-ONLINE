"""
routers/vinculacao_config.py — Tabela de Cabos, Tabela de Estruturas e scripts da vinculação cabo<->estrutura por projeto (TASK-084).

Nuvem primeiro, fallback local, erro de sincronização visível (mesmo padrão de routers/regras_vinculacao.py). Sem configuração salva:
 * Pernambuco ("081") nasce com o ponto de partida de data/vinculacao_pernambuco_seed.json;
 * os demais projetos MIGRAM as regras antigas (regras_vinculacao: tipo da estrutura, qtd_cabos, compatibilidade) para a Tabela de Estruturas.
A configuração só é gravada quando alguém salva — a migração é feita na leitura, sem efeito colateral.
"""
import json

from fastapi import APIRouter, HTTPException
from pydantic import BaseModel

from config import logger, VINCULACAO_PE_SEED_PATH
from database import get_connection
from routers.regras_vinculacao import _get_regras as _get_regras_legadas
from services import vinculacao_regras as vr
from services.supabase_client import get_supabase

router = APIRouter(prefix="/api/vinculacao-config", tags=["vinculacao-config"])

PROJETO_PERNAMBUCO = "081"


class ConfigPayload(BaseModel):
    projeto_codigo: str
    config: dict


def _seed_pernambuco() -> dict:
    with open(VINCULACAO_PE_SEED_PATH, "r", encoding="utf-8") as f:
        return json.load(f)


def _config_salva(projeto_codigo: str):
    supabase = get_supabase()
    if supabase:
        try:
            res = supabase.table("vinculacao_config").select("config_json").eq("projeto_codigo", projeto_codigo).execute()
            if res.data:
                cfg = json.loads(res.data[0]["config_json"])
                if isinstance(cfg, dict) and (cfg.get("cabos") or cfg.get("estruturas") or cfg.get("scripts")):
                    return cfg
        except Exception as e:
            logger.warning(f"Falha ao buscar vinculacao_config no Supabase: {e}")
    conn = get_connection()
    try:
        row = conn.execute("SELECT config_json FROM vinculacao_config WHERE projeto_codigo = ?", (projeto_codigo,)).fetchone()
    finally:
        conn.close()
    if row:
        try:
            cfg = json.loads(row[0])
            if isinstance(cfg, dict):
                return cfg
        except Exception:
            pass
    return None


def obter_config(projeto_codigo: str) -> dict:
    """Configuração efetiva do projeto, normalizada (salva > ponto de partida de Pernambuco > migração das regras antigas)."""
    return obter_config_com_origem(projeto_codigo)[0]


def obter_config_com_origem(projeto_codigo: str):
    salva = _config_salva(projeto_codigo)
    if salva is not None:
        return vr.normalizar_config(salva), "salva"
    if projeto_codigo == PROJETO_PERNAMBUCO:
        try:
            return vr.normalizar_config(_seed_pernambuco()), "pernambuco"
        except (OSError, json.JSONDecodeError) as e:
            logger.warning(f"Seed da vinculação de Pernambuco ilegível: {e}")
    legado = _get_regras_legadas(projeto_codigo)
    if legado:
        return vr.config_de_legado(legado), "migrada"
    return vr.config_vazia(), "vazia"


def salvar_config(projeto_codigo: str, cfg: dict) -> dict:
    erros = vr.validar_config(cfg)
    if erros:
        raise HTTPException(status_code=400, detail={"erros": erros})
    valor = json.dumps(vr.normalizar_config(cfg), ensure_ascii=False)
    conn = get_connection()
    try:
        conn.execute("INSERT OR REPLACE INTO vinculacao_config (projeto_codigo, config_json) VALUES (?, ?)", (projeto_codigo, valor))
        conn.commit()
    finally:
        conn.close()
    supabase = get_supabase()
    if supabase:
        try:
            supabase.table("vinculacao_config").upsert({"projeto_codigo": projeto_codigo, "config_json": valor}).execute()
        except Exception as e:
            logger.warning(f"Erro ao salvar vinculacao_config no Supabase: {e}")
            raise HTTPException(status_code=500, detail=f"Configuração salva localmente, mas falhou ao sincronizar com a nuvem (Supabase): {e}")
    return {"status": "success"}


@router.get("")
def get_vinculacao_config(projeto_codigo: str = "DEFAULT"):
    cfg, origem = obter_config_com_origem(projeto_codigo)
    return {"config": cfg, "origem": origem}


@router.post("")
def post_vinculacao_config(payload: ConfigPayload):
    return salvar_config(payload.projeto_codigo, payload.config)
