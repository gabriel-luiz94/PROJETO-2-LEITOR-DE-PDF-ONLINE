"""
models.py — Modelos Pydantic usados em toda a aplicação.
"""
from pydantic import BaseModel
from typing import List, Dict, Any, Optional


class ObraModel(BaseModel):
    id: str
    nome: str
    data: str
    dados_json: str


class RegraModel(BaseModel):
    conteudo: str


class RecModel(BaseModel):
    numero_obra: str
    dados_json: str


class ChatRequest(BaseModel):
    prompt: str
    table_context: str = ""
    history: List[Dict[str, Any]]
    provider: str = "gemini"          # "gemini" | "openai"
    openai_base_url: str = ""         # ex: http://localhost:11434/v1 (Ollama) ou https://openrouter.ai/api/v1


class OrcamentoRequest(BaseModel):
    cabos: List[Dict[str, Any]]
    outros: List[Dict[str, Any]]
    projeto: Optional[str] = None


class SalvarOrcamentoRequest(BaseModel):
    dados: List[Dict[str, Any]]


class ProjetoRequest(BaseModel):
    nome: str
    codigo: str


class DetalhesRequest(BaseModel):
    codigos: List[str]


class RecSaveRequest(BaseModel):
    numero_obra: str
    dados: List[Dict[str, Any]]
