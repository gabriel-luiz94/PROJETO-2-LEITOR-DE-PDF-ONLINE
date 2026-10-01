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
    projeto: str = "229"


class RegraModel(BaseModel):
    conteudo: str
    projeto_codigo: str = "229"


class RecModel(BaseModel):
    numero_obra: str
    dados_json: str
    projeto: str = "229"


class ChatRequest(BaseModel):
    prompt: str
    table_context: str = ""
    history: List[Dict[str, Any]]
    provider: str = "gemini"          # "gemini" | "openai"
    openai_base_url: str = ""         # ex: http://localhost:11434/v1 (Ollama) ou https://openrouter.ai/api/v1
    projeto_codigo: str = "229"


class OrcamentoRequest(BaseModel):
    cabos: List[Dict[str, Any]]
    outros: List[Dict[str, Any]]
    projeto: Optional[str] = None


class PayloadCalculo(BaseModel):
    cabos: List[Dict[str, Any]] = []
    outros: List[Dict[str, Any]] = []


class ValidacaoPlanilhasRequest(BaseModel):
    cabos: List[Dict[str, Any]] = []
    outros: List[Dict[str, Any]] = []
    projeto: Optional[str] = None
    # Código do projeto (ex.: "229") para escolher as regras de domínio; sem ele, usa o padrão DEFAULT.
    projeto_codigo: Optional[str] = None
    incluir_dominio: bool = True
    # Payload de cálculo (Totalizadora, após as regras de conversão). Se enviado, também checa a base técnica.
    payload_calculo: Optional[PayloadCalculo] = None


class ValidacaoIARequest(BaseModel):
    cabos: List[Dict[str, Any]] = []
    outros: List[Dict[str, Any]] = []
    projeto_codigo: Optional[str] = None
    # Achados das camadas 1 e 2, para a IA não repeti-los
    achados_previos: List[Dict[str, Any]] = []
    prompt_id: str = "validar-planilhas"


class CorrecaoIARequest(BaseModel):
    cabos: List[Dict[str, Any]] = []
    outros: List[Dict[str, Any]] = []
    projeto_codigo: Optional[str] = None
    # Achados a corrigir (linha_id, regra_id, mensagem e, se houver, sugestao). Só as linhas citadas vão à IA.
    achados: List[Dict[str, Any]] = []
    prompt_id: str = "corrigir-planilhas"


class AjustesIARequest(BaseModel):
    """Pedido de ajustes (ações estruturadas) à IA — TASK-025. `cabos`/`outros` = tabelas COMPLETAS (o diff é calculado nelas)."""
    cabos: List[Dict[str, Any]] = []
    outros: List[Dict[str, Any]] = []
    projeto_codigo: Optional[str] = None
    achados: List[Dict[str, Any]] = []
    prompt_id: str = "ajustar-planilhas"


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
    projeto: str = "229"
