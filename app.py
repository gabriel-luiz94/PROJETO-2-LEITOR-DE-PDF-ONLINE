from fastapi import FastAPI, UploadFile, File, WebSocket, WebSocketDisconnect, Request, Form, HTTPException
from fastapi.responses import HTMLResponse, JSONResponse
from fastapi.staticfiles import StaticFiles
from fastapi.middleware.cors import CORSMiddleware
import pymupdf
import ezdxf
from ezdxf.colors import int2rgb, aci2rgb
from ezdxf.tools.text import plain_mtext
import tempfile
import os
import sys
import re
import json
import threading
import webbrowser
import sqlite3
import urllib.request
import csv
import io
from pydantic import BaseModel
from typing import List, Dict, Any, Optional


app = FastAPI()

app.add_middleware(
    CORSMiddleware,
    allow_origins=["*"],
    allow_credentials=True,
    allow_methods=["*"],
    allow_headers=["*"],
)

# =====================================================
# WEBSOCKET
# =====================================================

class ConnectionManager:
    def __init__(self):
        self.active_connections = []

    async def connect(self, websocket):
        await websocket.accept()
        self.active_connections.append(websocket)

    def disconnect(self, websocket):
        if websocket in self.active_connections:
            self.active_connections.remove(websocket)

    async def broadcast(self, message):
        for c in self.active_connections:
            try:
                await c.send_text(message)
            except:
                pass

manager = ConnectionManager()

@app.websocket("/ws")
async def websocket_endpoint(websocket: WebSocket):
    await manager.connect(websocket)
    try:
        while True:
            await websocket.receive_text()
    except WebSocketDisconnect:
        manager.disconnect(websocket)

# =====================================================
# BANCO DE DADOS (MEMÓRIA DE OBRAS E REGRAS)
# =====================================================
DB_PATH = os.path.join(os.path.dirname(os.path.abspath(__file__)), "banco_resumo.db")

def init_db():
    conn = sqlite3.connect(DB_PATH)
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
    cursor.execute('''
        CREATE TABLE IF NOT EXISTS configuracoes (
            chave TEXT PRIMARY KEY,
            valor TEXT NOT NULL
        )
    ''')
    cursor.execute('''
        CREATE TABLE IF NOT EXISTS tabela_orcamento (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            ativo TEXT,
            desc_ativo TEXT,
            componente TEXT,
            projeto TEXT,
            mdo TEXT,
            codigo TEXT,
            desc_codigo TEXT,
            fator_i REAL,
            fator_r REAL
        )
    ''')
    cursor.execute('''
        CREATE TABLE IF NOT EXISTS historico_rec (
            numero_obra TEXT PRIMARY KEY,
            dados_json TEXT NOT NULL,
            data_criacao TEXT NOT NULL
        )
    ''')
    conn.commit()
    conn.close()

init_db()

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

@app.get("/api/obras")
def get_obras():
    conn = sqlite3.connect(DB_PATH)
    cursor = conn.cursor()
    cursor.execute("SELECT id, nome, data, dados_json FROM obras ORDER BY data DESC")
    rows = cursor.fetchall()
    conn.close()
    return [{"id": r[0], "nome": r[1], "data": r[2], "dados_json": r[3]} for r in rows]

@app.post("/api/obras")
def save_obra(obra: ObraModel):
    conn = sqlite3.connect(DB_PATH)
    cursor = conn.cursor()
    cursor.execute("INSERT OR REPLACE INTO obras (id, nome, data, dados_json) VALUES (?, ?, ?, ?)",
                   (obra.id, obra.nome, obra.data, obra.dados_json))
    conn.commit()
    conn.close()
    return {"status": "success"}

@app.delete("/api/obras/{obra_id}")
def delete_obra(obra_id: str):
    conn = sqlite3.connect(DB_PATH)
    cursor = conn.cursor()
    cursor.execute("DELETE FROM obras WHERE id = ?", (obra_id,))
    conn.commit()
    conn.close()
    return {"status": "success"}

@app.get("/api/regras")
def get_regras():
    conn = sqlite3.connect(DB_PATH)
    cursor = conn.cursor()
    cursor.execute("SELECT id, conteudo FROM regras ORDER BY id ASC")
    rows = cursor.fetchall()
    conn.close()
    return [{"id": r[0], "conteudo": r[1]} for r in rows]

@app.post("/api/regras")
def save_regra(regra: RegraModel):
    conn = sqlite3.connect(DB_PATH)
    cursor = conn.cursor()
    
    cursor.execute("INSERT INTO regras (conteudo, embedding) VALUES (?, NULL)", (regra.conteudo,))
    conn.commit()
    conn.close()
    return {"status": "success"}

# =====================================================
# API RECS
# =====================================================
from datetime import datetime

@app.get("/api/recs")
def get_recs():
    conn = sqlite3.connect(DB_PATH)
    cursor = conn.cursor()
    cursor.execute("SELECT numero_obra, data_criacao FROM historico_rec ORDER BY data_criacao DESC")
    rows = cursor.fetchall()
    conn.close()
    return [{"numero_obra": r[0], "data_criacao": r[1]} for r in rows]

@app.get("/api/recs/{numero_obra}")
def get_rec(numero_obra: str):
    conn = sqlite3.connect(DB_PATH)
    cursor = conn.cursor()
    cursor.execute("SELECT dados_json FROM historico_rec WHERE numero_obra = ?", (numero_obra,))
    row = cursor.fetchone()
    conn.close()
    from fastapi import HTTPException
    if row:
        return {"dados_json": row[0]}
    raise HTTPException(status_code=404, detail="REC não encontrado")

@app.post("/api/recs")
def save_rec(rec: RecModel):
    conn = sqlite3.connect(DB_PATH)
    cursor = conn.cursor()
    agora = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
    cursor.execute("INSERT OR REPLACE INTO historico_rec (numero_obra, dados_json, data_criacao) VALUES (?, ?, ?)",
                   (rec.numero_obra, rec.dados_json, agora))
    conn.commit()
    conn.close()
    return {"status": "success"}

@app.delete("/api/recs/{numero_obra}")
def delete_rec(numero_obra: str):
    conn = sqlite3.connect(DB_PATH)
    cursor = conn.cursor()
    cursor.execute("DELETE FROM historico_rec WHERE numero_obra = ?", (numero_obra,))
    conn.commit()
    conn.close()
    return {"status": "success"}

# =====================================================
# API GEMINI (PROXY)
# =====================================================
class ChatRequest(BaseModel):
    prompt: str
    table_context: str = ""
    history: List[Dict[str, Any]]
    provider: str = "gemini"          # "gemini" | "openai"
    openai_base_url: str = ""         # ex: http://localhost:11434/v1 (Ollama) ou https://openrouter.ai/api/v1


@app.get("/api/gemini/models")
def get_gemini_models(request: Request):
    api_key = request.headers.get("X-Gemini-Key")
    provider = request.headers.get("X-Provider", "gemini")
    openai_base_url = request.headers.get("X-OpenAI-Base-URL", "")

    conn = sqlite3.connect("banco_resumo.db")
    cursor = conn.cursor()
    cursor.execute("SELECT valor FROM configuracoes WHERE chave = 'gemini_api_key'")
    row = cursor.fetchone()
    saved_key = row[0] if row else None
    conn.close()

    if not api_key or api_key == "SAVED_IN_BACKEND":
        if saved_key:
            api_key = saved_key
        else:
            raise HTTPException(status_code=401, detail="API Key não encontrada.")

    # Provedores OpenAI-compatíveis (Ollama, OpenRouter, LM Studio, OpenAI)
    if provider == "openai":
        try:
            from openai import OpenAI
            base = openai_base_url or "https://api.openai.com/v1"
            client = OpenAI(api_key=api_key if api_key else "ollama", base_url=base)
            models_raw = [m.id for m in client.models.list().data]
            modelos_formatados = [{"id": m, "label": m} for m in sorted(models_raw)]
            return {"models": modelos_formatados, "provider": "openai"}
        except Exception as e:
            raise HTTPException(status_code=400, detail=str(e))

    # Gemini (padrão) — usando SDK novo google.genai
    try:
        from google import genai as _genai
        _client = _genai.Client(api_key=api_key)
        preferencias_gemini = ["gemini-3.1-flash-lite", "gemini-3.6-flash", "gemini-2.5-flash", "gemini-1.5-flash"]
        # Lista modelos disponíveis para a chave
        models_page = _client.models.list()
        modelos_raw = []
        for m in models_page:
            name = m.name.split("/")[-1] if "/" in m.name else m.name
            if name.startswith("gemini"):
                modelos_raw.append(name)
        modelos_formatados = []
        for m in modelos_raw:
            label = m
            if m == "gemini-3.1-flash-lite":
                label = f"{m} ★ (Padrão)"
            elif m in preferencias_gemini:
                label = f"{m} (Reserva)"
            modelos_formatados.append({"id": m, "label": label})
        # Garante que o modelo padrão aparece sempre na lista
        ids = [m["id"] for m in modelos_formatados]
        if "gemini-3.1-flash-lite" not in ids:
            modelos_formatados.insert(0, {"id": "gemini-3.1-flash-lite", "label": "gemini-3.1-flash-lite ★ (Padrão)"})
        return {"models": modelos_formatados, "provider": "gemini"}
    except Exception as e:
        raise HTTPException(status_code=400, detail=str(e))


@app.post("/api/gemini/chat")
async def gemini_chat(req: ChatRequest, request: Request):
    import re, json, math, asyncio
    from fastapi.responses import StreamingResponse, PlainTextResponse

    # ── 1. Resolver API Key e Modelo ──
    api_key = request.headers.get("X-Gemini-Key")
    custom_model = request.headers.get("X-Gemini-Model")
    provider = req.provider or "gemini"
    openai_base_url = req.openai_base_url or ""

    conn = sqlite3.connect("banco_resumo.db")
    cursor = conn.cursor()
    cursor.execute("SELECT valor FROM configuracoes WHERE chave = 'gemini_api_key'")
    row = cursor.fetchone()
    saved_key = row[0] if row else None
    cursor.execute("SELECT valor FROM configuracoes WHERE chave = 'gemini_model'")
    row_model = cursor.fetchone()
    saved_model = row_model[0] if row_model else None
    cursor.execute("SELECT valor FROM configuracoes WHERE chave = 'ai_provider'")
    row_prov = cursor.fetchone()
    saved_provider = row_prov[0] if row_prov else "gemini"
    cursor.execute("SELECT valor FROM configuracoes WHERE chave = 'openai_base_url'")
    row_base = cursor.fetchone()
    saved_base_url = row_base[0] if row_base else ""

    if api_key and api_key != "SAVED_IN_BACKEND":
        cursor.execute("INSERT OR REPLACE INTO configuracoes (chave, valor) VALUES ('gemini_api_key', ?)", (api_key,))
        conn.commit()
    elif saved_key:
        api_key = saved_key
    else:
        conn.close()
        raise HTTPException(status_code=401, detail="API Key não fornecida ou não salva internamente.")

    if custom_model == "SAVED_IN_BACKEND":
        custom_model = saved_model
    elif custom_model is not None:
        cursor.execute("INSERT OR REPLACE INTO configuracoes (chave, valor) VALUES ('gemini_model', ?)", (custom_model,))
        conn.commit()

    if provider == "gemini" and saved_provider != "gemini":
        provider = saved_provider
    if not openai_base_url and saved_base_url:
        openai_base_url = saved_base_url
    conn.close()

    # ── 2. Fast-Path (comandos de adição simples — sem chamar IA) ──
    fast_match = re.match(
        r'^(adicionar|add|adc|coloca|inserir|remover|rem|tira|excluir)\s+([\d.,]+)\s+(.+)$',
        req.prompt.strip(), re.IGNORECASE
    )
    if fast_match:
        acao = fast_match.group(1).lower()
        qtd  = fast_match.group(2).replace(",", ".")
        ativo = fast_match.group(3).upper().strip()
        op = "R" if acao.startswith("rem") or acao in ["tira", "excluir"] else "I"
        fake_json = {"outros": [{"ativo": f"{qtd}-{ativo}", "operacao": op}]}
        return PlainTextResponse(f"*(Fast-Path)*\n```json\n{json.dumps(fake_json, indent=2)}\n```")

    # ── 3. Detectar se precisa de contexto da tabela ──
    prompt_lower = req.prompt.lower()
    palavras_contexto = [
        'analis', 'alterar', 'editar', 'excluir', 'remover', 'substituir',
        'tudo', 'todas', 'tabela', 'leia', 'leitura', 'completo', 'lista',
        'quantos', 'total', 'verifique', 'cheque', 'corrig'
    ]
    precisa_contexto = any(kw in prompt_lower for kw in palavras_contexto)
    if precisa_contexto and req.table_context and req.table_context.strip():
        prompt_final = req.prompt.strip() + "\n\n" + req.table_context.strip()
    else:
        prompt_final = req.prompt.strip()

    # ── 4. Carregar System Prompt ──
    prompt_path = os.path.join(os.path.dirname(os.path.abspath(__file__)), "prompt_rede_eletrica.txt")
    system_instruction = "Você é um especialista em redes elétricas."
    if os.path.exists(prompt_path):
        with open(prompt_path, "r", encoding="utf-8") as f:
            system_instruction = f.read()

    # ── 5. Injeção de Regras (Simplificado para todas as IAs) ──
    try:
        conn2 = sqlite3.connect(DB_PATH)
        cursor2 = conn2.cursor()
        cursor2.execute("SELECT conteudo FROM regras")
        regras = cursor2.fetchall()
        conn2.close()

        if regras:
            regras_txt = "\n".join(f"- {r[0]}" for r in regras)
            system_instruction += f"\n\n🔹 REGRAS APRENDIDAS:\n{regras_txt}"
    except Exception as e:
        print(f"Erro ao carregar regras: {e}")

    # ── 6. Stream ──

    # ── Gemini ──
    if provider == "gemini":
        from google import genai
        from google.genai import types
        client = genai.Client(api_key=api_key)

        preferencias = [
            "gemini-3.1-flash-lite",
            "gemini-2.5-flash",
            "gemini-3.6-flash",
            "gemini-1.5-flash",
            "gemini-1.5-pro",
        ]
        modelos = [custom_model] if custom_model else preferencias
        if len(req.prompt.split()) < 15 and not precisa_contexto:
            modelos = ["gemini-3.1-flash-lite"] + [m for m in modelos if m != "gemini-3.1-flash-lite"]

        sdk_history = []
        for msg in req.history:
            parts = msg.get("parts", [])
            if parts:
                texto = parts[0].get("text", "") if isinstance(parts[0], dict) else str(parts[0])
                sdk_history.append({"role": msg.get("role", "user"), "parts": [{"text": texto}]})

        async def _stream_gemini():
            last_err = None
            for modelo in modelos:
                try:
                    chat = client.aio.chats.create(
                        model=modelo,
                        config=types.GenerateContentConfig(system_instruction=system_instruction),
                        history=sdk_history
                    )
                    response = await chat.send_message_stream(prompt_final)
                    async for chunk in response:
                        if chunk.text:
                            yield chunk.text
                    return  # sucesso
                except Exception as e:
                    err = str(e).lower()
                    if any(k in err for k in ["429", "quota", "exhausted", "not found", "404", "unavailable"]):
                        last_err = e
                        print(f"Gemini {modelo} indisponiível ({type(e).__name__}). Fallback...")
                        continue
                    # Erro inesperado: emite como texto para o frontend mostrar
                    yield f"\n[ERRO] {str(e)}"
                    return
            # Nenhum modelo funcionou
            msg = str(last_err) if last_err else "Nenhum modelo Gemini disponível."
            yield f"\n[ERRO] {msg}"

        return StreamingResponse(_stream_gemini(), media_type="text/plain")




@app.get("/trigger-file")
async def trigger_file(path: str):
    await manager.broadcast(json.dumps({"type": "load_file", "path": path}))
    return {"status": "success"}

# =====================================================
# PDF HELPERS
# =====================================================

def color_to_hex(c):
    if c is None: return "#000000"
    try:
        if isinstance(c, (list, tuple)):
            if len(c) >= 3:
                rgb = [max(0, min(255, int(x * 255))) for x in c[:3]]
                return "#%02x%02x%02x" % tuple(rgb)
        val = int(c)
        return "#%06x" % (val & 0xFFFFFF)
    except: return "#000000"

def flags_decomposer(flags):
    l = []
    if flags & 1: l.append("superscript")
    if flags & 2: l.append("italic")
    if flags & 4: l.append("serifed")
    else: l.append("sans")
    if flags & 8: l.append("monospaced")
    else: l.append("proportional")
    if flags & 16: l.append("bold")
    return ", ".join(l)

# =====================================================
# DXF HELPERS (CORRIGIDO)
# =====================================================

def _layer_name_to_hex(layer_name: str) -> str | None:
    """
    Extrai cor RGB do padrão de nomenclatura de layer usado neste DXF.
    Exemplo: 'TXT_FF0000' → '#FF0000', 'TXT_0000FF' → '#0000FF'.
    Retorna None se o layer não seguir o padrão.
    """
    m = re.match(r'.*_([0-9A-Fa-f]{6})$', layer_name)
    return f"#{m.group(1).upper()}" if m else None


def _aci_to_hex(aci: int) -> str:
    """
    Converte índice ACI para hex RGB.
    ACI 7 = branco (fundo escuro) → mapeado para preto no papel branco.
    ACI 0 = BYBLOCK → preto.
    """
    if aci in (0, 7):
        return "#000000"
    try:
        rgb = aci2rgb(aci)
        return f"#{rgb[0]:02x}{rgb[1]:02x}{rgb[2]:02x}"
    except:
        return "#000000"


# Remove códigos de formatação de parágrafo MTEXT: \pxql; \pxqc; \pxqr; etc.
# IMPORTANTE:
#   - SEM re.IGNORECASE: \p (minúsculo) = formatação de parágrafo
#                        \P (maiúsculo) = quebra de parágrafo — NÃO deve ser removido
#   - Ponto-e-vírgula OBRIGATÓRIO (sem ?): garante que \P seguido de texto
#     (ex: \PCAZ 9,5) não seja confundido com código de formatação.
_RE_PARA_FMT = re.compile(r'\\p[^;\\\\]*;')


def _get_entity_base_color(entity, doc) -> str:
    """
    Resolve a cor da entidade a partir de atributos de nivel de entidade/layer.
    NAO verifica codigos inline (\\C, \\c) -- esses sao resolvidos por
    paragrafo em _resolve_part_color para permitir heranca correta.
    Hierarquia:
      1. true_color da entidade
      2. ACI da entidade (nao-BYLAYER, nao-BYBLOCK)
      3. RGB no nome do layer (TXT_RRGGBB)
      4. true_color do layer
      5. ACI do layer
      6. Fallback preto
    """
    if entity.dxf.hasattr("true_color"):
        try:
            rgb = int2rgb(entity.dxf.true_color)
            return f"#{rgb[0]:02x}{rgb[1]:02x}{rgb[2]:02x}"
        except:
            pass
    aci = entity.dxf.color
    if aci not in (0, 256):
        return _aci_to_hex(aci)
    try:
        layer_name = entity.dxf.layer
        h = _layer_name_to_hex(layer_name)
        if h:
            return h
        layer = doc.layers.get(layer_name)
        if layer.dxf.hasattr("true_color"):
            rgb = int2rgb(layer.dxf.true_color)
            return f"#{rgb[0]:02x}{rgb[1]:02x}{rgb[2]:02x}"
        return _aci_to_hex(layer.color)
    except:
        pass
    return "#000000"


def _resolve_part_color(part_raw: str, inherited: str) -> str:
    """
    Resolve a cor de um parágrafo individual.
    Inline \\c<decimal> (true-color) e \\C<n> (ACI) dentro do part
    substituem a cor herdada. O ÚLTIMO código encontrado no parágrafo 
    é o que define a cor da linha e o que será herdado pelo próximo.
    """
    # Procura todos os códigos de cor no parágrafo
    # \c é true-color (RGB em decimal), \C é ACI
    matches = re.findall(r'\\([cC])(\d+)', part_raw)
    if not matches:
        return inherited
    
    # Pega o último código (é o que "vence" no final do parágrafo)
    code_type, value = matches[-1]
    
    if code_type == 'c':
        try:
            rgb = int2rgb(int(value))
            return f"#{rgb[0]:02x}{rgb[1]:02x}{rgb[2]:02x}"
        except:
            pass
    else:
        aci = int(value)
        if aci == 0 or aci == 256:
            return inherited
        return _aci_to_hex(aci)
        
    return inherited


def get_dxf_color(entity, doc) -> str:
    """
    Resolve a cor real de uma entidade TEXT/MTEXT.
    Para MTEXT com multiplos paragrafos, use _get_entity_base_color +
    _resolve_part_color por paragrafo (feito em _process_entity).
    """
    raw = entity.text if entity.dxftype() == "MTEXT" else None
    if raw:
        m = re.search(r'\\c(\d+)', raw)
        if m:
            try:
                rgb = int2rgb(int(m.group(1)))
                return f"#{rgb[0]:02x}{rgb[1]:02x}{rgb[2]:02x}"
            except:
                pass
        m = re.search(r'\\C(\d+)', raw)
        if m:
            return _aci_to_hex(int(m.group(1)))
    return _get_entity_base_color(entity, doc)


def clean_dxf_text(raw: str) -> str:
    """
    Limpa uma fatia de texto MTEXT ja separada por paragrafo.
    Aplica plain_mtext apos remover \\pxq...; e depois limpa residuos.
    """
    if not raw:
        return ""
    text = _RE_PARA_FMT.sub('', raw)                        # remove \pxql; etc.
    text = plain_mtext(text)                                 # limpa \f \C \c {}
    text = re.sub(r'\\[a-zA-Z][^\\;\s]*;?', '', text)      # residuos
    text = re.sub(r'\\{2,}', '', text)                      # barras soltas
    text = re.sub(r'[ \t]+', ' ', text)
    return text.strip()


def _process_entity(entity, doc, base_color=None, block_name=None) -> list[dict]:
    """
    Transforma uma entidade DXF (TEXT, MTEXT, ATTRIB) em uma lista de dicionários formatados.
    Suporta quebra de parágrafos MTEXT e herança de cor.
    """
    tp = entity.dxftype()
    # TEXT e ATTRIB usam .dxf.text, MTEXT usa .text
    raw = entity.dxf.text if tp in ("TEXT", "ATTRIB") else entity.text

    # Cor base da entidade (sem codigos inline — esses sao por paragrafo)
    if base_color and entity.dxf.color == 0:
        entity_base = base_color
    else:
        entity_base = _get_entity_base_color(entity, doc)

    if tp == "TEXT" or tp == "ATTRIB":
        height = getattr(entity.dxf, "height", 0)
    else:
        height = getattr(entity.dxf, "char_height", 0)
    
    style = getattr(entity.dxf, "style", "Standard")
    pos = entity.dxf.insert

    if tp == "MTEXT":
        # Passo 1: remove \pxql; \pxqc; \pxqr; etc.
        pre = _RE_PARA_FMT.sub('', raw)
        # Passo 2: divide por \P, quebras de linha ou mudanças de cor internas
        # Lookahead (?=...) divide ANTES do código, Lookbehind (?<=\}) divide APÓS o bloco.
        # (?<!\{) garante que não dividimos entre o { e o código de cor (\C1;).
        parts = re.split(r'\\P|[\r\n]+|\^M|\^J|(?=\{\\[cC]\d+;)|(?<!\{)(?=\\[cC]\d+;)|(?<=\})', pre)
    else:
        parts = [raw]

    rows = []
    current_color = entity_base  # cor herdada entre paragrafos (\P)

    for part in parts:
        if not part or not part.strip():
            continue
        
        # Resolve a cor desta parte. 
        # Se estiver entre chaves { }, a cor é local.
        p_stripped = part.strip()
        is_block = p_stripped.startswith('{') and p_stripped.endswith('}')
        item_color = _resolve_part_color(part, current_color)
        
        if tp == "MTEXT":
            # Passo 4: limpa o texto do paragrafo/bloco.
            txt = clean_dxf_text(part)
        else:
            txt = re.sub(r'[ \t]+', ' ', part).strip()

        if txt:
            flags = f"DXF_{tp}"
            if block_name:
                flags += f" | BLOCO:{block_name}"

            rows.append({
                "pagina": 1,
                "texto": txt,
                "fonte": style,
                "tamanho": round(height, 2),
                "cor": item_color,
                "flags": flags,
                "layer": getattr(entity.dxf, "layer", ""),
                "_y": float(pos.y),
                "_x": float(pos.x),
            })
            
            # Se for uma quebra de paragrafo real ou mudança global (sem chaves), 
            # atualiza a cor herdada. Se for bloco { }, não atualiza a herança global.
            if not is_block:
                current_color = item_color
    return rows


def extract_dxf_content(doc) -> list[dict]:
    """
    Extrai todo o conteúdo textual do model space, blocos e atributos,
    incluindo a identificação dos nomes dos blocos.
    """
    msp = doc.modelspace()
    all_rows: list[dict] = []

    # 1. Entidades diretas no model space (Texto e MText)
    for entity in msp.query("TEXT MTEXT"):
        all_rows.extend(_process_entity(entity, doc))

    # 2. Entidades dentro de blocos (INSERT)
    for insert in msp.query("INSERT"):
        bname = insert.dxf.name
        try:
            block = doc.blocks.get(bname)
            bcolor = _get_entity_base_color(insert, doc)
            
            # Extrai o nome do bloco como uma entrada (útil para identificar símbolos)
            all_rows.append({
                "pagina": 1,
                "texto": f"[BLOCO: {bname}]",
                "fonte": "BlockName",
                "tamanho": 0,
                "cor": bcolor,
                "flags": "DXF_INSERT",
                "layer": getattr(insert.dxf, "layer", ""),
                "_y": float(insert.dxf.insert.y),
                "_x": float(insert.dxf.insert.x),
            })

            # Extrai textos, mtexts e atributos de DENTRO da definição do bloco
            for entity in block.query("TEXT MTEXT ATTRIB"):
                all_rows.extend(_process_entity(entity, doc, base_color=bcolor, block_name=bname))
            
            # Extrai atributos específicos desta INSTÂNCIA (insert.attribs)
            for attr in insert.attribs:
                all_rows.extend(_process_entity(attr, doc, base_color=bcolor, block_name=bname))
                
        except Exception as e:
            print(f"Erro ao processar bloco {bname}: {e}")
            pass

    # Ordenação: maior Y primeiro (topo), desempate por X (esquerda)
    all_rows.sort(key=lambda r: (-r["_y"], r["_x"]))
    for r in all_rows:
        r.pop("_y", None)
        r.pop("_x", None)

    return all_rows

# =====================================================
# API ENDPOINTS
# =====================================================

@app.post("/upload")
async def upload_file(file: UploadFile = File(...)):
    try:
        contents = await file.read()
        if file.filename.lower().endswith(".dxf"):
            with tempfile.NamedTemporaryFile(delete=False, suffix=".dxf") as tmp:
                tmp.write(contents)
                tmp_path = tmp.name
            try:
                doc = ezdxf.readfile(tmp_path)
                return {"data": extract_dxf_content(doc)}
            finally:
                if os.path.exists(tmp_path): os.remove(tmp_path)

        # PDF
        doc = pymupdf.open(stream=contents, filetype="pdf")
        extracted = []
        for i, page in enumerate(doc):
            blocks = page.get_text("dict", flags=11)["blocks"]
            for b in blocks:
                if "lines" not in b: continue
                for l in b["lines"]:
                    if "spans" not in l: continue
                    for s in l["spans"]:
                        text = s["text"].strip()
                        if text:
                            extracted.append({
                                "pagina": i + 1,
                                "texto": text,
                                "fonte": s["font"],
                                "tamanho": round(s["size"], 2),
                                "cor": color_to_hex(s.get("color", 0)),
                                "flags": flags_decomposer(s["flags"])
                            })
        return {"data": extracted}
    except Exception as e:
        return {"error": str(e), "data": []}

@app.get("/extract-local")
async def extract_local(path: str):
    if not os.path.exists(path): return {"error": "Arquivo não encontrado"}
    try:
        if path.lower().endswith(".dxf"):
            doc = ezdxf.readfile(path)
            return {"data": extract_dxf_content(doc), "filename": os.path.basename(path)}
        # PDF
        doc = pymupdf.open(path)
        extracted = []
        for i, page in enumerate(doc):
            blocks = page.get_text("dict", flags=11)["blocks"]
            for b in blocks:
                if "lines" not in b: continue
                for l in b["lines"]:
                    if "spans" not in l: continue
                    for s in l["spans"]:
                        text = s["text"].strip()
                        if text:
                            extracted.append({
                                "pagina": i + 1,
                                "texto": text,
                                "fonte": s["font"],
                                "tamanho": round(s["size"], 2),
                                "cor": color_to_hex(s.get("color", 0)),
                                "flags": flags_decomposer(s["flags"])
                            })
        return {"data": extracted, "filename": os.path.basename(path)}
    except Exception as e:
        return {"error": str(e), "data": []}

@app.post("/api/importar-rec-pdf")
async def importar_rec_pdf(orcamento: UploadFile = File(...), lista: UploadFile = File(...)):
    try:
        orcamento_contents = await orcamento.read()
        lista_contents = await lista.read()
        
        extracted_items = []
        
        # Processa ORÇAMENTO
        doc_orc = pymupdf.open(stream=orcamento_contents, filetype="pdf")
        current_mdo = "MÃO-DE-OBRA"
        
        for page in doc_orc:
            blocks = page.get_text("dict", flags=11)["blocks"]
            texts = []
            for b in blocks:
                if "lines" not in b: continue
                for l in b["lines"]:
                    if "spans" not in l: continue
                    for s in l["spans"]:
                        text = s["text"].strip()
                        if text:
                            texts.append(text)
            
            for i, text in enumerate(texts):
                if text == "MAO-DE-OBRA":
                    current_mdo = "MÃO-DE-OBRA"
                elif text == "MATERIAL":
                    current_mdo = "MATERIAL"
                
                if re.match(r'^\d{6}$', text):
                    codigo = text.lstrip('0') or '0'
                    desc = texts[i-1] if i > 0 else ""
                    
                    qtd = 1.0
                    for j in range(i-1, max(-1, i-5), -1):
                        # Aceita números no formato 1.040,43 ou 1040.43 ou 1040,43
                        if re.match(r'^-?\d{1,3}(\.\d{3})*(,\d+)?$', texts[j]) or re.match(r'^-?\d+(,\d+)?$', texts[j]) or re.match(r'^-?\d+(\.\d+)?$', texts[j]):
                            try:
                                # Remove pontos de milhar, troca vírgula por ponto
                                clean_val = texts[j].replace('.', '') if ',' in texts[j] and texts[j].count('.') >= 1 else texts[j]
                                clean_val = clean_val.replace(',', '.')
                                qtd = float(clean_val)
                                break
                            except ValueError:
                                pass
                    
                    op = texts[i+1] if i + 1 < len(texts) else "I"
                    if op not in ['I', 'R']:
                        op = 'I'
                    
                    extracted_items.append({
                        "operacao": op,
                        "mdo": current_mdo,
                        "codigo": codigo,
                        "desc_codigo": desc,
                        "total": qtd
                    })
        doc_orc.close()

        # Processa LISTA
        doc_lista = pymupdf.open(stream=lista_contents, filetype="pdf")
        for page in doc_lista:
            blocks = page.get_text("dict", flags=11)["blocks"]
            texts = []
            for b in blocks:
                if "lines" not in b: continue
                for l in b["lines"]:
                    if "spans" not in l: continue
                    for s in l["spans"]:
                        text = s["text"].strip()
                        if text:
                            texts.append(text)
            
            for i, text in enumerate(texts):
                if re.match(r'^\d{6}$', text):
                    codigo = text.lstrip('0') or '0'
                    requisitar = None
                    devolver = None
                    desc = ""
                    
                    # Look ahead up to 10 elements to find Requisitar, Devolver, and Description
                    for j in range(i+1, min(i+10, len(texts))):
                        if re.match(r'^-?\d{1,3}(\.\d{3})*(,\d+)?$', texts[j]) or re.match(r'^-?\d+(,\d+)?$', texts[j]) or re.match(r'^-?\d+(\.\d+)?$', texts[j]):
                            try:
                                clean_val = texts[j].replace('.', '') if ',' in texts[j] and texts[j].count('.') >= 1 else texts[j]
                                clean_val = clean_val.replace(',', '.')
                                val = float(clean_val)
                                if requisitar is None:
                                    requisitar = val
                                elif devolver is None:
                                    devolver = val
                                    if j + 1 < len(texts):
                                        desc = texts[j+1]
                                    break
                            except ValueError:
                                pass
                    
                    if devolver and devolver > 0:
                        extracted_items.append({
                            "operacao": "R",
                            "mdo": "MATERIAL",
                            "codigo": codigo,
                            "desc_codigo": desc,
                            "total": devolver
                        })
        doc_lista.close()
        
        return {"resultado": extracted_items}
    except Exception as e:
        return {"error": str(e)}

# =====================================================
# STATIC FILES
# =====================================================

if getattr(sys, "frozen", False):
    base_dir = sys._MEIPASS
else:
    base_dir = os.path.dirname(os.path.abspath(__file__))

static_dir = os.path.join(base_dir, "static")
if not os.path.exists(static_dir): os.makedirs(static_dir)

app.mount("/static", StaticFiles(directory=static_dir), name="static")

NO_CACHE_HEADERS = {
    "Cache-Control": "no-store, no-cache, must-revalidate, max-age=0",
    "Pragma": "no-cache",
    "Expires": "0",
}

@app.get("/")
async def serve_index():
    index_path = os.path.join(static_dir, "index.html")
    if not os.path.exists(index_path): return HTMLResponse("index.html não encontrado na pasta static")
    with open(index_path, "r", encoding="utf-8") as f:
        return HTMLResponse(f.read(), headers=NO_CACHE_HEADERS)

@app.get("/resumo")
async def serve_resumo():
    resumo_path = os.path.join(static_dir, "resumo.html")
    if not os.path.exists(resumo_path): return HTMLResponse("resumo.html não encontrado na pasta static")
    with open(resumo_path, "r", encoding="utf-8") as f:
        return HTMLResponse(f.read(), headers=NO_CACHE_HEADERS)

@app.get("/static/resumo.js")
async def serve_resumo_js():
    js_path = os.path.join(static_dir, "resumo.js")
    if not os.path.exists(js_path): return HTMLResponse("resumo.js não encontrado")
    with open(js_path, "r", encoding="utf-8") as f:
        from fastapi.responses import Response
        return Response(content=f.read(), media_type="application/javascript", headers=NO_CACHE_HEADERS)

def open_browser(url):
    webbrowser.open(url)

class OrcamentoRequest(BaseModel):
    cabos: List[Dict[str, Any]]
    outros: List[Dict[str, Any]]

@app.post("/api/orcamento/upload")
async def upload_orcamento(file: UploadFile = File(...)):
    try:
        contents = await file.read()
        # Decode and parse CSV
        text = contents.decode("utf-8-sig")
        reader = csv.DictReader(io.StringIO(text), delimiter=";")
        
        # fallback to comma if no columns found
        if not reader.fieldnames or len(reader.fieldnames) < 2:
            reader = csv.DictReader(io.StringIO(text), delimiter=",")
            
        required_cols = ["ATIVO", "DESC ATIVO", "COMPONENTE", "PROJETO", "MDO", "CODIGO", "DESC CODIGO", "FATOR I", "FATOR R"]
        # Convert to upper just in case
        actual_cols = [c.strip().upper() for c in (reader.fieldnames or []) if c]
        
        conn = sqlite3.connect(DB_PATH)
        cursor = conn.cursor()
        
        # Limpar tabela atual
        cursor.execute("DELETE FROM tabela_orcamento")
        
        for row in reader:
            # Map by ignoring case
            row_upper = {k.strip().upper(): v for k, v in row.items() if k}
            ativo = row_upper.get("ATIVO", "")
            codigo = row_upper.get("CODIGO", "")
            if not ativo and not codigo: continue
            
            try:
                fator_i = float(row_upper.get("FATOR I", "0").replace(",", "."))
            except ValueError:
                fator_i = 0.0
                
            try:
                fator_r = float(row_upper.get("FATOR R", "0").replace(",", "."))
            except ValueError:
                fator_r = 0.0
                
            cursor.execute('''
                INSERT INTO tabela_orcamento (ativo, desc_ativo, componente, projeto, mdo, codigo, desc_codigo, fator_i, fator_r)
                VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?)
            ''', (
                ativo,
                row_upper.get("DESC ATIVO", ""),
                row_upper.get("COMPONENTE", ""),
                row_upper.get("PROJETO", ""),
                row_upper.get("MDO", ""),
                row_upper.get("CODIGO", ""),
                row_upper.get("DESC CODIGO", ""),
                fator_i,
                fator_r
            ))
            
        conn.commit()
        conn.close()
        return {"status": "ok", "message": "Tabela carregada com sucesso."}
    except Exception as e:
        return JSONResponse(status_code=400, content={"error": str(e)})

@app.get("/api/orcamento/dados")
def get_orcamento_dados():
    conn = sqlite3.connect(DB_PATH)
    conn.row_factory = sqlite3.Row
    cursor = conn.cursor()
    cursor.execute("SELECT * FROM tabela_orcamento")
    rows = cursor.fetchall()
    conn.close()
    return {"dados": [dict(r) for r in rows]}

from pydantic import BaseModel
from typing import List, Dict, Any

class SalvarOrcamentoRequest(BaseModel):
    dados: List[Dict[str, Any]]

@app.post("/api/orcamento/salvar")
def salvar_orcamento(req: SalvarOrcamentoRequest):
    try:
        conn = sqlite3.connect(DB_PATH)
        cursor = conn.cursor()
        cursor.execute("DELETE FROM tabela_orcamento")
        for row in req.dados:
            ativo = row.get("ativo", "").strip()
            codigo = row.get("codigo", "").strip()
            if not ativo and not codigo: continue
            
            try:
                fator_i = float(str(row.get("fator_i", "0")).replace(",", "."))
            except ValueError:
                fator_i = 0.0
                
            try:
                fator_r = float(str(row.get("fator_r", "0")).replace(",", "."))
            except ValueError:
                fator_r = 0.0
                
            cursor.execute('''
                INSERT INTO tabela_orcamento (ativo, desc_ativo, componente, projeto, mdo, codigo, desc_codigo, fator_i, fator_r)
                VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?)
            ''', (
                ativo,
                row.get("desc_ativo", "").strip(),
                row.get("componente", "").strip(),
                row.get("projeto", "").strip(),
                row.get("mdo", "").strip(),
                codigo,
                row.get("desc_codigo", "").strip(),
                fator_i,
                fator_r
            ))
            
        conn.commit()
        conn.close()
        return {"status": "ok", "message": "Tabela atualizada com sucesso."}
    except Exception as e:
        return JSONResponse(status_code=400, content={"error": str(e)})

class RecSaveRequest(BaseModel):
    numero_obra: str
    dados: List[Dict[str, Any]]

@app.post("/api/rec/salvar")
def salvar_rec(req: RecSaveRequest):
    try:
        from datetime import datetime
        import json
        conn = sqlite3.connect(DB_PATH)
        cursor = conn.cursor()
        data_agora = datetime.now().isoformat()
        dados_str = json.dumps(req.dados)
        
        cursor.execute("INSERT OR REPLACE INTO historico_rec (numero_obra, dados_json, data_criacao) VALUES (?, ?, ?)",
                       (req.numero_obra.strip(), dados_str, data_agora))
        conn.commit()
        conn.close()
        return {"status": "ok"}
    except Exception as e:
        return JSONResponse(status_code=400, content={"error": str(e)})

@app.get("/api/rec/{numero_obra}")
def recuperar_rec(numero_obra: str):
    try:
        import json
        conn = sqlite3.connect(DB_PATH)
        conn.row_factory = sqlite3.Row
        cursor = conn.cursor()
        cursor.execute("SELECT dados_json FROM historico_rec WHERE numero_obra = ?", (numero_obra.strip(),))
        row = cursor.fetchone()
        conn.close()
        
        if row:
            return {"status": "ok", "dados": json.loads(row["dados_json"])}
        else:
            return JSONResponse(status_code=404, content={"error": "Obra não encontrada"})
    except Exception as e:
        return JSONResponse(status_code=400, content={"error": str(e)})

class DetalhesRequest(BaseModel):
    codigos: List[str]

@app.post("/api/orcamento/detalhes")
def get_detalhes_codigos(req: DetalhesRequest):
    if not req.codigos:
        return {}
    
    conn = sqlite3.connect(DB_PATH)
    conn.row_factory = sqlite3.Row
    cursor = conn.cursor()
    
    placeholders = ",".join(["?"] * len(req.codigos))
    cursor.execute(f"SELECT codigo, desc_codigo, mdo FROM tabela_orcamento WHERE codigo IN ({placeholders})", tuple(req.codigos))
    rows = cursor.fetchall()
    conn.close()
    
    resultado = {}
    for r in rows:
        resultado[r["codigo"]] = {
            "desc_codigo": r["desc_codigo"],
            "mdo": r["mdo"]
        }
    return resultado

@app.post("/api/orcamento/calcular")
def calcular_orcamento(req: OrcamentoRequest):
    conn = sqlite3.connect(DB_PATH)
    conn.row_factory = sqlite3.Row
    cursor = conn.cursor()
    cursor.execute("SELECT * FROM tabela_orcamento")
    orcamento_rows = cursor.fetchall()
    conn.close()

    # ── 1. Índice ATIVO → linhas do orçamento (O(1) lookup) ─────────────────
    orcamento_dict = {}
    for r in orcamento_rows:
        key = r["ativo"].strip().upper()
        orcamento_dict.setdefault(key, []).append(dict(r))

    # ── 2. Extrair tabela interna [ativo, qtd, operacao] ────────────────────
    ativos_qtd = []

    # 2a. CABOS: Regra rígida de formato "ATIVO FASE COMPRIMENTO"
    for item in req.cabos:
        ativo_raw = item.get("ativo", "").strip().upper()
        operacao  = item.get("operacao", "").strip().upper()
        if not ativo_raw:
            continue
            
        txt = ativo_raw
        txt = txt.replace("CAA ", "CAA")
        txt = txt.replace("CA ", "CA")
        txt = txt.replace("CU ", "CU")
        txt = txt.replace("CAZ ", "CAZ")
        txt = txt.replace("P ", "P")
        
        # O 'm' final pode ser removido
        txt = re.sub(r'\s+m\s*$', '', txt, flags=re.IGNORECASE)
        
        parts = txt.split()
        if not parts:
            continue
            
        nome_ativo = parts[0]
        
        fases = 1.0
        comprimento = 1.0
        
        if len(parts) >= 3:
            try:
                fases = float(parts[1].replace(",", "."))
            except Exception:
                pass
            try:
                comprimento = float(parts[2].replace(",", "."))
            except Exception:
                pass
        elif len(parts) == 2:
            # Fallback caso tenham omitido fase ou comprimento
            try:
                comprimento = float(parts[1].replace(",", "."))
            except:
                pass
                
        # Prioriza qtdAtivos calculado pelo frontend; fallback para lógica local
        qtd_ativos_frontend = item.get("qtdAtivos")
        if qtd_ativos_frontend is not None:
            try:
                multiplicador = float(qtd_ativos_frontend)
            except (ValueError, TypeError):
                multiplicador = fases
        elif nome_ativo.startswith("CAZ") or nome_ativo.startswith("M"):
            multiplicador = 1.0
        else:
            multiplicador = fases
            
        # Adiciona a margem de 5% sobre a quantidade processada (comprimento * multiplicador)
        qtd_final = (comprimento * multiplicador) * 1.05
        
        ativos_qtd.append({"ativo": nome_ativo, "qtd": qtd_final, "operacao": operacao})

    # 2b. OUTROS: texto separado por espaços no modelo "[qtd]-[ativo] [qtd]-[ativo] ..."
    for item in req.outros:
        ativo_raw = item.get("ativo", "").strip()
        operacao  = item.get("operacao", "").strip().upper()
        if not ativo_raw:
            continue
        txt = ativo_raw.strip()
        # Se inicia com DT ou CV sem quantidade, adiciona "1-"
        if re.match(r'^(DT|CV)', txt, re.IGNORECASE) and not re.match(r'^\d', txt):
            txt = "1-" + txt
        # "-" → " " (separa quantidade do ativo)
        txt = txt.replace("-", " ")
        # "*" → "-" (representa negativo / linha viva)
        txt = txt.replace("*", "-")
        # Posições pares = quantidades, posições ímpares = ativos
        tokens = txt.split()
        i = 0
        while i + 1 < len(tokens):
            try:
                qtd = float(tokens[i].replace(",", "."))
            except ValueError:
                qtd = 1.0
            nome_str = tokens[i + 1].upper()
            ativos_qtd.append({"ativo": nome_str, "qtd": qtd, "operacao": operacao})
            i += 2

    # ── 3. Cross-reference: Ativo → 1º match → Componente → todos os Códigos ─
    # Índice auxiliar: COMPONENTE → lista de todas as linhas daquele componente
    componente_dict = {}
    for rows_list in orcamento_dict.values():
        for row in rows_list:
            comp = row.get("componente", "").strip().upper()
            if comp:
                componente_dict.setdefault(comp, []).append(row)

    # Mapa garantido: codigo → row representativa (prioriza as que têm MDO)
    codigo_lookup = {}
    for rows_list in orcamento_dict.values():
        for row in rows_list:
            codigo = row["codigo"]
            if codigo not in codigo_lookup:
                codigo_lookup[codigo] = row
            elif (row.get("mdo") or "").strip() and not (codigo_lookup[codigo].get("mdo") or "").strip():
                codigo_lookup[codigo] = row

    componentes_qtd = []
    nao_encontrados = []

    # Índice auxiliar: DESC_ATIVO (upper) → lista de rows (para busca parcial)
    desc_ativo_index = {}
    for rows_list in orcamento_dict.values():
        for row in rows_list:
            desc = (row.get("desc_ativo") or "").strip().upper()
            if desc:
                desc_ativo_index.setdefault(desc, []).append(row)

    for av in ativos_qtd:
        nome_ativo = av["ativo"]
        linhas_ativo = orcamento_dict.get(nome_ativo, [])

        if not linhas_ativo:
            # ── Passo 2: busca por COMPONENTE exato ─────────────────────────
            if nome_ativo in componente_dict:
                linhas_ativo = componente_dict[nome_ativo]

        if not linhas_ativo:
            # ── Passo 3: busca por DESC_ATIVO exato ─────────────────────────
            if nome_ativo in desc_ativo_index:
                linhas_ativo = desc_ativo_index[nome_ativo]

        if not linhas_ativo:
            # ── Passo 4: busca parcial em DESC_ATIVO (nome_ativo contido na desc) ─
            matches = [rows for desc, rows in desc_ativo_index.items() if nome_ativo in desc]
            if matches:
                linhas_ativo = matches[0]  # Usa o primeiro match parcial

        if not linhas_ativo:
            # ── Passo 5: busca como CÓDIGO direto ───────────────────────────
            if nome_ativo in codigo_lookup:
                row_clone = dict(codigo_lookup[nome_ativo])
                row_clone["fator_i"] = 1.0
                row_clone["fator_r"] = 1.0
                componentes_qtd.append({
                    "qtd":      av["qtd"],
                    "operacao": av["operacao"],
                    "row":      row_clone,
                })
            else:
                nao_encontrados.append(nome_ativo)
            continue

        # Pega o COMPONENTE do primeiro match do ativo
        componente = linhas_ativo[0].get("componente", "").strip().upper()
        if not componente:
            continue

        # Agora busca TODOS os códigos desse componente
        # Para cada código, garante que a row escolhida tem mdo preenchido (se existir)
        linhas_componente = componente_dict.get(componente, [])
        
        # Agrupa linhas por codigo, priorizando as que têm mdo não-vazio
        melhor_row_por_codigo = {}
        for row in linhas_componente:
            codigo = row["codigo"]
            mdo_row = (row.get("mdo") or "").strip()
            if codigo not in melhor_row_por_codigo:
                melhor_row_por_codigo[codigo] = row
            elif mdo_row and not (melhor_row_por_codigo[codigo].get("mdo") or "").strip():
                # Substitui por essa que tem mdo preenchido
                melhor_row_por_codigo[codigo] = row

        for codigo, row in melhor_row_por_codigo.items():
            componentes_qtd.append({
                "qtd":      av["qtd"],
                "operacao": av["operacao"],
                "row":      row,
            })

    # ── 4. Soma I e Soma R por CÓDIGO (equivalente ao SOMASE do Excel) ───────
    codigos = {}
    for c in componentes_qtd:
        row     = c["row"]
        codigo  = row["codigo"]
        # MDO: da linha, ou fallback pelo código — nunca fica vazio (usa "-" se não achar em lugar nenhum)
        fallback_mdo = codigo_lookup.get(codigo, {}).get("mdo", "")
        mdo     = (row.get("mdo") or "").strip() or (fallback_mdo or "").strip() or "-"
        key     = (codigo, mdo)
        is_inst = c["operacao"] in ("I", "*I")
        fator_i = float(row.get("fator_i") or 0)
        fator_r = float(row.get("fator_r") or 0)
        if key not in codigos:
            codigos[key] = {
                "mdo":       mdo,
                "codigo":    codigo,
                "desc_codigo": row.get("desc_codigo", ""),
                "soma_i":    0.0,
                "soma_r":    0.0,
            }
        if is_inst:
            codigos[key]["soma_i"] += c["qtd"] * fator_i
        else:
            codigos[key]["soma_r"] += c["qtd"] * fator_r

    # ── 5. Montar resultado com uma linha por operação presente ──────────────
    resultado = []
    for key in codigos.keys():
        v = codigos[key]
        if v["soma_i"] > 0:
            resultado.append({"operacao": "I", "mdo": v["mdo"], "codigo": v["codigo"],
                               "desc_codigo": v["desc_codigo"], "total": round(v["soma_i"], 2)})
        if v["soma_r"] > 0:
            resultado.append({"operacao": "R", "mdo": v["mdo"], "codigo": v["codigo"],
                               "desc_codigo": v["desc_codigo"], "total": round(v["soma_r"], 2)})

    # Ordenar: Descrição A-Z, depois Operação A-Z, depois MDO A-Z
    resultado.sort(key=lambda x: (x["desc_codigo"].upper(), x["operacao"], x["mdo"].upper()))

    return {"resultado": resultado, "nao_encontrados": list(set(nao_encontrados))}



if __name__ == "__main__":
    import uvicorn
    port = 8000
    if "PORT" not in os.environ:
        threading.Timer(1.5, open_browser, args=(f"http://127.0.0.1:{port}",)).start()
    uvicorn.run(app, host="0.0.0.0", port=port)