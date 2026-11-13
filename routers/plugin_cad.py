"""
routers/plugin_cad.py — API para o plugin do nanoCAD (plugin_nanocad/, TASK-104).

O plugin lê o desenho dentro do CAD com as regras do projeto. Com a aplicação aberta (localhost ou online), ele passa a usar as
regras GRAVADAS AQUI e a conferir a obra contra a Base de Orçamento, sem repetir a lógica do cálculo:

    GET  /api/plugin/versao?projeto=081    impressão digital do que muda o resultado (regras do leitor, vínculo, conversão, base).
                                           O plugin consulta a cada poucos segundos; mudou → baixa de novo e reconfere.
    GET  /api/plugin/regras?projeto=081    regras do leitor (processamento + classificação) e de vínculo do projeto, numa chamada.
    POST /api/plugin/conferir              {"projeto", "dados": {"cabos", "outros"}} (as linhas do .obra.json) → para cada linha, os
                                           ativos que a Totalizadora gera e se cada um existe na Base de Orçamento do projeto
                                           (mesmo caminho do modo autônomo: montar_totalizadora → processar_calculo, RN-06).
    POST /api/plugin/obra                  {"chave", "envelope"} (TASK-106): "Abrir na aplicação" — grava a obra do desenho como obra
                                           PARTICULAR do usuário (a mesma obra a cada envio do mesmo desenho, sem duplicar) e devolve o
                                           link do Resumo que a abre já carregada (/resumo?obra=…&projeto=…).

    GET  /api/plugin/projetos              (TASK-131) projetos cadastrados (o plugin lista todos na tela inicial).
    GET  /api/plugin/leitura?projeto=081   (TASK-131) configuração de LEITURA do desenho do projeto (layers da rede, blocos de poste,
    PUT  /api/plugin/leitura?projeto=081   guias, raios) guardada aqui para valer em qualquer computador; também vai em /regras e na versão.
    GET  /api/plugin/catalogo?projeto=081  (TASK-144) a Base de Orçamento do projeto para o autocompletar do plugin: nomes (+ descrição, origem),
                                           apelidos e os parâmetros de cada ativo (mesma lógica da dica da aplicação, services/ativo_parametros).
    GET  /api/plugin/desktop               (TASK-143, pública) "sou o OrçaCore (antes GridReader) instalado neste computador?"
    POST /api/plugin/receber               (TASK-143, pública, só 127.0.0.1 e só no programa instalado) o plugin entrega a obra; a janela
                                           vem para a frente em /resumo?plugin_pendente=<id> · GET /api/plugin/pendente/<id> (login) a devolve.
    GET  /api/plugin/instalacao            (TASK-131) o programa traz o plugin: versão trazida × instalada e caminho para o APPLOAD
    POST /api/plugin/instalacao            (desktop) instalar/atualizar agora · POST /api/plugin/instalacao/abrir-pasta (desktop)
    GET  /api/plugin/atualizacao?alvo=nanocad26          (TASK-130) versão do plugin publicada aqui + SHA-256 + endereço do arquivo.
    GET  /api/plugin/atualizacao/arquivo?alvo=nanocad26  a DLL. As duas são PÚBLICAS (a DLL não tem dados): o plugin se atualiza
                                           sozinho mesmo sem login. A DLL servida é a do repositório (plugin_nanocad/appload/<pasta>/),
                                           com o manifesto versao.json que o compilar.sh grava; mesclou no main e publicou = atualização.

As quatro primeiras exigem login (o middleware protege /api/*); a 4ª grava só a obra do próprio usuário. Não registrar o conteúdo das obras em log (RULES Regra 11).
"""
import hashlib
import json
import math
from datetime import datetime
from urllib.parse import quote

import os

from fastapi import APIRouter, HTTPException, Request
from fastapi.responses import FileResponse
from pydantic import BaseModel
from typing import Any, Dict, Optional

import database
from services import obras_arquivo

router = APIRouter(prefix="/api/plugin", tags=["plugin-cad"])

LIMITE_LINHAS = obras_arquivo.LIMITE_LINHAS


def _projeto(codigo: str) -> dict:
    conn = database.get_row_connection()
    try:
        r = conn.execute("SELECT nome, codigo FROM projetos WHERE codigo = ?", ((codigo or "").strip(),)).fetchone()
    finally:
        conn.close()
    if not r:
        raise HTTPException(status_code=404, detail=f"Projeto '{codigo}' não cadastrado.")
    return {"codigo": r["codigo"], "nome": r["nome"]}


def _user_id(request: Request):
    user = getattr(request.state, "user", None)
    return user["user_id"] if user else None


def _fontes(codigo: str, user_id) -> dict:
    """Tudo que muda o resultado do plugin, das MESMAS fontes que a tela e o modo autônomo usam."""
    from routers.regras import get_regras_conversao
    from routers.regras_leitor import _get_regras
    from routers.vinculacao_config import obter_config_com_origem
    from services.sync_service import get_merged_orcamento
    from services.autonomo.totalizadora import REGRA_PADRAO
    cfg, origem = obter_config_com_origem(codigo)
    return {
        "processamento": _get_regras("processamento", codigo),
        "classificacao": _get_regras("classificacao", codigo),
        "vinculacao": cfg,
        "vinculacao_origem": origem,
        "conversao": (get_regras_conversao(codigo) or {}).get("regras") or [dict(REGRA_PADRAO)],
        "base": get_merged_orcamento(user_id),
        "leitura": leitura_salva(codigo),
    }


def _hash(valor) -> str:
    return hashlib.sha256(json.dumps(valor, ensure_ascii=False, sort_keys=True, default=str).encode("utf-8")).hexdigest()[:16]


def _versoes(f: dict) -> dict:
    partes = {"regras": _hash([f["processamento"], f["classificacao"], f["vinculacao"], f.get("leitura")]),
              "conversao": _hash(f["conversao"]), "base": _hash(f["base"])}
    partes["geral"] = _hash(partes)
    return partes


@router.get("/versao")
def versao(request: Request, projeto: str):
    """Leve o bastante para o plugin consultar a cada poucos segundos: só devolve as impressões digitais."""
    p = _projeto(projeto)
    return {"projeto": p, "versao": _versoes(_fontes(p["codigo"], _user_id(request)))}


@router.get("/regras")
def regras(request: Request, projeto: str):
    p = _projeto(projeto)
    f = _fontes(p["codigo"], _user_id(request))
    return {"projeto": p, "versao": _versoes(f), "processamento": f["processamento"], "classificacao": f["classificacao"],
            "vinculacao": f["vinculacao"], "vinculacao_origem": f["vinculacao_origem"], "leitura": f["leitura"]}


@router.get("/projetos")
def projetos():
    """Todos os projetos cadastrados, para a tela inicial do plugin (com ou sem configuração de leitura guardada)."""
    conn = database.get_row_connection()
    try:
        rows = conn.execute("SELECT codigo, nome FROM projetos WHERE codigo IS NOT NULL AND codigo <> '' ORDER BY nome").fetchall()
    finally:
        conn.close()
    return {"projetos": [{"codigo": r["codigo"], "nome": r["nome"], "tem_leitura": leitura_salva(r["codigo"]) is not None} for r in rows]}


# ── configuração de leitura do desenho por projeto (TASK-131) ──
_LISTAS_LEITURA = ("layers_rede_mt", "layers_rede_bt", "blocos_poste", "layers_poste", "layers_guia")
_NUMEROS_LEITURA = ("raio_alturas", "guia_max_alturas", "tol_poste_alturas", "guia_toque_alturas", "raio_reserva_alturas")


def validar_leitura(cfg) -> dict:
    """Só os campos que o plugin entende (ConfigLeitura.ParaJson), com tipos e limites; erro → 400 com a lista."""
    import re
    if not isinstance(cfg, dict):
        raise HTTPException(status_code=400, detail={"erros": ["A configuração de leitura tem de ser um objeto."]})
    erros, ok = [], {}
    for k in _LISTAS_LEITURA:
        if k not in cfg:
            continue
        v = cfg[k]
        if not isinstance(v, list) or len(v) > 50 or any(not isinstance(x, str) or len(x) > 120 for x in v):
            erros.append(f"{k}: lista de até 50 textos")
        else:
            ok[k] = [x.strip() for x in v if x.strip()]
    for k in _NUMEROS_LEITURA:
        if k not in cfg:
            continue
        v = cfg[k]
        if isinstance(v, bool) or not isinstance(v, (int, float)) or not (0 < v <= 1000):
            erros.append(f"{k}: número entre 0 e 1000")
        else:
            ok[k] = v
    if "cor_confirma" in cfg:
        if not isinstance(cfg["cor_confirma"], bool):
            erros.append("cor_confirma: verdadeiro/falso")
        else:
            ok["cor_confirma"] = cfg["cor_confirma"]
    if "placa_regex" in cfg:
        v = cfg["placa_regex"]
        try:
            if not isinstance(v, str) or len(v) > 200:
                raise ValueError
            re.compile(v)
            ok["placa_regex"] = v
        except (ValueError, re.error):
            erros.append("placa_regex: expressão regular inválida")
    if erros:
        raise HTTPException(status_code=400, detail={"erros": erros})
    return ok


def leitura_salva(projeto_codigo: str):
    """Configuração de leitura guardada para o projeto (nuvem primeiro, depois local) ou None."""
    from services.supabase_client import get_supabase
    supabase = get_supabase()
    if supabase:
        try:
            res = supabase.table("plugin_leitura_config").select("config_json").eq("projeto_codigo", projeto_codigo).execute()
            if res.data:
                return json.loads(res.data[0]["config_json"])
        except Exception as e:
            from config import logger
            logger.warning(f"Falha ao buscar plugin_leitura_config no Supabase: {e}")
    conn = database.get_connection()
    try:
        row = conn.execute("SELECT config_json FROM plugin_leitura_config WHERE projeto_codigo = ?", (projeto_codigo,)).fetchone()
    finally:
        conn.close()
    if row:
        try:
            v = json.loads(row[0])
            return v if isinstance(v, dict) else None
        except ValueError:
            return None
    return None


class LeituraPayload(BaseModel):
    leitura: Dict[str, Any]


@router.get("/leitura")
def obter_leitura(projeto: str):
    p = _projeto(projeto)
    return {"projeto": p, "leitura": leitura_salva(p["codigo"])}


@router.put("/leitura")
def salvar_leitura(projeto: str, payload: LeituraPayload):
    """Grava a configuração de leitura do projeto (local + nuvem). Falha na nuvem não perde a gravação local: vira aviso."""
    from services.supabase_client import get_supabase
    p = _projeto(projeto)
    cfg = validar_leitura(payload.leitura)
    valor = json.dumps(cfg, ensure_ascii=False)
    conn = database.get_connection()
    try:
        conn.execute("INSERT OR REPLACE INTO plugin_leitura_config (projeto_codigo, config_json, updated_at) VALUES (?, ?, datetime('now'))",
                     (p["codigo"], valor))
        conn.commit()
    finally:
        conn.close()
    aviso = None
    supabase = get_supabase()
    if supabase:
        try:
            supabase.table("plugin_leitura_config").upsert({"projeto_codigo": p["codigo"], "config_json": valor}).execute()
        except Exception as e:
            aviso = f"Gravada neste servidor, mas não na nuvem (Supabase): {e}"
    return {"projeto": p, "leitura": cfg, "aviso": aviso}


class ConferirPayload(BaseModel):
    projeto: str
    dados: Dict[str, Any] = {}
    incluir_totalizadora: Optional[bool] = False


def _qtd(q):
    if isinstance(q, float) and (math.isnan(q) or math.isinf(q)):
        return None
    return q


@router.post("/conferir")
def conferir(payload: ConferirPayload, request: Request):
    """Para cada linha de Cabos/Outros: os ativos que ela gera na Totalizadora e quais NÃO existem na Base de Orçamento do projeto."""
    from services.autonomo.totalizadora import montar_totalizadora, normalizar_para_json, payload_calculo
    from services.autonomo.vinculacao import itens_vinculo_para_totalizadora
    from services.orcamento_calc import processar_calculo
    from services.vinculacao_regras import normalizar_config

    p = _projeto(payload.projeto)
    dados = obras_arquivo.normalizar_dados(payload.dados if isinstance(payload.dados, dict) else {})
    cabos, outros = dados["cabos"], dados["outros"]
    if len(cabos) > LIMITE_LINHAS or len(outros) > LIMITE_LINHAS:
        raise HTTPException(status_code=413, detail=f"Linhas demais (máx. {LIMITE_LINHAS} por tabela).")
    f = _fontes(p["codigo"], _user_id(request))
    tot = montar_totalizadora(cabos, outros, f["conversao"], f["base"], p["nome"],
                              itens_vinculo=itens_vinculo_para_totalizadora(cabos, outros), cfg_vinculo=normalizar_config(f["vinculacao"]))
    calc = payload_calculo(tot)
    nao_encontrados = list(processar_calculo(calc["cabos"], calc["outros"], p["nome"], f["base"])["nao_encontrados"])

    def nao_achado(ativo):
        a = str(ativo or "")
        return any(n == a or n.startswith(a + " (") for n in nao_encontrados)

    # TOT-k → linha de origem: o contador da Totalizadora anda uma vez por linha não vazia, Cabos e depois Outros
    origem, k = {}, 1
    for tabela, linhas in (("cabos", cabos), ("outros", outros)):
        for i, linha in enumerate(linhas):
            if linha:
                origem[f"TOT-{k}"] = (tabela, i)
                k += 1
    por_linha = {}
    for item in tot:
        chave = origem.get(item.get("id") or item.get("baseId"))
        if chave is None:
            continue
        # naoEncontrado = o que a Totalizadora da tela pinta de vermelho (sem linha do projeto nem genérica na base);
        # foraDoCalculo = nem a cascata do cálculo (RN-06/RN-08, que ainda tenta outros projetos) achou — some do orçamento
        fora = nao_achado(item.get("ativo"))
        por_linha.setdefault(chave, []).append({"ativo": item.get("ativo"), "operacao": item.get("operacao"), "qtd": _qtd(item.get("qtd")),
                                                "origem": item.get("origem"), "naoEncontrado": bool(item.get("naoEncontrado")) or fora,
                                                "foraDoCalculo": fora, "desc": item.get("desc") or ""})
    linhas = [{"tabela": t, "indice": i, "itens": itens} for (t, i), itens in sorted(por_linha.items(), key=lambda x: (x[0][0] != "cabos", x[0][1]))]
    resposta = {"projeto": p, "versao": _versoes(f), "linhas": linhas, "nao_encontrados": sorted(set(nao_encontrados)),
                "resumo": {"linhas": len(linhas), "com_pendencia": sum(1 for l in linhas if any(x["naoEncontrado"] for x in l["itens"]))}}
    if payload.incluir_totalizadora:
        resposta["totalizadora"] = normalizar_para_json(tot)
    return resposta


class ObraPluginPayload(BaseModel):
    chave: str                      # identifica o desenho (o plugin manda o caminho do DWG; aqui só entra no hash do id)
    envelope: Dict[str, Any]        # o mesmo conteúdo do .obra.json


@router.post("/obra")
def enviar_obra(payload: ObraPluginPayload, request: Request):
    """Grava (cria ou atualiza) a obra do desenho para o usuário e devolve o link que a abre no Resumo (TASK-106)."""
    from routers.obras import _existente, gravar_obra
    from middleware.auth_middleware import get_current_user_from_state
    user = get_current_user_from_state(request)
    env = payload.envelope if isinstance(payload.envelope, dict) else {}
    p = _projeto(str(env.get("projeto") or ""))
    if len(json.dumps(env, ensure_ascii=False)) > obras_arquivo.LIMITE_BYTES:
        raise HTTPException(status_code=413, detail={"erros": ["Obra grande demais."]})
    try:
        pronta = obras_arquivo.validar_importacao(env, p["codigo"])
    except obras_arquivo.ErroArquivoObra as e:
        raise HTTPException(status_code=400, detail={"erros": e.mensagens})
    if pronta["tipo"] != "obra":
        raise HTTPException(status_code=400, detail={"erros": ["O plugin só envia obras (não modelos)."]})
    chave = (payload.chave or "").strip()
    if not chave:
        raise HTTPException(status_code=400, detail={"erros": ["Falta a chave do desenho."]})
    obra_id = "plugin_" + hashlib.sha256(f"{user['user_id']}|{chave.lower()}".encode("utf-8")).hexdigest()[:20]
    existe, dono, _, _ = _existente(obra_id)
    if existe and dono and dono != user["user_id"]:
        raise HTTPException(status_code=403, detail={"erros": ["Essa obra pertence a outro usuário."]})
    nome = (pronta["nome"] + " (nanoCAD)")[:120]
    gravar_obra(user, {"id": obra_id, "nome": nome, "data": datetime.now().strftime("%d/%m/%Y, %H:%M:%S"),
                       "dados_json": obras_arquivo.dados_json_para_gravar(pronta["dados"]), "projeto": p["codigo"], "publica": False, "tipo": "obra"})
    url = f"/resumo?obra={quote(obra_id)}&projeto={quote(p['codigo'])}&projeto_nome={quote(p['nome'])}"
    return {"id": obra_id, "nome": nome, "projeto": p, "url": url, "atualizada": existe,
            "cabos": len(pronta["dados"]["cabos"]), "outros": len(pronta["dados"]["outros"])}


# ── atualização automática do plugin (TASK-130) ──
def _pasta_appload() -> str:
    from config import _get_resource_path        # dev: pasta do repositório · .exe: empacotada pelo build.bat
    return _get_resource_path(os.path.join("plugin_nanocad", "appload"))


PASTA_APPLOAD = _pasta_appload()
ALVOS_PLUGIN = {"nanocad26": "nanoCAD26", "nanocad5": "nanoCAD5"}
ARQUIVO_DLL = "OrcamentoNanoCAD.dll"
_cache_sha: Dict[str, Any] = {}


def _pacote(alvo: str) -> dict:
    """Versão publicada para o alvo, conferida: o SHA-256 do manifesto tem de ser o do arquivo (senão o manifesto está velho)."""
    pasta = ALVOS_PLUGIN.get((alvo or "").strip().lower())
    if not pasta:
        raise HTTPException(status_code=404, detail="Alvo desconhecido.")
    dll = os.path.join(PASTA_APPLOAD, pasta, ARQUIVO_DLL)
    manifesto = os.path.join(PASTA_APPLOAD, pasta, "versao.json")
    if not (os.path.isfile(dll) and os.path.isfile(manifesto)):
        raise HTTPException(status_code=404, detail="Nenhuma versão do plugin publicada para esse alvo.")
    st = os.stat(dll)
    chave = (dll, st.st_mtime, st.st_size)
    if _cache_sha.get(dll, (None,))[0] != chave:
        with open(dll, "rb") as f:
            _cache_sha[dll] = (chave, hashlib.sha256(f.read()).hexdigest())
    sha = _cache_sha[dll][1]
    with open(manifesto, encoding="utf-8") as f:
        m = json.load(f)
    if str(m.get("sha256", "")).lower() != sha:
        raise HTTPException(status_code=503, detail="Manifesto do plugin desatualizado (versao.json não confere com a DLL).")
    return {"alvo": alvo.lower(), "versao": str(m.get("versao", "")), "sha256": sha, "tamanho": st.st_size, "dll": dll}


@router.get("/atualizacao")
def atualizacao_plugin(alvo: str = "nanocad26"):
    """Versão do plugin publicada nesta aplicação (pública: o plugin se atualiza sem login)."""
    p = _pacote(alvo)
    return {"alvo": p["alvo"], "versao": p["versao"], "sha256": p["sha256"], "tamanho": p["tamanho"],
            "url": f"/api/plugin/atualizacao/arquivo?alvo={quote(p['alvo'])}"}


@router.get("/atualizacao/arquivo")
def atualizacao_arquivo(alvo: str = "nanocad26"):
    p = _pacote(alvo)
    return FileResponse(p["dll"], media_type="application/octet-stream", filename=ARQUIVO_DLL,
                        headers={"X-Plugin-Versao": p["versao"], "X-Plugin-Sha256": p["sha256"], "Cache-Control": "no-store"})


# ── o programa principal traz o plugin (TASK-131): situação, instalar agora, abrir a pasta ──
def _desktop() -> bool:
    from config import APP_MODE
    return APP_MODE == "desktop"


@router.get("/instalacao")
def instalacao():
    """Desktop: versão que o programa traz × instalada e o caminho para o APPLOAD. Servidor: só os links de download."""
    from services import plugin_instalador
    downloads = []
    for alvo, (_, nome) in plugin_instalador.ALVOS.items():
        try:
            p = _pacote(alvo)
            downloads.append({"alvo": alvo, "nome": nome, "versao": p["versao"], "url": f"/api/plugin/atualizacao/arquivo?alvo={alvo}"})
        except HTTPException:
            pass
    r = {"desktop": _desktop(), "downloads": downloads}
    if r["desktop"]:
        r["pasta"] = plugin_instalador.pasta_destino()
        r["alvos"] = plugin_instalador.situacao()
    return r


class InstalarPayload(BaseModel):
    forcar: Optional[bool] = False


@router.post("/instalacao")
def instalar_plugin(payload: InstalarPayload):
    if not _desktop():
        raise HTTPException(status_code=403, detail="Só no programa instalado no computador (no online, baixe a DLL).")
    from services import plugin_instalador
    return {"resultado": plugin_instalador.instalar(bool(payload.forcar)), "alvos": plugin_instalador.situacao()}


@router.post("/instalacao/abrir-pasta")
def abrir_pasta_plugin():
    if not _desktop():
        raise HTTPException(status_code=403, detail="Só no programa instalado no computador.")
    from services import plugin_instalador
    pasta = plugin_instalador.pasta_destino()
    os.makedirs(pasta, exist_ok=True)
    if hasattr(os, "startfile"):
        os.startfile(pasta)  # noqa: S606 — pasta fixa do próprio programa
        return {"aberta": True, "pasta": pasta}
    return {"aberta": False, "pasta": pasta}



# ── o plugin entrega a obra ao programa aberto no computador (TASK-143) ──
@router.get("/desktop")
def desktop():
    from config import APP_VERSION
    # "app" continua "GridReader": é o identificador que o plugin do nanoCAD já instalado confere (TASK-152: o nome visível passou a ser GridR e depois OrçaCore; "nome" é o novo)
    return {"app": "GridReader", "nome": "OrçaCore", "desktop": _desktop(), "versao": APP_VERSION}


class ReceberPayload(BaseModel):
    chave: str = ""
    envelope: Dict[str, Any]


def _do_proprio_computador(request: Request) -> bool:
    host = (request.client.host if request.client else "") or ""
    return host in ("127.0.0.1", "::1", "localhost", "testclient")


@router.post("/receber")
def receber_obra(payload: ReceberPayload, request: Request):
    """Sem login (o plugin não tem o token deste programa), mas só no programa instalado e só do próprio computador:
    a obra fica em memória até a janela — com o usuário logado — buscá-la e gravá-la para ele."""
    if not _desktop():
        raise HTTPException(status_code=403, detail="Só no programa instalado no computador.")
    if not _do_proprio_computador(request):
        raise HTTPException(status_code=403, detail="Só o plugin deste computador pode entregar obras.")
    from services import janela_desktop, programa_local
    erro = programa_local.validar(payload.chave, payload.envelope)
    if erro:
        raise HTTPException(status_code=400, detail=erro)
    i = programa_local.guardar(payload.chave, payload.envelope)
    url = programa_local.url_pendente(i)
    janela = janela_desktop.mostrar(url, request.url.port)
    return {"id": i, "url": url, "janela": janela}


@router.get("/pendente/{i}")
def obra_pendente(i: str):
    from services import programa_local
    v = programa_local.pegar(i)
    if not v:
        raise HTTPException(status_code=404, detail="Obra do nanoCAD não encontrada (passou mais de 1 hora ou o programa foi reaberto). Envie de novo pelo plugin.")
    return {"chave": v["chave"], "envelope": v["envelope"]}



# ── catálogo para o autocompletar do plugin (TASK-144) ──
def montar_catalogo(linhas, projeto: str) -> dict:
    """Mesmo conteúdo que a dica de ativos da aplicação usa (static/ativo_dica.js): nomes válidos no projeto, apelidos e parâmetros."""
    from services import ativo_parametros as ap
    p = str(projeto or "").strip().upper()
    validas = []
    for l in linhas or []:
        lp = str(l.get("projeto") or "").strip().upper()
        if not lp or not p or p in [x.strip() for x in lp.split("/")]:
            validas.append(l)
    ativos, vistos = [], {}
    for l in validas:
        cel = str(l.get("ativo") or "").strip()
        if not cel:
            continue
        a = ap.analisar_celula_base(cel)
        if a.get("erro") or not a.get("base"):
            continue
        nome = a["base"].upper()
        origem = str(l.get("origem") or "").strip().upper() or None
        if nome in vistos:
            item = ativos[vistos[nome]]
            if not item["desc"] and l.get("desc_ativo"):
                item["desc"] = str(l.get("desc_ativo"))
            continue
        vistos[nome] = len(ativos)
        ativos.append({"nome": nome, "desc": str(l.get("desc_ativo") or ""), "origem": origem})
    apelidos = {str(k).upper(): str(v).upper() for k, v in (ap.apelidos_do_projeto(linhas or [], projeto or "") or {}).items()}
    parametros = {}
    for base, sig in (ap.assinaturas_do_projeto(linhas or [], projeto or "") or {}).items():
        nomes = sig.get("nomes") or []
        livres = sig.get("livres") or []
        if not nomes and not livres:
            continue
        lista = [{"nome": ap.nome_tela(n), "padrao": (sig.get("padroes") or {}).get(n), "valores": list((sig.get("valores") or {}).get(n) or []),
                  "mascaras": list((sig.get("mascaras") or {}).get(n) or []), "aberto": n not in (sig.get("fechados") or []), "porNome": False} for n in nomes]
        lista += [{"nome": ap.nome_tela(n), "padrao": None, "valores": [], "mascaras": [], "aberto": True, "porNome": True} for n in livres]
        parametros[str(base).upper()] = lista
    corpo = {"ativos": ativos, "apelidos": apelidos, "parametros": parametros}
    corpo["versao"] = _hash(corpo)
    return corpo


@router.get("/catalogo")
def catalogo(request: Request, projeto: str):
    from services.sync_service import get_merged_orcamento
    p = _projeto(projeto)
    return dict(montar_catalogo(get_merged_orcamento(_user_id(request)), p["nome"]), projeto=p)     # a base guarda o NOME do projeto
