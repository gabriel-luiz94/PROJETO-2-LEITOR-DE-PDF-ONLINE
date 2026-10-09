"""
routers/obras.py — Rotas para gerenciamento de obras.
"""
import json
import logging
import secrets
import time
from datetime import datetime
from typing import Optional
from urllib.parse import quote

from fastapi import APIRouter, HTTPException, Request
from fastapi.responses import JSONResponse
from pydantic import BaseModel
from database import get_connection
from models import ObraModel
from middleware.auth_middleware import get_current_user_from_state
from services import modelos_obra, obras_arquivo

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


def _tipo_de(linha: dict) -> str:
    """'modelo' ou 'obra'. Nuvem sem a coluna `tipo` (SQL da TASK-058 não rodado): reconhece o modelo pela marca em `dados_json`."""
    if linha.get("tipo") in ("obra", "modelo"):
        return linha["tipo"]
    dj = linha.get("dados_json")
    return "modelo" if isinstance(dj, str) and '"modelo"' in dj and '"parametros"' in dj else "obra"


def _apresentar(linha: dict, user) -> dict:
    """Linha do banco -> resposta da API. Nunca expõe `user_id`/e-mail de outros; só o nome curto do dono."""
    uid = _uid(user)
    dono = linha.get("user_id") or None
    saida = {k: linha[k] for k in ("id", "nome", "data", "projeto", "dados_json") if k in linha}
    saida["tipo"] = _tipo_de(linha)
    saida["publica"] = bool(linha.get("publica")) or saida["tipo"] == "modelo"
    saida["minha"] = dono is not None and dono == uid
    saida["sem_dono"] = dono is None
    if saida["minha"]:
        saida["dono"] = linha.get("dono_nome") or _nome_curto((user or {}).get("email")) or "você"
    elif dono is None:
        saida["dono"] = ""
    else:
        saida["dono"] = linha.get("dono_nome") or "outro usuário"
    return saida


_COLS_LEVES = "id, nome, data, projeto, user_id, publica, dono_nome, tipo"
_COLS_LEVES_057 = "id, nome, data, projeto, user_id, publica, dono_nome"
_COLS_LEVES_ANTIGAS = "id, nome, data, projeto, user_id"   # nuvem ainda sem as colunas novas (SQL da TASK-057 não rodado)


def _nuvem(sb, completo, montar):
    if completo:
        return montar(sb.table("obras").select("*")).execute().data or []
    for cols in (_COLS_LEVES, _COLS_LEVES_057):      # nuvem sem `tipo` (TASK-058) e/ou sem `publica` (TASK-057)
        try:
            return montar(sb.table("obras").select(cols)).execute().data or []
        except Exception:
            continue
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
        cols = "id, nome, data, projeto, user_id, publica, dono_nome, tipo" + (", dados_json" if completo else "")
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
def get_obras(request: Request, projeto: str = None, visibilidade: str = "todas", origem: str = "todas", tipo: str = "todos"):
    """Obras visíveis ao usuário (TASK-057): as dele + as públicas de outros no `projeto`. Filtros: visibilidade
    (todas|particulares|publicas) e origem (todas|minhas|outros). Sem `projeto` só vêm as dele."""
    user = get_current_user_from_state(request)
    itens = [_apresentar(l, user) for l in _ler_linhas(user, projeto, completo=True)]
    if visibilidade == "publicas":
        itens = [o for o in itens if o["publica"]]
    elif visibilidade == "particulares":
        itens = [o for o in itens if not o["publica"]]
    if tipo == "obras":
        itens = [o for o in itens if o["tipo"] == "obra"]
    elif tipo == "modelos":
        itens = [o for o in itens if o["tipo"] == "modelo"]
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
    modelo = pronta["tipo"] == "modelo"
    dados_json = (obras_arquivo.dados_json_para_gravar(pronta["dados"], pronta["parametros"]) if modelo
                  else obras_arquivo.dados_json_para_gravar(pronta["dados"]))
    if modelo:      # o modelo importado vira modelo de quem importa; os rótulos/padrões vêm do arquivo (TASK-058)
        dados_json, _ = _preparar_modelo(dados_json, pronta["parametros"], None, user, pronta["projeto"])
    gravar_obra(user, {"id": obra_id, "nome": nome, "data": datetime.now().strftime("%d/%m/%Y, %H:%M:%S"),
                       "dados_json": dados_json, "projeto": pronta["projeto"], "publica": modelo, "tipo": pronta["tipo"]})
    return {"status": "success", "id": obra_id, "nome": nome, "projeto": pronta["projeto"], "tipo": pronta["tipo"],
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
    """(existe, user_id, publica, tipo) da obra `obra_id` em qualquer dono — para proteger contra sobrescrita (TASK-057)."""
    conn = get_connection()
    try:
        r = conn.execute("SELECT user_id, publica, tipo FROM obras WHERE id = ?", (obra_id,)).fetchone()
    finally:
        conn.close()
    if r:
        return True, (r[0] or None), bool(r[1]), (r[2] or "obra")
    sb = get_supabase()
    if sb:
        try:
            d = sb.table("obras").select("id, user_id").eq("id", obra_id).execute().data
            if d:
                return True, (d[0].get("user_id") or None), False, "obra"
        except Exception as e:
            registrar_falha("obras.existente", e)
    return False, None, False, "obra"


def gravar_obra(user, registro: dict) -> None:
    """Grava a obra do `user` (Supabase quando disponível + SQLite). `registro['publica']` (bool, padrão falso).
    Usado por `save_obra`, pela importação e pelo modo autônomo."""
    uid = _uid(user)
    tipo = registro.get("tipo") if registro.get("tipo") in ("obra", "modelo") else "obra"
    publica = bool(registro.get("publica")) or tipo == "modelo"      # modelo é sempre público no projeto (TASK-058)
    dono_nome = _nome_curto((user or {}).get("email")) or None
    base = {"id": registro["id"], "nome": registro["nome"], "data": registro["data"], "dados_json": registro["dados_json"],
            "user_id": uid, "projeto": registro["projeto"]}
    supabase = get_supabase()
    if supabase:
        try:
            for extra in ({"publica": publica, "dono_nome": dono_nome, "tipo": tipo}, {"publica": publica, "dono_nome": dono_nome}, {}):
                try:      # nuvem sem as colunas novas (SQL das TASK-057/058 não rodado): recua até gravar como antes
                    supabase.table("obras").upsert({**base, **extra}).execute()
                    break
                except Exception:
                    if not extra:
                        raise
        except Exception as e:
            registrar_falha("obras.save_obra", e)

    conn = get_connection()
    cursor = conn.cursor()
    cursor.execute("INSERT OR REPLACE INTO obras (id, nome, data, dados_json, user_id, projeto, publica, dono_nome, tipo) VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?)",
                   (registro["id"], registro["nome"], registro["data"], registro["dados_json"], uid, registro["projeto"], int(publica), dono_nome, tipo))
    conn.commit()
    conn.close()


def _contexto_modelo(user, projeto):
    """(grupos do projeto, universo de ativos da base técnica) — para expandir `X(@GRUPO)` e avisar de opção fora da base (TASK-060). Nunca levanta."""
    grupos, universo = {}, set()
    try:
        from routers.validacao_regras import regras_efetivas
        grupos = regras_efetivas(projeto or "DEFAULT")[1] or {}
    except Exception as e:  # noqa: BLE001
        logging.getLogger(__name__).warning("Grupos do projeto indisponíveis: %s", type(e).__name__)
    try:
        from services.sync_service import get_merged_orcamento
        conn = get_connection()
        try:
            r = conn.execute("SELECT nome FROM projetos WHERE codigo = ?", (projeto,)).fetchone()
        finally:
            conn.close()
        nome = (r[0] if r else str(projeto or "")).strip().upper()
        for linha in get_merged_orcamento(_uid(user)):
            projs = [p.strip() for p in (linha.get("projeto") or "").strip().upper().split("/") if p.strip()]
            ativo = (linha.get("ativo") or "").strip().upper()
            if ativo and (not projs or nome in projs):
                universo.add(ativo)
    except Exception as e:  # noqa: BLE001
        logging.getLogger(__name__).warning("Base técnica indisponível para conferir opções: %s", type(e).__name__)
    return grupos, universo


def _preparar_modelo(dados_json, parametros, atuais_json=None, user=None, projeto=None):
    """Valida e normaliza o `dados_json` de um modelo (TASK-058/060): precisa ter ao menos uma variável (`V` ou `X(...)`); sem Totalizadora;
    guarda os parâmetros. Devolve (dados_json a gravar, avisos) — os avisos (ex.: opção fora da base técnica) não impedem salvar."""
    try:
        snap = json.loads(dados_json) if isinstance(dados_json, str) else None
    except ValueError:
        snap = None
    if not isinstance(snap, dict):
        raise HTTPException(status_code=400, detail={"erros": ["Os dados do modelo estão ilegíveis."]})
    cabos, outros, cfg = modelos_obra.separar_dados(snap)
    if len(cabos) > obras_arquivo.LIMITE_LINHAS or len(outros) > obras_arquivo.LIMITE_LINHAS:
        raise HTTPException(status_code=400, detail={"erros": ["Linhas demais no modelo."]})
    variaveis = modelos_obra.detectar_variaveis(cabos, outros)
    if not variaveis:
        raise HTTPException(status_code=400, detail={"erros": ["O modelo precisa ter ao menos uma variável (ex.: 'CAA 2 ABC V m', 'V-U4' ou 'V-X(U3,U4)')."]})
    try:
        if parametros is None:       # não enviados: mantém os já gravados (ou os do próprio dados_json)
            cfg_final = cfg or (modelos_obra.separar_dados(atuais_json)[2] if atuais_json else [])
        else:
            cfg_final = modelos_obra.validar_configuracao(parametros)
    except modelos_obra.ErroModelo as e:
        raise HTTPException(status_code=400, detail={"erros": e.mensagens})
    grupos, universo = _contexto_modelo(user, projeto)
    analise = modelos_obra.analisar(cabos, outros, cfg_final, grupos, universo)
    if analise["erros"]:
        raise HTTPException(status_code=400, detail={"erros": analise["erros"]})
    params = [{"chave": p["chave"], "rotulo": p["rotulo"], "padrao": p["padrao"]} for p in analise["parametros"]]
    return obras_arquivo.dados_json_para_gravar({"cabos": cabos, "outros": outros}, params), analise["avisos"]


def _dados_atuais(user, obra_id):
    o = buscar_obra(user, obra_id)
    return o.get("dados_json") if o else None


@router.post("")
def save_obra(obra: ObraModel, request: Request):
    user = get_current_user_from_state(request)
    existe, dono, publica_atual, tipo_atual = _existente(obra.id)
    if existe and dono != _uid(user):
        raise HTTPException(status_code=403, detail="Esta obra pertence a outro usuário. Salve como uma nova obra.")
    if obra.tipo not in (None, "obra", "modelo"):
        raise HTTPException(status_code=400, detail="Tipo inválido (use obra ou modelo).")
    tipo = obra.tipo or (tipo_atual if existe else "obra")
    publica = obra.publica if obra.publica is not None else (publica_atual if existe else False)
    dados_json, avisos = obra.dados_json, []
    if tipo == "modelo":
        dados_json, avisos = _preparar_modelo(dados_json, obra.parametros, _dados_atuais(user, obra.id) if existe else None, user, obra.projeto)
        publica = True
    gravar_obra(user, {"id": obra.id, "nome": obra.nome, "data": obra.data, "dados_json": dados_json,
                       "projeto": obra.projeto, "publica": publica, "tipo": tipo})
    if (obra.origem_execucao or obra.baseline_manual) and tipo != "modelo":
        _registrar_aprendizado(user, obra)
    return {"status": "success", "publica": bool(publica), "tipo": tipo, **({"avisos": avisos} if avisos else {})}


def _registrar_aprendizado(user, obra) -> None:
    """TASK-059: obra do modo autônomo corrigida e salva -> guarda só as diferenças e atualiza as propostas. Nunca atrapalha o salvar."""
    if (user or {}).get("role") != "admin":      # as propostas são do administrador (viram regras/ajustes do projeto)
        return
    try:
        from services.autonomo import aprendizado_servico
        aprendizado_servico.registrar_sessao(_uid(user), obra.projeto, obra.origem_execucao, obra.baseline_manual, obra.dados_json, obra.id)
    except Exception as e:  # noqa: BLE001 — o aprendizado é um extra
        logging.getLogger(__name__).warning("Aprendizado do autônomo não registrado: %s", type(e).__name__)


class DetectarModel(BaseModel):
    cabos: list = []
    outros: list = []
    projeto: Optional[str] = None
    parametros: Optional[list] = None      # rótulos/padrões já escolhidos (para conferir o padrão dos campos de ativo)


@router.post("/modelo/detectar")
def detectar_variaveis_modelo(corpo: DetectarModel, request: Request):
    """Variáveis (`V` e `X(...)`) das tabelas enviadas, para a janela de "Salvar como modelo" — TASK-058/060.
    Devolve `parametros` (campos de ativo com `opcoes` já expandidas), `erros` (impedem salvar) e `avisos` (não impedem)."""
    user = get_current_user_from_state(request)
    if len(corpo.cabos) > obras_arquivo.LIMITE_LINHAS or len(corpo.outros) > obras_arquivo.LIMITE_LINHAS:
        raise HTTPException(status_code=400, detail="Linhas demais.")
    try:
        cfg = modelos_obra.validar_configuracao(corpo.parametros) if corpo.parametros is not None else None
    except modelos_obra.ErroModelo as e:
        raise HTTPException(status_code=400, detail={"erros": e.mensagens})
    grupos, universo = _contexto_modelo(user, corpo.projeto)
    return modelos_obra.analisar(corpo.cabos, corpo.outros, cfg, grupos, universo)


def _modelo_visivel(request: Request, obra_id: str, projeto):
    user = get_current_user_from_state(request)
    obra = buscar_obra(user, obra_id, projeto)
    if obra is None:
        raise HTTPException(status_code=404, detail="Obra não encontrada.")
    if obra["tipo"] != "modelo":
        raise HTTPException(status_code=400, detail="Esta obra não é um modelo.")
    return obra


@router.get("/{obra_id}/modelo")
def parametros_do_modelo(obra_id: str, request: Request, projeto: str = None):
    """Parâmetros do modelo (quantidades e escolhas de ativo, com as opções de HOJE — grupos expandidos agora), para a janela "Parâmetros do modelo" — TASK-058/060."""
    obra = _modelo_visivel(request, obra_id, projeto)
    cabos, outros, cfg = modelos_obra.separar_dados(obra["dados_json"])
    grupos, universo = _contexto_modelo(get_current_user_from_state(request), projeto)
    analise = modelos_obra.analisar(cabos, outros, cfg, grupos, universo)
    return {"id": obra["id"], "nome": obra["nome"], "parametros": analise["parametros"], "erros": analise["erros"]}


class GerarModel(BaseModel):
    valores: dict = {}


@router.post("/{obra_id}/gerar")
def gerar_obra_do_modelo(obra_id: str, corpo: GerarModel, request: Request, projeto: str = None):
    """Gera uma obra PADRÃO (sem V) a partir do modelo e dos valores informados. Não grava nada — TASK-058."""
    obra = _modelo_visivel(request, obra_id, projeto)
    cabos, outros, cfg = modelos_obra.separar_dados(obra["dados_json"])
    try:
        gerada = modelos_obra.gerar(cabos, outros, cfg, corpo.valores, *_contexto_modelo(get_current_user_from_state(request), projeto))
    except modelos_obra.ErroModelo as e:
        raise HTTPException(status_code=400, detail={"erros": e.mensagens})
    return {"nome": obra["nome"], "projeto": obra["projeto"], **gerada}


class VisibilidadeModel(BaseModel):
    publica: bool


@router.put("/{obra_id}/visibilidade")
def alterar_visibilidade(obra_id: str, corpo: VisibilidadeModel, request: Request):
    """Torna a obra pública (visível no projeto) ou particular. Só o dono (TASK-057)."""
    user = get_current_user_from_state(request)
    existe, dono, _, _ = _existente(obra_id)
    if not existe or dono is None:
        raise HTTPException(status_code=404, detail="Obra não encontrada.")
    if dono != _uid(user):
        # quem não enxerga a obra recebe 404 (não revela que existe); quem a enxerga (pública) recebe 403
        visivel = any(l["id"] == obra_id for l in _ler_linhas(user, _projeto_da(obra_id)))
        raise HTTPException(status_code=403 if visivel else 404, detail="Só o dono da obra altera a visibilidade." if visivel else "Obra não encontrada.")
    if _existente(obra_id)[3] == "modelo":
        raise HTTPException(status_code=400, detail="Modelos são sempre públicos no projeto.")
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
    existe, dono, _, _ = _existente(obra_id)
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
