"""
routers/barra_modais.py — Configuração da BARRA DE MODAIS por projeto (TASK-117).

Cada projeto guarda quais modais ficam ocultos na barra do Resumo e a ordem dos botões. Os modais do sistema (resumo-rede, postes, cabos…)
só são ocultados; os montáveis (id "mm:<id>", ver routers/modais_montaveis.py) também podem ser excluídos/copiados entre projetos.
Leitura: qualquer usuário autenticado. Gravação: só administrador. Nuvem (Supabase) primeiro, fallback SQLite; usa a tabela já existente
`atalhos_modais` com modal = "barra_modais" (sem SQL novo).

Formato: {"versao": 1, "ocultos": ["conexoes", "mm:ramais_colados"], "ordem": ["resumo-rede", "ajustar", ...]}  (ordem vazia = padrão)
"""
import re

from fastapi import APIRouter, Depends, HTTPException, Request
from pydantic import BaseModel

from middleware.auth_middleware import require_role
from routers import modais_montaveis as mm

router = APIRouter(prefix="/api/barra-modais", tags=["barra-modais"])

CHAVE = "barra_modais"
MAX_ITENS = 100
_RE_ID = re.compile(r"^[A-Za-z0-9_:-]{1,60}$")


def _vazio() -> dict:
    return {"versao": 1, "ocultos": [], "ordem": []}


def validar(dados) -> list:
    erros = []
    if not isinstance(dados, dict):
        return ["A configuração deve ser um objeto."]
    for campo in ("ocultos", "ordem"):
        lista = dados.get(campo, [])
        if not isinstance(lista, list) or len(lista) > MAX_ITENS:
            erros.append(f"'{campo}' deve ser uma lista de até {MAX_ITENS} ids.")
            continue
        for v in lista:
            if not isinstance(v, str) or not _RE_ID.match(v):
                erros.append(f"'{campo}': id inválido ({str(v)[:30]!r}).")
        if len(set(lista)) != len(lista):
            erros.append(f"'{campo}' tem ids repetidos.")
    return erros


def normalizar(dados) -> dict:
    d = dados if isinstance(dados, dict) else {}
    ok = lambda lista: [x for x in (lista if isinstance(lista, list) else []) if isinstance(x, str) and _RE_ID.match(x)][:MAX_ITENS]
    ocultos = list(dict.fromkeys(ok(d.get("ocultos"))))
    ordem = list(dict.fromkeys(ok(d.get("ordem"))))
    return {"versao": 1, "ocultos": ocultos, "ordem": ordem}


def _ler(projeto: str) -> dict:
    return normalizar(mm._ler_chave(projeto, CHAVE, _vazio()))


class BarraPayload(BaseModel):
    projeto_codigo: str
    dados: dict


class CopiarPayload(BaseModel):
    origem: str
    destino: str
    ids: list[str]
    copiar_ocultos: bool = True


@router.get("")
def obter(projeto_codigo: str = "DEFAULT"):
    return {"dados": _ler(projeto_codigo)}


@router.put("", dependencies=[Depends(require_role("admin"))])
def gravar(payload: BarraPayload):
    erros = validar(payload.dados)
    if erros:
        raise HTTPException(status_code=400, detail={"erros": erros[:20]})
    dados = normalizar(payload.dados)
    mm._escrever_chave(payload.projeto_codigo, CHAVE, mm.json.dumps(dados, ensure_ascii=False), obrigatorio_nuvem=True)
    return {"status": "success", "dados": dados}


def _id_livre(base: str, usados: set) -> str:
    if base not in usados:
        return base
    n = 2
    while True:
        suf = f"_{n}"
        cand = base[:40 - len(suf)] + suf
        if cand not in usados:
            return cand
        n += 1


@router.post("/copiar", dependencies=[Depends(require_role("admin"))])
def copiar(payload: CopiarPayload, request: Request):
    """Copia modais MONTÁVEIS (por id) de um projeto para outro. Id repetido no destino ganha sufixo _2, _3…; os do sistema não se copiam."""
    if payload.origem == payload.destino:
        raise HTTPException(status_code=400, detail="Origem e destino são o mesmo projeto.")
    origem = mm._ler(payload.origem).get("modais", [])
    escolhidos = [m for m in origem if isinstance(m, dict) and m.get("id") in set(payload.ids)]
    if not escolhidos:
        raise HTTPException(status_code=400, detail="Nenhum dos modais escolhidos existe no projeto de origem.")
    destino = mm._ler(payload.destino)
    modais = [m for m in destino.get("modais", []) if isinstance(m, dict)]
    if len(modais) + len(escolhidos) > mm.MAX_MODAIS:
        raise HTTPException(status_code=400, detail=f"O destino ficaria com mais de {mm.MAX_MODAIS} modais.")
    usados = {m.get("id") for m in modais}
    ocultos_origem = set(_ler(payload.origem)["ocultos"])
    novos_ocultos, mapa = [], {}
    for m in escolhidos:
        novo = dict(m)
        novo["id"] = _id_livre(str(m["id"]), usados)
        usados.add(novo["id"])
        mapa[m["id"]] = novo["id"]
        modais.append(novo)
        if payload.copiar_ocultos and f"mm:{m['id']}" in ocultos_origem:
            novos_ocultos.append(f"mm:{novo['id']}")
    erros = mm.validar({"versao": 1, "modais": modais})
    if erros:
        raise HTTPException(status_code=400, detail={"erros": erros[:20]})
    mm._gravar_interno(payload.destino, {"versao": 1, "modais": modais}, mm._email(request))
    if novos_ocultos:
        cfg = _ler(payload.destino)
        cfg["ocultos"] = list(dict.fromkeys(cfg["ocultos"] + novos_ocultos))
        mm._escrever_chave(payload.destino, CHAVE, mm.json.dumps(cfg, ensure_ascii=False), obrigatorio_nuvem=True)
    return {"status": "success", "copiados": [{"de": a, "para": b} for a, b in mapa.items()]}
