"""
services/projeto_obra.py — Arquivo `.obra` do PROJETO DE TRABALHO (TASK-154). Funções puras: sem banco, sem rede, sem janela.

O `.obra` é o mesmo envelope do `.obra.json` (formato "leitor-obra", versão 1: `nome`, `projeto`, `dados` {cabos, outros, totalizadora}),
mais o bloco `projeto_trabalho` com o que só existe quando se trabalha dentro de um projeto:
    {"nome": "...", "numero_obra": "...", "sap": {bairro, municipio, utd, ...}, "resultado": {"assinatura": "...", "linhas": [...]}, "salvo_em": "..."}
Um `.obra`/`.obra.json` ANTIGO (sem esse bloco) abre normalmente: o nome vem do envelope e o resto fica vazio.
TASK-163: o plugin do nanoCAD grava o MESMO arquivo (`<desenho>.obra`, na pasta do desenho) com a seção `cad` (a obra dele: geometria,
vínculos, ajustes). Este programa não usa a `cad`, mas a PRESERVA ao gravar — senão o plugin perderia a obra dele.
Programas antigos abrem o `.obra` novo como uma obra comum (ignoram o bloco extra).
Tudo que vem de arquivo é tratado como não confiável (tamanho, tipos, campos). Não registrar o conteúdo em log.
"""
import hashlib
import json
import re
from datetime import datetime, timezone

from services import obras_arquivo
from services.fundo_cad import fundo_do_cad

EXTENSAO = ".obra"
LIMITE_BYTES = 60 * 1024 * 1024
LIMITE_NOME = 120
LIMITE_NUMERO = 40
LIMITE_TEXTO = 300
LIMITE_CONTEXTO = 80                                 # TASK-165: contexto selecionado (parâmetro OBRA.CONTEXTO)
LIMITE_LINHAS_RESULTADO = 20000
# Campos da caixinha "Informações do relatório SAP" (título e nota vêm do nome e do número da obra)
CAMPOS_SAP = ("sap", "bairro", "municipio", "utd", "utep", "projetista", "data", "id_investimento", "id_investimento_odd")


class ErroProjeto(ValueError):
    """Arquivo ou pedido recusado; `args[0]` = lista de mensagens em português."""

    @property
    def mensagens(self) -> list:
        return self.args[0]


def assinatura(conteudo) -> str:
    """SHA-1 do conteúdo (bytes ou texto): detecta alteração feita por outro programa/computador."""
    if isinstance(conteudo, str):
        conteudo = conteudo.encode("utf-8")
    return hashlib.sha1(conteudo).hexdigest()


def nome_arquivo(nome: str) -> str:
    """`<nome limpo>.obra` (sem caracteres proibidos no Windows)."""
    base = re.sub(r"[^\w\-. ]+", "_", str(nome or "projeto"), flags=re.UNICODE).strip(" ._")[:80] or "projeto"
    return base + EXTENSAO


def _texto(v, limite) -> str:
    return v.strip()[:limite] if isinstance(v, str) else ""


_RE_CHAVE_SAP = re.compile(r"^[a-z][a-z0-9_]{0,30}$")
LIMITE_CAMPOS_SAP = 30


def _chaves_do_perfil(projeto) -> tuple:
    """TASK-176: as chaves do cabeçalho do perfil de exportação do projeto (além das de sempre); sem perfil, nenhuma."""
    codigo = _texto(projeto, 40)
    if not codigo:
        return ()
    try:
        from routers.vinculacao_config import perfil_do_projeto
        from services import perfil_exportacao as pe
        perfil = perfil_do_projeto(codigo)
        return tuple(k for k in pe.campos_sap(perfil) if k not in CAMPOS_SAP) if perfil else ()
    except Exception:  # noqa: BLE001
        return ()


def sap_limpo(sap, projeto="") -> dict:
    """Cabeçalho do relatório gravado no `.obra`: os campos de sempre mais as chaves do perfil de exportação do projeto (TASK-176, ex.: `contrato`)."""
    if not isinstance(sap, dict) or not sap:
        return {}
    chaves = CAMPOS_SAP + (("titulo",)) + _chaves_do_perfil(projeto)          # `titulo` = título do SAP guardado só para os relatórios (o nome do projeto pode ser a nota)
    return {k: _texto(sap.get(k), LIMITE_TEXTO) for k in chaves if isinstance(sap.get(k), str)}


def _linha_resultado(x):
    if not isinstance(x, dict):
        return None
    saida = {}
    for k in ("operacao", "mdo", "codigo", "desc_codigo", "filtro"):
        v = x.get(k, "")
        saida[k] = v[:500] if isinstance(v, str) else ""
    t = x.get("total", 0)
    saida["total"] = t if isinstance(t, (int, float)) and not isinstance(t, bool) else 0
    return saida


def resultado_limpo(resultado):
    """`{"assinatura": str, "linhas": [...]}` validado, ou None. A assinatura é a do cálculo que gerou as linhas (ver resultado_orcamento.html)."""
    if not isinstance(resultado, dict) or not isinstance(resultado.get("linhas"), list):
        return None
    linhas = [_linha_resultado(x) for x in resultado["linhas"][:LIMITE_LINHAS_RESULTADO]]
    return {"assinatura": _texto(resultado.get("assinatura"), 64), "linhas": [x for x in linhas if x is not None]}


def montar(parcial: dict, existente: dict = None, projeto_codigo: str = "") -> str:
    """Texto JSON do `.obra`. `parcial` traz só o que mudou (nome, numero_obra, sap, dados, resultado): o resto vem de `existente`
    (o arquivo já gravado, ou None num projeto novo). Levanta ErroProjeto se faltar o nome."""
    existente = existente if isinstance(existente, dict) else {}
    pt_antigo = existente.get("projeto_trabalho") if isinstance(existente.get("projeto_trabalho"), dict) else {}
    nome = _texto(parcial.get("nome"), LIMITE_NOME) if "nome" in parcial else (_texto(pt_antigo.get("nome"), LIMITE_NOME) or _texto(existente.get("nome"), LIMITE_NOME))
    if not nome:
        raise ErroProjeto(["Informe o nome do projeto."])
    numero = _texto(parcial.get("numero_obra"), LIMITE_NUMERO) if "numero_obra" in parcial else _texto(pt_antigo.get("numero_obra"), LIMITE_NUMERO)
    codigo = _texto(projeto_codigo, 40) or _texto(existente.get("projeto"), 40)
    sap = sap_limpo(parcial.get("sap"), codigo) if "sap" in parcial else sap_limpo(pt_antigo.get("sap"), codigo)
    contexto = _texto(parcial.get("contexto"), LIMITE_CONTEXTO) if "contexto" in parcial else _texto(pt_antigo.get("contexto"), LIMITE_CONTEXTO)          # TASK-165: sem "contexto" no parcial preserva o do arquivo
    if "dados" in parcial:
        dados = obras_arquivo.normalizar_dados(parcial.get("dados"))
    else:
        dados = obras_arquivo.normalizar_dados(existente.get("dados")) if existente else {"cabos": [], "outros": [], "totalizadora": []}
    if "resultado" in parcial:
        resultado = resultado_limpo(parcial.get("resultado"))
    else:
        resultado = resultado_limpo(pt_antigo.get("resultado"))
    trabalho = {"nome": nome, "numero_obra": numero, "sap": sap, "salvo_em": datetime.now(timezone.utc).isoformat(timespec="seconds")}
    if contexto:
        trabalho["contexto"] = contexto
    if resultado is not None:
        trabalho["resultado"] = resultado
    envelope = {"formato": obras_arquivo.FORMATO, "versao": obras_arquivo.VERSAO, "exportado_em": trabalho["salvo_em"], "tipo": "obra",
                "nome": nome, "projeto": codigo, "dados": dados, "projeto_trabalho": trabalho}
    if isinstance(existente.get("cad"), dict):                 # TASK-163: a obra do plugin do nanoCAD fica como está
        envelope["cad"] = existente["cad"]
    texto = json.dumps(envelope, ensure_ascii=False, separators=(",", ":"))
    if len(texto.encode("utf-8")) > LIMITE_BYTES:
        raise ErroProjeto([f"Projeto grande demais (máx. {LIMITE_BYTES // (1024 * 1024)} MB)."])
    return texto


def ler(texto, projeto_selecionado: str) -> dict:
    """Valida um `.obra` (novo ou antigo `.obra.json`) e devolve {nome, numero_obra, sap, dados, resultado, projeto, legado}.
    Levanta ErroProjeto com as mensagens. `projeto_selecionado` = código do projeto de regras escolhido na tela (a obra precisa ser dele)."""
    if isinstance(texto, bytes):
        try:
            texto = texto.decode("utf-8-sig")
        except UnicodeDecodeError:
            raise ErroProjeto(["O arquivo não é um .obra válido (texto ilegível)."])
    try:
        carga = json.loads(texto)
    except (ValueError, TypeError):
        raise ErroProjeto(["O arquivo não é um .obra válido (JSON inválido)."])
    return ler_carga(carga, projeto_selecionado)


def ler_carga(carga, projeto_selecionado: str) -> dict:
    """Valida o envelope já decodificado e prepara os dados de trabalho."""
    try:
        pronta = obras_arquivo.validar_importacao(carga, projeto_selecionado)
    except obras_arquivo.ErroArquivoObra as e:
        raise ErroProjeto(e.mensagens)
    if pronta.get("tipo") != "obra":
        raise ErroProjeto(["Este arquivo é um modelo de obra, não um projeto."])
    trabalho = carga.get("projeto_trabalho") if isinstance(carga.get("projeto_trabalho"), dict) else None
    return {"nome": (_texto(trabalho.get("nome"), LIMITE_NOME) if trabalho else "") or pronta["nome"],
            "numero_obra": _texto(trabalho.get("numero_obra"), LIMITE_NUMERO) if trabalho else "",
            "sap": sap_limpo(trabalho.get("sap"), pronta["projeto"] or projeto_selecionado) if trabalho else {},
            "contexto": _texto(trabalho.get("contexto"), LIMITE_CONTEXTO) if trabalho else "",
            "resultado": resultado_limpo(trabalho.get("resultado")) if trabalho else None,
            "dados": pronta["dados"], "projeto": pronta["projeto"], "legado": trabalho is None,
            "cad": isinstance(carga.get("cad"), dict),          # TASK-166: veio do plugin do nanoCAD (tem geometria): o programa vincula por poste ao abrir
            "fundo": fundo_do_cad(carga.get("cad"))}            # TASK-233: rede e blocos de poste sem anotação, esmaecidos ao fundo do mapa
