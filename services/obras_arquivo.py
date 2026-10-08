"""
services/obras_arquivo.py — Arquivo de obra (`.obra.json`) para exportar/importar (TASK-056). Funções puras: sem banco, sem rede.

Formato (versão 1):
    {"formato": "leitor-obra", "versao": 1, "exportado_em": "...", "tipo": "obra", "nome": "...", "projeto": "229",
     "dados": {"cabos": [...], "outros": [...], "totalizadora": [...]}}
O arquivo leva SÓ os dados da obra: nada de regras de conversão (são do projeto), marcas internas (`autonomo`), ids, usuário ou e-mail.
Tudo que vem de um arquivo é tratado como não confiável: tamanho, quantidade de linhas, tipos e campos são validados e os campos
desconhecidos são descartados. Não registrar o conteúdo das obras em log (RULES Regra 11).
"""
import json
import re
from datetime import datetime, timezone

FORMATO = "leitor-obra"
VERSAO = 1
LIMITE_BYTES = 5 * 1024 * 1024
LIMITE_LINHAS = 5000
LIMITE_ERROS = 20
OPERACOES = ("", "I", "*I", "R", "*R", "M", "*M", "0")
_RE_ENTIDADE = re.compile(r"^[A-Za-z0-9_]{1,30}$")

CAMPOS_LINHA = {"entidade": str, "operacao": str, "ativo": str, "qtdAtivos": (int, float, str), "texto": str}
CAMPOS_TOTALIZADORA = {"id": str, "baseId": str, "obs": str, "operacao": str, "ativo": str, "qtd": (int, float, str, type(None)), "desc": str,
                       "naoEncontrado": bool, "origem": str}


class ErroArquivoObra(ValueError):
    """Arquivo recusado; `args[0]` = lista de mensagens em português."""

    @property
    def mensagens(self) -> list:
        return self.args[0]


def _lista(dados, chave):
    t = (dados or {}).get(chave) if isinstance(dados, dict) else None
    if isinstance(t, dict):
        t = t.get("data")                      # forma `tableStates`: {bodyId, data: [...]}
    return [x for x in t if x is not None] if isinstance(t, list) else []


def _tipo_ok(valor, tipos) -> bool:
    tipos = tipos if isinstance(tipos, tuple) else (tipos,)
    if isinstance(valor, bool):
        return bool in tipos          # bool é subclasse de int: só vale onde o campo é de fato booleano
    return isinstance(valor, tipos)


def normalizar_dados(dados_json) -> dict:
    """Aceita o `dados_json` salvo (texto ou objeto, formas `{cabos:{data}}` e `{cabos:[...]}`) e devolve listas limpas."""
    if isinstance(dados_json, str):
        try:
            dados_json = json.loads(dados_json)
        except ValueError:
            dados_json = {}
    saida = {}
    for chave, campos in (("cabos", CAMPOS_LINHA), ("outros", CAMPOS_LINHA), ("totalizadora", CAMPOS_TOTALIZADORA)):
        linhas = []
        for x in _lista(dados_json, chave):
            if isinstance(x, dict):
                linhas.append({k: x[k] for k, tipo in campos.items() if k in x and _tipo_ok(x[k], tipo)})
        saida[chave] = linhas
    return saida


def montar_exportacao(obra: dict) -> dict:
    """`obra` = linha do banco (nome, projeto, dados_json, tipo). Devolve o envelope pronto para virar JSON.
    Modelo (TASK-058): `tipo` = "modelo", leva a lista de parâmetros e não leva Totalizadora."""
    modelo = obra.get("tipo") == "modelo"
    envelope = {"formato": FORMATO, "versao": VERSAO, "exportado_em": datetime.now(timezone.utc).isoformat(timespec="seconds"),
                "tipo": "modelo" if modelo else "obra", "nome": obra.get("nome") or "obra", "projeto": str(obra.get("projeto") or ""),
                "dados": normalizar_dados(obra.get("dados_json"))}
    if modelo:
        envelope["dados"].pop("totalizadora", None)
        envelope["parametros"] = [{"chave": p["chave"], "rotulo": p["rotulo"], "padrao": p["padrao"]}
                                  for p in _params(obra.get("dados_json"))]
    return envelope


def _params(dados_json):
    from services.modelos_obra import detectar_variaveis, montar_parametros, separar_dados
    cabos, outros, cfg = separar_dados(dados_json)
    return montar_parametros(detectar_variaveis(cabos, outros), cfg)


def nome_de_arquivo(nome: str) -> str:
    base = re.sub(r"[^\w\-. ]+", "_", str(nome or "obra"), flags=re.UNICODE).strip(" ._")[:80] or "obra"
    return f"{base}.obra.json"


def _validar_linhas(rotulo, linhas, erros, completo=True):
    if not isinstance(linhas, list):
        erros.append(f"{rotulo}: precisa ser uma lista.")
        return
    if len(linhas) > LIMITE_LINHAS:
        erros.append(f"{rotulo}: linhas demais (máx. {LIMITE_LINHAS}).")
        return
    for i, x in enumerate(linhas, 1):
        if len(erros) >= LIMITE_ERROS:
            return
        if not isinstance(x, dict):
            erros.append(f"{rotulo} linha {i}: precisa ser um objeto.")
            continue
        if completo:
            if not isinstance(x.get("ativo", ""), str) or len(x.get("ativo", "")) > 500:
                erros.append(f"{rotulo} linha {i}: 'ativo' precisa ser um texto de até 500 caracteres.")
            if x.get("operacao", "") not in OPERACOES:
                erros.append(f"{rotulo} linha {i}: operação '{str(x.get('operacao'))[:10]}' inválida (use I, *I, R, *R, M, *M).")
            ent = x.get("entidade", "0")
            if not isinstance(ent, str) or (ent != "0" and not _RE_ENTIDADE.match(ent)):
                erros.append(f"{rotulo} linha {i}: 'entidade' inválida.")


def validar_importacao(carga, projeto_selecionado: str) -> dict:
    """Valida o arquivo e devolve {nome, projeto, dados} prontos para gravar. Levanta ErroArquivoObra com todas as mensagens."""
    erros = []
    if not isinstance(carga, dict):
        raise ErroArquivoObra(["O arquivo não é uma obra exportada (esperado um objeto JSON)."])
    if carga.get("formato") != FORMATO:
        raise ErroArquivoObra(["O arquivo não é uma obra exportada por este programa (formato desconhecido)."])
    versao = carga.get("versao")
    if not isinstance(versao, int) or isinstance(versao, bool) or versao < 1:
        raise ErroArquivoObra(["Versão do arquivo inválida."])
    if versao > VERSAO:
        raise ErroArquivoObra([f"Este arquivo é de uma versão mais nova ({versao}); atualize o programa para importá-lo."])
    tipo = carga.get("tipo", "obra")
    if tipo not in ("obra", "modelo"):
        erros.append("Este arquivo é de um tipo desconhecido (esperado obra ou modelo).")
    nome = carga.get("nome")
    if not isinstance(nome, str) or not nome.strip() or len(nome) > 120:
        erros.append("O nome da obra precisa ser um texto de 1 a 120 caracteres.")
    projeto = carga.get("projeto")
    if not isinstance(projeto, str) or not projeto.strip():
        erros.append("O arquivo não informa o projeto da obra.")
    elif str(projeto_selecionado or "").strip() != projeto.strip():
        erros.append(f"Esta obra é do projeto {projeto.strip()}, mas o projeto selecionado é {str(projeto_selecionado or '(nenhum)').strip()}. "
                     f"Troque o projeto na tela e importe de novo.")
    dados = carga.get("dados")
    if not isinstance(dados, dict):
        erros.append("O arquivo não tem a seção 'dados'.")
    else:
        for chave, rotulo in (("cabos", "Cabos"), ("outros", "Outros")):
            if chave not in dados:
                erros.append(f"{rotulo}: seção ausente.")
            else:
                _validar_linhas(rotulo, dados[chave], erros)
        if "totalizadora" in dados:
            _validar_linhas("Totalizadora", dados["totalizadora"], erros, completo=False)
    parametros = []
    if tipo == "modelo" and not erros:
        from services.modelos_obra import ErroModelo, detectar_variaveis, validar_configuracao
        try:
            parametros = validar_configuracao(carga.get("parametros"))
        except ErroModelo as e:
            erros += e.mensagens
        if not detectar_variaveis(dados.get("cabos"), dados.get("outros")):
            erros.append("O modelo não tem nenhuma variável V.")
    if erros:
        raise ErroArquivoObra(erros[:LIMITE_ERROS])
    pronta = {"nome": nome.strip(), "projeto": projeto.strip(), "dados": normalizar_dados(dados), "tipo": tipo}
    if tipo == "modelo":
        pronta["dados"].pop("totalizadora", None)
        pronta["parametros"] = parametros
    return pronta


def dados_json_para_gravar(dados: dict, parametros=None) -> str:
    """Mesmo formato que o botão Salvar obra grava (`tableStates`), para o Carregar/Adicionar/Subtrair funcionarem como nas obras normais.
    Com `parametros` (modelo, TASK-058) grava a lista de parâmetros em `modelo` e não grava a Totalizadora."""
    saida = {"cabos": {"bodyId": "body-cabos", "data": dados["cabos"]}, "outros": {"bodyId": "body-outros", "data": dados["outros"]}}
    if parametros is None:
        saida["totalizadora"] = {"bodyId": "body-totalizadora", "data": dados["totalizadora"]}
    else:
        saida["modelo"] = {"parametros": parametros}
    return json.dumps(saida, ensure_ascii=False)
