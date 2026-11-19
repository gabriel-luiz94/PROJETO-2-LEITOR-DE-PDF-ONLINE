"""
services/assistente/skills.py (TASK-191) — Skills do assistente de IA: arquivos .md que ensinam o assistente (catálogo de ativos, como montar um projeto, como orçar).

Estrutura (uma pasta por escopo; `_geral` vale para todos os projetos, as demais são o nome do projeto em minúsculas, sem acento: `pernambuco`):
    data/skills/_geral/montar_projeto.md
    data/skills/pernambuco/ativos.md
Há duas raízes: a do PACOTE (acompanha o programa) e a do USUÁRIO (onde o editor grava; no .exe fica ao lado dele). A do usuário vence a do pacote
quando o nome é o mesmo; a skill do projeto vence a `_geral`. No início do arquivo, uma linha opcional `descricao: ...` resume quando usar a skill
(o assistente vê só esse resumo até pedir o conteúdo). Sem DOM nem rede; só arquivos.
"""
import os
import re
import unicodedata

import config

GERAL = "_geral"
MAX_CHARS = 30000                     # teto por skill: ela entra no contexto da IA
_RE_NOME = re.compile(r"^[a-z0-9][a-z0-9_-]{0,59}$")
_RE_DESC = re.compile(r"^\s*descri[cç][aã]o\s*:\s*(.+?)\s*$", re.IGNORECASE)


class ErroSkill(ValueError):
    """Nome/escopo inválido, conteúdo grande demais ou skill inexistente — mensagem pronta para o usuário."""


def slug(texto) -> str:
    """"PERNAMBUCO" → "pernambuco"; "Rondônia 2" → "rondonia_2"; vazio → "" (sem escopo de projeto)."""
    t = unicodedata.normalize("NFKD", str(texto or "")).encode("ascii", "ignore").decode().lower()
    return re.sub(r"[^a-z0-9]+", "_", t).strip("_")[:60]


def _raizes():
    """Pacote primeiro, usuário depois (o último vence). Lidas a cada chamada para os testes poderem trocar `config`."""
    raizes = [config.SKILLS_DIR_PACOTE]
    if os.path.abspath(config.SKILLS_DIR_USUARIO) != os.path.abspath(config.SKILLS_DIR_PACOTE):
        raizes.append(config.SKILLS_DIR_USUARIO)
    return raizes


def _validar(escopo: str, nome: str):
    if escopo != GERAL and not _RE_NOME.match(escopo or ""):
        raise ErroSkill("Escopo inválido.")
    if not _RE_NOME.match(nome or ""):
        raise ErroSkill("Nome inválido: use letras minúsculas, números, '_' e '-' (até 60).")


def _arquivos(raiz: str, escopo: str) -> dict:
    pasta = os.path.join(raiz, escopo)
    if not os.path.isdir(pasta):
        return {}
    return {f[:-3]: os.path.join(pasta, f) for f in sorted(os.listdir(pasta)) if f.endswith(".md") and _RE_NOME.match(f[:-3])}


def _ler(caminho: str) -> str:
    with open(caminho, "r", encoding="utf-8") as f:
        return f.read()


def descricao_de(texto: str) -> str:
    """`descricao: ...` no começo do arquivo; senão a primeira linha de texto que não seja título."""
    for linha in texto.splitlines()[:12]:
        m = _RE_DESC.match(linha)
        if m:
            return m.group(1)
    for linha in texto.splitlines():
        t = linha.strip()
        if t and not t.startswith("#"):
            return t[:140]
    return ""


def corpo_de(texto: str) -> str:
    """O texto sem a linha `descricao:` (é só para o índice)."""
    saida, tirou = [], False
    for linha in texto.splitlines():
        if not tirou and _RE_DESC.match(linha):
            tirou = True
            continue
        saida.append(linha)
    return "\n".join(saida).strip("\n")


def listar(projeto_slug: str = "") -> list:
    """Skills visíveis para o projeto: [{nome, escopo, descricao, origem, chars}]. A do projeto esconde a `_geral` de mesmo nome."""
    achadas = {}
    for escopo in [GERAL] + ([projeto_slug] if projeto_slug and projeto_slug != GERAL else []):
        for raiz in _raizes():
            origem = "usuario" if raiz == config.SKILLS_DIR_USUARIO and raiz != config.SKILLS_DIR_PACOTE else "pacote"
            for nome, caminho in _arquivos(raiz, escopo).items():
                texto = _ler(caminho)
                achadas[nome] = {"nome": nome, "escopo": "projeto" if escopo != GERAL else "geral", "descricao": descricao_de(texto),
                                 "origem": origem, "chars": len(texto)}
    return [achadas[n] for n in sorted(achadas)]


def carregar(projeto_slug: str, nome: str) -> str:
    """Conteúdo (sem a linha de descrição) da skill `nome`; ErroSkill se não existir. Projeto > geral; usuário > pacote."""
    nome = (nome or "").strip().lower().removesuffix(".md")
    if not _RE_NOME.match(nome):
        raise ErroSkill(f"Skill '{nome}' não existe.")
    for escopo in ([projeto_slug] if projeto_slug and projeto_slug != GERAL else []) + [GERAL]:
        for raiz in reversed(_raizes()):
            caminho = _arquivos(raiz, escopo).get(nome)
            if caminho:
                return corpo_de(_ler(caminho))
    raise ErroSkill(f"Skill '{nome}' não existe.")


def texto_bruto(escopo: str, nome: str) -> str:
    """Arquivo inteiro (com `descricao:`) para o editor; ErroSkill se não existir naquele escopo."""
    _validar(escopo, nome)
    for raiz in reversed(_raizes()):
        caminho = _arquivos(raiz, escopo).get(nome)
        if caminho:
            return _ler(caminho)
    raise ErroSkill(f"Skill '{nome}' não existe.")


def salvar(escopo: str, nome: str, texto: str) -> dict:
    """Grava na raiz do USUÁRIO (cria a pasta). Nunca altera o pacote."""
    _validar(escopo, nome)
    texto = str(texto or "").replace("\r\n", "\n")
    if not texto.strip():
        raise ErroSkill("A skill está vazia.")
    if len(texto) > MAX_CHARS:
        raise ErroSkill(f"A skill passa de {MAX_CHARS} caracteres; divida em duas.")
    pasta = os.path.join(config.SKILLS_DIR_USUARIO, escopo)
    os.makedirs(pasta, exist_ok=True)
    with open(os.path.join(pasta, nome + ".md"), "w", encoding="utf-8", newline="\n") as f:
        f.write(texto if texto.endswith("\n") else texto + "\n")
    return {"nome": nome, "escopo": escopo, "chars": len(texto)}


def apagar(escopo: str, nome: str) -> bool:
    """Apaga a skill do USUÁRIO (em dev a pasta é a mesma do pacote); False se não havia. No .exe a do pacote nunca é tocada: salve uma própria para sobrepor."""
    _validar(escopo, nome)
    caminho = os.path.join(config.SKILLS_DIR_USUARIO, escopo, nome + ".md")
    if not os.path.isfile(caminho):
        return False
    os.remove(caminho)
    return True


def indice_para_prompt(projeto_slug: str) -> str:
    """Bloco do prompt de sistema: o que existe e quando usar (o conteúdo só vem com `carregar_skill`)."""
    itens = listar(projeto_slug)
    if not itens:
        return ""
    linhas = ["SKILLS DISPONÍVEIS (chame carregar_skill com o nome antes de agir quando o pedido tocar no assunto):"]
    linhas += [f"- {s['nome']}: {s['descricao'] or '(sem descrição)'}" for s in itens]
    return "\n".join(linhas)
