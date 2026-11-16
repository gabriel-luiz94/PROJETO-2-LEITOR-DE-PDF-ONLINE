"""
services/perfil_exportacao.py — Perfil de exportação por projeto (TASK-176, ADR-007).

Os formatos de saída (PDF/XLSX ODI–ODD, visão SAP e `.rec`) seguem um perfil do projeto, guardado na configuração da vinculação (`exportacao`). Sem perfil vale o de
Neoenergia/Pernambuco (PADRAO), idêntico ao que era fixo no código (saídas iguais, comparadas em tests/test_perfil_exportacao_176.py).

  {"cabecalho": [{"chave": "bairro", "rotulo": "BAIRRO", "tela": "Bairro", "coluna": "esq"|"meio", "obrigatorio": false, "obra": true, "lembrar": false}, …],
   "centro": {"MATERIAL": "1600", "MAO-DE-OBRA": "1000"},
   "pdf": {"cor_titulo": "#00A443", "rodape": "Internal Use", "rotulo_ordem": "TIPO DE ORDEM", "rotulo_id": "ID INVESTIMENTO"},
   "xlsx": {"colunas": ["mdo", "codigo", "qtd", "centro", "desc"]},
   "arquivos": ["pdf_sap", "xlsx", "rec", "csv"]}
`cabecalho`: campos do cabeçalho do relatório. `coluna` esq = 1ª coluna do cabeçalho do PDF (até 5 linhas), meio = 2ª (até 4 linhas, a partir da 2ª linha); `id` = o ID do
investimento (ODI/ODD) do canto. `titulo` e `nota` vêm do nome e do número do projeto. `obra: true` = o campo vira o parâmetro OBRA.<CHAVE> da base (TASK-165).
O `.obra` grava as chaves do perfil (menos título e nota). Mesma lógica de static/perfil_exportacao.js (paridade testada).
"""
import json
import re

COLUNAS_XLSX = {"mdo": "TIPO MDO", "codigo": "CÓDIGO", "qtd": "QUANTIDADE", "centro": "CENTRO", "desc": "DESCRIÇÃO"}
ARQUIVOS = ("pdf_sap", "xlsx", "rec", "csv")
_RE_CHAVE = re.compile(r"^[a-z][a-z0-9_]{0,30}$")
_RE_COR = re.compile(r"^#[0-9A-Fa-f]{6}$")


def _c(chave, rotulo, tela, coluna, obra=False, lembrar=False, obrigatorio=False):
    return {"chave": chave, "rotulo": rotulo, "tela": tela, "coluna": coluna, "obrigatorio": obrigatorio, "obra": obra, "lembrar": lembrar}


PADRAO = {
    "cabecalho": [_c("titulo", "TÍTULO:", "Título", "esq"), _c("nota", "Nº NOTA:", "Nº da nota", "esq"), _c("sap", "Nº SAP:", "Nº SAP", "meio"),
                  _c("bairro", "BAIRRO", "Bairro", "esq", obra=True), _c("municipio", "MUNICIPIO:", "Município", "meio", obra=True, lembrar=True),
                  _c("utd", "UTD:", "UTD", "esq", obra=True, lembrar=True), _c("utep", "UTEP:", "UTEP", "meio", obra=True, lembrar=True),
                  _c("projetista", "PROJETISTA:", "Projetista", "esq", lembrar=True), _c("data", "DATA:", "Data", "meio"),
                  _c("id_investimento", "ID INVESTIMENTO", "ID Investimento (ODI)", "id", lembrar=True),
                  _c("id_investimento_odd", "ID INVESTIMENTO (ODD)", "ID Investimento (ODD) — vazio = igual ao ODI", "id", lembrar=True)],
    "centro": {"MATERIAL": "1600", "MAO-DE-OBRA": "1000"},
    "pdf": {"cor_titulo": "#00A443", "rodape": "Internal Use", "rotulo_ordem": "TIPO DE ORDEM", "rotulo_id": "ID INVESTIMENTO"},
    "xlsx": {"colunas": ["mdo", "codigo", "qtd", "centro", "desc"]},
    "arquivos": list(ARQUIVOS),
}


def _trim(t) -> str:
    return str(t if t is not None else "").strip()


def _copia(o):
    return json.loads(json.dumps(o))


def normalizar(perfil) -> dict:
    """Perfil completo: campos que faltam usam o padrão de Neoenergia; sem `cabecalho` válido vale o padrão."""
    if not isinstance(perfil, dict):
        return _copia(PADRAO)
    p = _copia(PADRAO)
    campos = []
    for x in perfil.get("cabecalho") if isinstance(perfil.get("cabecalho"), list) else []:
        if isinstance(x, dict) and _RE_CHAVE.match(_trim(x.get("chave"))):
            ch = _trim(x["chave"])
            col = _trim(x.get("coluna")).lower()
            campos.append({"chave": ch, "rotulo": _trim(x.get("rotulo")) or ch.upper() + ":", "tela": _trim(x.get("tela")) or _trim(x.get("rotulo")).rstrip(":") or ch,
                           "coluna": col if col in ("esq", "meio", "id") else "meio", "obrigatorio": x.get("obrigatorio") is True, "obra": x.get("obra") is True, "lembrar": x.get("lembrar") is True})
    if campos:
        p["cabecalho"] = campos
    if isinstance(perfil.get("centro"), dict) and perfil["centro"]:
        p["centro"] = {_trim(k).upper(): _trim(v) for k, v in perfil["centro"].items() if _trim(k)}
    if isinstance(perfil.get("pdf"), dict):
        for k in p["pdf"]:
            if _trim(perfil["pdf"].get(k)):
                p["pdf"][k] = _trim(perfil["pdf"][k])
    cols = perfil.get("xlsx", {}).get("colunas") if isinstance(perfil.get("xlsx"), dict) else None
    if isinstance(cols, list) and cols and all(c in COLUNAS_XLSX for c in cols):
        p["xlsx"]["colunas"] = list(dict.fromkeys(cols))
    if isinstance(perfil.get("arquivos"), list):
        p["arquivos"] = [a for a in perfil["arquivos"] if a in ARQUIVOS]
    return p


def validar(perfil) -> list:
    erros = []
    if perfil is None:
        return erros
    if not isinstance(perfil, dict):
        return ["'exportacao' precisa ser um objeto."]
    cab = perfil.get("cabecalho")
    if cab is not None:
        if not isinstance(cab, list):
            erros.append("Exportação: 'cabecalho' precisa ser uma lista de campos.")
        else:
            vistas = set()
            for n, x in enumerate(cab, 1):
                if not isinstance(x, dict) or not _RE_CHAVE.match(_trim(x.get("chave"))):
                    erros.append(f"Exportação, campo {n} do cabeçalho: 'chave' precisa ser minúscula, sem espaço (ex.: contrato).")
                    continue
                if x["chave"] in vistas:
                    erros.append(f"Exportação: a chave '{x['chave']}' do cabeçalho está repetida.")
                vistas.add(x["chave"])
                if _trim(x.get("coluna")) and _trim(x["coluna"]).lower() not in ("esq", "meio", "id"):
                    erros.append(f"Exportação, campo '{x['chave']}': 'coluna' precisa ser esq, meio ou id.")
            if cab and len([x for x in cab if isinstance(x, dict) and _trim(x.get("coluna")).lower() == "esq"]) > 5:
                erros.append("Exportação: a 1ª coluna do cabeçalho do PDF comporta até 5 campos.")
            if cab and len([x for x in cab if isinstance(x, dict) and _trim(x.get("coluna")).lower() in ("", "meio")]) > 4:
                erros.append("Exportação: a 2ª coluna do cabeçalho do PDF comporta até 4 campos.")
    c = perfil.get("centro")
    if c is not None and (not isinstance(c, dict) or not all(isinstance(v, (str, int)) for v in c.values())):
        erros.append("Exportação: 'centro' precisa ser {\"MATERIAL\": \"1600\", \"MAO-DE-OBRA\": \"1000\"}.")
    pdf = perfil.get("pdf")
    if pdf is not None:
        if not isinstance(pdf, dict):
            erros.append("Exportação: 'pdf' precisa ser um objeto.")
        elif pdf.get("cor_titulo") not in (None, "") and not _RE_COR.match(_trim(pdf["cor_titulo"])):
            erros.append("Exportação: 'pdf.cor_titulo' precisa ser #RRGGBB.")
    x = perfil.get("xlsx")
    if x is not None and not (isinstance(x, dict) and (x.get("colunas") is None or (isinstance(x["colunas"], list) and x["colunas"] and all(k in COLUNAS_XLSX for k in x["colunas"])))):
        erros.append(f"Exportação: 'xlsx.colunas' precisa ser uma lista com {list(COLUNAS_XLSX)}.")
    a = perfil.get("arquivos")
    if a is not None and (not isinstance(a, list) or any(v not in ARQUIVOS for v in a)):
        erros.append(f"Exportação: 'arquivos' precisa ser uma lista com {list(ARQUIVOS)}.")
    return erros


def campos_sap(perfil=None) -> list:
    """Chaves do cabeçalho que o `.obra` grava (todas, menos título e nota — esses vêm do nome e do número do projeto)."""
    return [x["chave"] for x in normalizar(perfil)["cabecalho"] if x["chave"] not in ("titulo", "nota")]


def chaves_obra(perfil=None) -> dict:
    """{chave: "OBRA.CHAVE"} dos campos marcados `obra` (padrão: municipio, bairro, utd, utep)."""
    return {x["chave"]: "OBRA." + x["chave"].upper() for x in normalizar(perfil)["cabecalho"] if x["obra"]}


def centro_de(mdo, perfil=None) -> str:
    """CENTRO do tipo de mão de obra: `mdo` 'M' (material) ou 'N' (mão de obra); outro = vazio."""
    c = normalizar(perfil)["centro"]
    return c.get({"M": "MATERIAL", "N": "MAO-DE-OBRA"}.get(_trim(mdo).upper(), ""), "")


def oferece(arquivo, perfil=None) -> bool:
    return arquivo in normalizar(perfil)["arquivos"]
