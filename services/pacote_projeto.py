"""
services/pacote_projeto.py — Pacote de projeto (TASK-177, ADR-007): exportar e importar tudo o que um projeto precisa num arquivo `.zip`, e pontos de partida embutidos.

Formato: `.zip` com `manifest.json` (`formato: "pacote-projeto"`, `versao_formato`, `versao_programa`, `projeto`, `descricao`, `partes` = [{nome, arquivo, hash}]) e um arquivo por parte.
Partes ausentes = "não mexer". Partes: base (`base.csv`), vinculacao, dicionario_cabos, postes, exportacao, desenho, leitor_processamento, leitor_classificacao, regional (`<parte>.json`).
Importar mostra antes (`previa`) o que muda por parte (novo / substitui / igual, contagens e avisos do validador) e só grava em `aplicar`, depois de confirmado.

Pontos de partida: cada pasta de `data/pacotes/<id>/` tem um `manifest.json` (`projeto_codigo_padrao`, `nome_projeto`, `partes` apontando para os arquivos) — um projeto sem configuração usa o
pacote do próprio cadastro (`projetos.pacote_inicial`) ou o que declara o código/nome dele. Nada no código diz "081" ou "PERNAMBUCO".
"""
import csv
import hashlib
import io
import json
import os
import zipfile
from datetime import datetime, timezone

import config
import database
from config import logger

FORMATO = "pacote-projeto"
VERSAO_FORMATO = 1
PARTES = ("base", "vinculacao", "dicionario_cabos", "postes", "exportacao", "desenho", "leitor_processamento", "leitor_classificacao", "regional", "casos")
ROTULOS = {"base": "Base técnica", "vinculacao": "Vinculação (cabos, estruturas, scripts)", "dicionario_cabos": "Dicionário de cabos", "postes": "Convenção de postes",
           "exportacao": "Perfil de exportação", "desenho": "Leitura do desenho (DXF)", "leitor_processamento": "Regras do leitor (processamento)",
           "leitor_classificacao": "Regras do leitor (classificação)", "regional": "Tabela regional (OBRA.REGIAO)", "casos": "Casos de referência"}
CHAVES_VINCULACAO = ("versao", "cabos", "estruturas", "ordem_maior", "scripts")
COLUNAS_BASE = ("ATIVO", "DESC ATIVO", "COMPONENTE", "PROJETO", "MDO", "CODIGO", "DESC CODIGO", "FATOR I", "FATOR R", "FILTRO", "ORIGEM", "LV", "ARRED")
LIMITE_BYTES = 80 * 1024 * 1024


class ErroPacote(ValueError):
    """Pedido ou arquivo recusado; `args[0]` = mensagem em português."""


# ── arquivos ──────────────────────────────────────────────────────────────────────────────────
def _hash(b: bytes) -> str:
    return hashlib.sha256(b).hexdigest()


def _json_bytes(v) -> bytes:
    return json.dumps(v, ensure_ascii=False, indent=1, sort_keys=True).encode("utf-8")


def _norm(v) -> str:
    return json.dumps(v, ensure_ascii=False, sort_keys=True)


def _trim(t) -> str:
    return str(t if t is not None else "").strip()


def versao_programa() -> str:
    return str(getattr(config, "APP_VERSION", "") or getattr(config, "VERSION", "") or "")


# ── pontos de partida (pacotes embutidos em data/pacotes/<id>/) ──────────────────────────────
def _pasta_pacotes() -> str:
    return config._get_resource_path(os.path.join("data", "pacotes"))              # funciona em dev e no .exe (data é empacotada)


def pacotes_embutidos() -> dict:
    """{id: manifesto + "pasta"} dos pacotes de partida embutidos."""
    saida = {}
    pasta = _pasta_pacotes()
    if not os.path.isdir(pasta):
        return saida
    for nome in sorted(os.listdir(pasta)):
        mf = os.path.join(pasta, nome, "manifest.json")
        if not os.path.isfile(mf):
            continue
        try:
            with open(mf, encoding="utf-8") as f:
                m = json.load(f)
        except (OSError, ValueError) as e:
            logger.warning(f"Pacote embutido '{nome}' ilegível: {e}")
            continue
        if isinstance(m, dict) and isinstance(m.get("partes"), dict):
            m = dict(m)
            m["id"] = m.get("id") or nome
            m["pasta"] = os.path.join(pasta, nome)
            saida[m["id"]] = m
    return saida


def parte_embutida(pacote_id, parte):
    """Conteúdo de uma parte de um pacote embutido (JSON já lido; `base` devolve o texto CSV); None se o pacote não tem a parte."""
    m = pacotes_embutidos().get(pacote_id)
    rel = (m or {}).get("partes", {}).get(parte)
    if not rel:
        return None
    caminho = os.path.normpath(os.path.join(m["pasta"], rel))
    try:
        if parte == "base":
            with open(caminho, encoding="utf-8-sig") as f:
                return f.read()
        with open(caminho, encoding="utf-8") as f:
            return json.load(f)
    except (OSError, ValueError) as e:
        logger.warning(f"Parte '{parte}' do pacote '{pacote_id}' ilegível: {e}")
        return None


def projeto_do_cadastro(codigo_ou_nome):
    """{"nome", "codigo", "pacote_inicial"} do cadastro de projetos (pelo código ou pelo nome); None se não existe."""
    chave = _trim(codigo_ou_nome)
    if not chave:
        return None
    conn = database.get_row_connection()
    try:
        r = conn.execute("SELECT nome, codigo, COALESCE(pacote_inicial, '') AS pacote_inicial FROM projetos WHERE codigo = ? OR UPPER(nome) = ? LIMIT 1", (chave, chave.upper())).fetchone()
    except Exception:  # noqa: BLE001
        return None
    finally:
        conn.close()
    return dict(r) if r else None


def pacote_inicial_de(codigo_ou_nome):
    """O pacote de partida do projeto: o marcado no cadastro (`pacote_inicial`) ou o que declara o código/nome dele; None se não há."""
    chave = _trim(codigo_ou_nome)
    if not chave:
        return None
    pacotes = pacotes_embutidos()
    proj = projeto_do_cadastro(chave) or {}
    if proj.get("pacote_inicial") in pacotes:
        return pacotes[proj["pacote_inicial"]]
    codigo, nome = _trim(proj.get("codigo") or chave), _trim(proj.get("nome") or chave).upper()
    for m in pacotes.values():
        if _trim(m.get("projeto_codigo_padrao")) == codigo or _trim(m.get("nome_projeto")).upper() == nome:
            return m
    return None


def parte_inicial(codigo_ou_nome, parte):
    """A parte do pacote de partida do projeto (None se o projeto não tem pacote ou o pacote não traz a parte)."""
    m = pacote_inicial_de(codigo_ou_nome)
    return parte_embutida(m["id"], parte) if m else None


def config_inicial(codigo_ou_nome):
    """Configuração da vinculação do pacote de partida (vinculação + dicionário, postes e exportação) ou None."""
    m = pacote_inicial_de(codigo_ou_nome)
    if not m:
        return None
    cfg = parte_embutida(m["id"], "vinculacao")
    if not isinstance(cfg, dict):
        return None
    cfg = dict(cfg)
    for parte in ("dicionario_cabos", "postes", "exportacao"):
        v = parte_embutida(m["id"], parte)
        if v not in (None, [], {}):
            cfg[parte] = v
    return cfg


# ── partes guardadas (regional e afins) ───────────────────────────────────────────────────────
def _garantir_tabela(conn):
    conn.execute("CREATE TABLE IF NOT EXISTS projeto_partes (projeto_codigo TEXT NOT NULL, parte TEXT NOT NULL, conteudo_json TEXT NOT NULL, updated_at TEXT DEFAULT (datetime('now')), PRIMARY KEY (projeto_codigo, parte))")


def parte_guardada(codigo, parte):
    conn = database.get_connection()
    try:
        _garantir_tabela(conn)
        r = conn.execute("SELECT conteudo_json FROM projeto_partes WHERE projeto_codigo = ? AND parte = ?", (codigo, parte)).fetchone()
    finally:
        conn.close()
    return json.loads(r[0]) if r else None


def guardar_parte(codigo, parte, valor):
    conn = database.get_connection()
    try:
        _garantir_tabela(conn)
        conn.execute("INSERT OR REPLACE INTO projeto_partes (projeto_codigo, parte, conteudo_json) VALUES (?, ?, ?)", (codigo, parte, json.dumps(valor, ensure_ascii=False)))
        conn.commit()
    finally:
        conn.close()


# ── estado atual de um projeto, por parte ─────────────────────────────────────────────────────
def _base_do_projeto(nome):
    conn = database.get_row_connection()
    try:
        return [dict(r) for r in conn.execute("SELECT * FROM tabela_orcamento_master WHERE UPPER(TRIM(projeto)) = ? ORDER BY id", (_trim(nome).upper(),)).fetchall()]
    finally:
        conn.close()


def _csv_base(linhas) -> str:
    buf = io.StringIO()
    w = csv.writer(buf, delimiter=";", lineterminator="\n")
    w.writerow(COLUNAS_BASE)
    for r in linhas:
        fi = _trim(r.get("formula_i")) or r.get("fator_i")
        fr = _trim(r.get("formula_r")) or r.get("fator_r")
        w.writerow([r.get("ativo") or "", r.get("desc_ativo") or "", r.get("componente") or "", r.get("projeto") or "", r.get("mdo") or "", r.get("codigo") or "", r.get("desc_codigo") or "",
                    "" if fi is None else str(fi).replace(".", ",") if isinstance(fi, float) else fi, "" if fr is None else str(fr).replace(".", ",") if isinstance(fr, float) else fr,
                    r.get("filtro") or "", r.get("origem") or "", r.get("lv") or "A", r.get("arred") or "NORMAL"])
    return buf.getvalue()


def _linhas_do_csv(texto):
    from services.fator_formula import normalizar_fator
    from services.orcamento_arred import normalizar_arred
    from services.orcamento_lv import normalizar_lv
    leitor = csv.DictReader(io.StringIO(texto), delimiter=";")
    if not leitor.fieldnames or len(leitor.fieldnames) < 2:
        leitor = csv.DictReader(io.StringIO(texto), delimiter=",")
    saida = []
    for n, row in enumerate(leitor, start=2):
        r = {k.strip().upper(): (v or "") for k, v in row.items() if k}
        ativo, codigo = r.get("ATIVO", "").strip(), r.get("CODIGO", "").strip()
        if not ativo and not codigo:
            continue
        try:
            fi, fmi = normalizar_fator(r.get("FATOR I"), "FATOR I")
            fr, fmr = normalizar_fator(r.get("FATOR R"), "FATOR R")
            lv, arred = normalizar_lv(r.get("LV")), normalizar_arred(r.get("ARRED"))
        except ValueError as e:
            raise ErroPacote(f"base.csv, linha {n}: {e}")
        saida.append({"ativo": ativo, "desc_ativo": r.get("DESC ATIVO", "").strip(), "componente": r.get("COMPONENTE", "").strip(), "projeto": r.get("PROJETO", "").strip(), "mdo": r.get("MDO", "").strip(),
                      "codigo": codigo, "desc_codigo": r.get("DESC CODIGO", "").strip(), "fator_i": fi, "fator_r": fr, "filtro": r.get("FILTRO", "").strip(), "origem": r.get("ORIGEM", "").strip(),
                      "lv": lv, "arred": arred, "formula_i": fmi, "formula_r": fmr})
    return saida


def _config_efetiva(codigo):
    from routers.vinculacao_config import obter_config
    return obter_config(codigo)


def valor_atual(projeto, parte):
    """O valor da parte hoje no projeto (`projeto` = {nome, codigo}); None = o projeto não tem a parte."""
    codigo, nome = projeto["codigo"], projeto["nome"]
    if parte == "base":
        linhas = _base_do_projeto(nome)
        return _csv_base(linhas) if linhas else None
    if parte == "vinculacao":
        cfg = _config_efetiva(codigo)
        return {k: cfg[k] for k in CHAVES_VINCULACAO if k in cfg}
    if parte in ("dicionario_cabos", "postes", "exportacao"):
        v = _config_efetiva(codigo).get(parte)
        return v if v else None
    if parte == "desenho":
        from routers.plugin_cad import leitura_salva
        return leitura_salva(codigo)
    if parte in ("leitor_processamento", "leitor_classificacao"):
        from routers.regras_leitor import _get_regras
        r = _get_regras(parte.split("_", 1)[1], codigo)
        return r or None
    if parte == "regional":
        return parte_guardada(codigo, "regional") or parte_inicial(codigo, "regional")
    if parte == "casos":
        from services import casos_referencia as cr
        return [{k: v for k, v in c.items() if k != "id"} for c in cr.listar(codigo)] or None
    raise ErroPacote(f"Parte desconhecida: {parte}.")


# ── exportar ──────────────────────────────────────────────────────────────────────────────────
def exportar(codigo_ou_nome, partes=None, descricao="") -> bytes:
    """`.zip` do pacote com as `partes` pedidas (padrão: todas as que o projeto tem)."""
    proj = projeto_do_cadastro(codigo_ou_nome)
    if not proj:
        raise ErroPacote(f"Projeto '{codigo_ou_nome}' não está cadastrado.")
    pedidas = list(partes) if partes else list(PARTES)
    ruins = [p for p in pedidas if p not in PARTES]
    if ruins:
        raise ErroPacote(f"Parte desconhecida: {', '.join(ruins)}.")
    arquivos, lista = {}, []
    for p in PARTES:
        if p not in pedidas:
            continue
        v = valor_atual(proj, p)
        if v is None:
            continue
        nome = "base.csv" if p == "base" else p + ".json"
        conteudo = v.encode("utf-8") if p == "base" else _json_bytes(v)
        arquivos[nome] = conteudo
        lista.append({"nome": p, "arquivo": nome, "hash": _hash(conteudo)})
    if not lista:
        raise ErroPacote("O projeto não tem nenhuma das partes pedidas para exportar.")
    manifesto = {"formato": FORMATO, "versao_formato": VERSAO_FORMATO, "versao_programa": versao_programa(), "projeto": {"nome": proj["nome"], "codigo": proj["codigo"]},
                 "descricao": _trim(descricao)[:300], "exportado_em": datetime.now(timezone.utc).isoformat(timespec="seconds"), "partes": lista}
    buf = io.BytesIO()
    with zipfile.ZipFile(buf, "w", zipfile.ZIP_DEFLATED) as z:
        z.writestr("manifest.json", _json_bytes(manifesto))
        for nome, c in arquivos.items():
            z.writestr(nome, c)
    return buf.getvalue()


# ── ler ───────────────────────────────────────────────────────────────────────────────────────
def ler_pacote(conteudo: bytes) -> dict:
    """Valida o `.zip` (formato, versão, hashes) e devolve {"manifesto", "partes": {nome: valor}} (base = texto CSV). Levanta ErroPacote."""
    if len(conteudo) > LIMITE_BYTES:
        raise ErroPacote(f"Pacote grande demais (máx. {LIMITE_BYTES // (1024 * 1024)} MB).")
    try:
        z = zipfile.ZipFile(io.BytesIO(conteudo))
    except zipfile.BadZipFile:
        raise ErroPacote("O arquivo não é um pacote de projeto (.zip inválido).")
    with z:
        if "manifest.json" not in z.namelist():
            raise ErroPacote("O arquivo não é um pacote de projeto (falta o manifest.json).")
        try:
            m = json.loads(z.read("manifest.json").decode("utf-8-sig"))
        except ValueError:
            raise ErroPacote("O manifest.json do pacote está ilegível.")
        if not isinstance(m, dict) or m.get("formato") != FORMATO:
            raise ErroPacote("O arquivo não é um pacote de projeto (formato diferente).")
        v = m.get("versao_formato")
        if not isinstance(v, int) or isinstance(v, bool):
            raise ErroPacote("O pacote não informa a versão do formato.")
        if v > VERSAO_FORMATO:
            raise ErroPacote(f"Este pacote usa o formato versão {v}, mais novo que o suportado por este programa (versão {VERSAO_FORMATO}). Atualize o OrçaCore para importá-lo.")
        partes = {}
        for item in m.get("partes") or []:
            nome, arq = _trim(item.get("nome")), _trim(item.get("arquivo"))
            if nome not in PARTES:
                raise ErroPacote(f"O pacote traz uma parte desconhecida: '{nome}'.")
            if arq not in z.namelist() or "/" in arq or "\\" in arq or arq.startswith("."):
                raise ErroPacote(f"Falta o arquivo da parte '{nome}' no pacote.")
            b = z.read(arq)
            if _trim(item.get("hash")) and _hash(b) != item["hash"]:
                raise ErroPacote(f"A parte '{nome}' foi alterada depois de exportada (a conferência de integridade falhou).")
            try:
                partes[nome] = b.decode("utf-8-sig") if nome == "base" else json.loads(b.decode("utf-8-sig"))
            except ValueError:
                raise ErroPacote(f"A parte '{nome}' está ilegível.")
    return {"manifesto": m, "partes": partes}


# ── prévia ────────────────────────────────────────────────────────────────────────────────────
def _contagem(parte, v) -> str:
    if v is None:
        return "—"
    if parte == "base":
        return f"{len(_linhas_do_csv(v))} linhas"
    if parte == "casos":
        return f"{len(v)} casos"
    if parte == "vinculacao":
        return f"{len(v.get('cabos') or [])} cabos · {len(v.get('estruturas') or [])} estruturas · {len(v.get('scripts') or [])} scripts"
    if isinstance(v, list):
        return f"{len(v)} itens"
    return f"{len(v)} campos" if isinstance(v, dict) else "1 item"


def _avisos_da_parte(parte, v, destino):
    """Erros/avisos do validador da própria parte (sem gravar nada)."""
    avisos = []
    try:
        if parte in ("vinculacao", "dicionario_cabos", "postes", "exportacao"):
            from services import vinculacao_regras as vr
            cfg = dict(v) if parte == "vinculacao" else {parte: v}
            avisos += vr.validar_config(cfg)
        elif parte == "base":
            from services.ativo_parametros import validar_base
            rows = _linhas_do_csv(v)
            r = validar_base(rows)
            avisos += [f"base: {e}" for e in r.get("erros", [])[:20]] + [f"base (aviso): {a}" for a in r.get("avisos", [])[:20]]
        elif parte == "casos":
            from services import casos_referencia as cr
            for c in v:
                cr.normalizar_caso(c)
        elif parte == "desenho":
            from services import leitura_desenho as ld
            avisos += ld.validar(v)
        elif parte in ("leitor_processamento", "leitor_classificacao"):
            from routers.regras_leitor import validar_regras
            avisos += [str(e) for e in validar_regras(parte.split("_", 1)[1], v)]
    except ErroPacote as e:
        avisos.append(str(e))
    except Exception as e:  # noqa: BLE001
        avisos.append(f"não consegui validar: {e}")
    return avisos


def previa(codigo_ou_nome, pacote: dict) -> dict:
    """O que a importação faria no projeto, por parte: situacao = novo | substitui | igual, contagens (antes → depois) e avisos. Não grava nada."""
    proj = projeto_do_cadastro(codigo_ou_nome)
    if not proj:
        raise ErroPacote(f"Projeto '{codigo_ou_nome}' não está cadastrado.")
    saida = []
    for parte, novo in pacote["partes"].items():
        atual = valor_atual(proj, parte)
        if parte == "base":
            origem = _trim((pacote["manifesto"].get("projeto") or {}).get("nome"))
            igual = atual is not None and _norm([{k: v for k, v in r.items() if k != "projeto"} for r in _linhas_do_csv(atual)]) == _norm([{k: v for k, v in r.items() if k != "projeto"} for r in _linhas_do_csv(novo)])
        else:
            igual = atual is not None and _norm(atual) == _norm(novo)
        situacao = "novo" if atual is None else ("igual" if igual else "substitui")
        saida.append({"parte": parte, "rotulo": ROTULOS[parte], "situacao": situacao, "antes": _contagem(parte, atual), "depois": _contagem(parte, novo), "avisos": _avisos_da_parte(parte, novo, proj)})
    res = {"projeto": proj, "origem": pacote["manifesto"].get("projeto"), "descricao": pacote["manifesto"].get("descricao", ""), "versao_programa": pacote["manifesto"].get("versao_programa", ""), "partes": saida}
    if "base" in pacote["partes"]:                                                                  # TASK-178: os casos de referência (do pacote ou do projeto) rodam com a base do pacote, sem gravar
        try:
            from routers.casos import conferir_projeto
            from services import casos_referencia as cr
            casos = [dict(cr.normalizar_caso(c), id=i + 1) for i, c in enumerate(pacote["partes"].get("casos") or [])] if pacote["partes"].get("casos") else None
            r = conferir_projeto(proj, None, None, pacote["partes"]["base"], casos)
            res["casos"] = {"resumo": r["resumo"], "total": r["total"], "segundos": r["segundos"], "mudaram": [c for c in r["casos"] if c["situacao"] != "igual"][:50]}
        except Exception as e:  # noqa: BLE001
            res["casos"] = {"erro": str(e)[:200]}
    return res


# ── aplicar ───────────────────────────────────────────────────────────────────────────────────
def _aplicar_base(proj, texto, user_id):
    from services.ativo_parametros import preencher_componentes, resumo_erros, validar_base
    from services.supabase_client import get_supabase
    from services.sync_service import espelhar_master_no_pessoal, normalizar_projeto, validar_linhas_do_projeto
    linhas = _linhas_do_csv(texto)
    alvo = normalizar_projeto(proj["nome"])
    linhas = [dict(l, projeto="") if normalizar_projeto(l["projeto"]) == normalizar_projeto(_ORIGEM.get("nome")) else l for l in linhas]       # a base exportada do projeto de origem passa a ser do destino
    try:
        linhas = validar_linhas_do_projeto(linhas, proj["nome"])
    except ValueError as e:
        raise ErroPacote(str(e))
    linhas, _ = preencher_componentes(linhas)
    val = validar_base(linhas)
    if val["erros"]:
        raise ErroPacote("A base do pacote tem erros: " + resumo_erros(val["erros"]))
    conn = database.get_connection()
    try:
        cur = conn.cursor()
        cur.execute("DELETE FROM tabela_orcamento_master WHERE UPPER(TRIM(projeto)) = ?", (alvo,))
        for r in linhas:
            cur.execute("INSERT INTO tabela_orcamento_master (ativo, desc_ativo, componente, projeto, mdo, codigo, desc_codigo, fator_i, fator_r, filtro, origem, lv, arred, formula_i, formula_r) "
                        "VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?)", (r["ativo"], r["desc_ativo"], r["componente"], r["projeto"], r["mdo"], r["codigo"], r["desc_codigo"], r["fator_i"], r["fator_r"],
                                                                                    r["filtro"], r["origem"], r["lv"], r["arred"], r["formula_i"], r["formula_r"]))
        conn.commit()
    finally:
        conn.close()
    espelhar_master_no_pessoal(user_id, proj["nome"])
    supabase = get_supabase()
    if supabase and linhas:
        try:
            supabase.table("tabela_orcamento_master").delete().ilike("projeto", alvo).execute()
            for i in range(0, len(linhas), 500):
                supabase.table("tabela_orcamento_master").insert(linhas[i:i + 500]).execute()
        except Exception as e:  # noqa: BLE001
            raise ErroPacote(f"Base gravada neste computador, mas falhou ao sincronizar com a nuvem (Supabase): {e}")


_ORIGEM = {}


def _aplicar_parte(proj, parte, valor, user_id, email):
    codigo = proj["codigo"]
    if parte == "base":
        _aplicar_base(proj, valor, user_id)
    elif parte in ("vinculacao", "dicionario_cabos", "postes", "exportacao"):
        from routers.vinculacao_config import salvar_config
        cfg = dict(_config_efetiva(codigo))
        if parte == "vinculacao":
            for k in CHAVES_VINCULACAO:
                cfg.pop(k, None)
            cfg.update({k: valor[k] for k in CHAVES_VINCULACAO if k in valor})
        else:
            cfg[parte] = valor
        salvar_config(codigo, cfg)
    elif parte == "desenho":
        from routers.plugin_cad import LeituraPayload, salvar_leitura
        salvar_leitura(codigo, LeituraPayload(leitura=valor))
    elif parte in ("leitor_processamento", "leitor_classificacao"):
        from routers.regras_leitor import RegrasLeitorPayload, _save_regras
        _save_regras(parte.split("_", 1)[1], RegrasLeitorPayload(projeto_codigo=codigo, regras=valor), email)
    elif parte == "casos":
        from services import casos_referencia as cr
        for c in cr.listar(codigo):
            cr.excluir(codigo, c["id"])
        for c in valor:
            cr.gravar(codigo, c, email)
    elif parte == "regional":
        guardar_parte(codigo, "regional", valor)
        from services import obra_params
        obra_params.limpar_cache()


def aplicar(codigo_ou_nome, pacote: dict, partes=None, user_id=None, email=None) -> dict:
    """Grava no projeto as `partes` escolhidas (padrão: todas do pacote). Antes valida TUDO (nada é gravado se algo não passa); depois grava parte por parte."""
    proj = projeto_do_cadastro(codigo_ou_nome)
    if not proj:
        raise ErroPacote(f"Projeto '{codigo_ou_nome}' não está cadastrado.")
    escolhidas = [p for p in PARTES if p in pacote["partes"] and (not partes or p in partes)]
    if not escolhidas:
        raise ErroPacote("Nenhuma parte escolhida para importar.")
    _ORIGEM.clear()
    _ORIGEM.update(pacote["manifesto"].get("projeto") or {})
    for p in escolhidas:                                                                        # validação completa antes de gravar qualquer coisa
        erros = [a for a in _avisos_da_parte(p, pacote["partes"][p], proj) if not a.startswith("base (aviso)")]
        if erros:
            raise ErroPacote(f"{ROTULOS[p]}: " + "; ".join(erros[:5]) + " Nada foi gravado.")
    feitas = []
    for p in escolhidas:
        _aplicar_parte(proj, p, pacote["partes"][p], user_id, email)
        feitas.append(p)
    return {"projeto": proj, "gravadas": feitas}
