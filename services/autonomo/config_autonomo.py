"""
services/autonomo/config_autonomo.py — Configuração do modo autônomo (TASK-031, fase C).

Guardada em `configuracoes` (chave `autonomo_config`, JSON), como as demais configurações locais. Pastas: uma base com `entrada/`, `processados/`,
`erros/` e `saida/` (cada uma pode ser trocada por um caminho absoluto). `user_id` = dono das obras geradas (decisão: admin escolhido).
"""
import json
import os

import config
import database

CHAVE = "autonomo_config"
NOMES_PASTAS = ("entrada", "processados", "erros", "saida")
PADRAO = {"ligado": False, "pasta_base": os.path.join(config.BASE_DIR, "autonomo"), "pasta_entrada": "", "pasta_processados": "", "pasta_erros": "",
          "pasta_saida": "", "user_id": "", "intervalo_s": 5, "estabilizacao_s": 3,
          # TASK-059 etapa 3: confirmação automática das exclusões (desligada por padrão; só nos projetos listados e com nível "confiável")
          "autoconfirmar_projetos": [], "confianca_janela": 10, "confianca_taxa": 0.9}


class ErroConfig(ValueError):
    pass


def carregar() -> dict:
    conn = database.get_connection()
    try:
        r = conn.execute("SELECT valor FROM configuracoes WHERE chave = ?", (CHAVE,)).fetchone()
    finally:
        conn.close()
    salvo = {}
    if r:
        try:
            salvo = json.loads(r[0])
        except ValueError:
            salvo = {}
    cfg = {**PADRAO, **{k: v for k, v in salvo.items() if k in PADRAO}}
    cfg["autoconfirmar_projetos"] = list(cfg["autoconfirmar_projetos"] or [])
    return cfg


def pastas(cfg: dict) -> dict:
    """{entrada, processados, erros, saida}: caminhos absolutos resolvidos (sem criar)."""
    base = cfg.get("pasta_base") or PADRAO["pasta_base"]
    return {n: os.path.abspath(cfg.get(f"pasta_{n}") or os.path.join(base, n)) for n in NOMES_PASTAS}


def _dentro(filho: str, pai: str) -> bool:
    filho, pai = os.path.normcase(os.path.abspath(filho)), os.path.normcase(os.path.abspath(pai))
    try:
        return os.path.commonpath([filho, pai]) == pai
    except ValueError:        # unidades diferentes no Windows
        return False


def validar(cfg: dict) -> dict:
    """Normaliza e valida; levanta ErroConfig com uma mensagem em português."""
    novo = {**PADRAO, **{k: v for k, v in cfg.items() if k in PADRAO}}
    novo["ligado"] = bool(novo["ligado"])
    novo["user_id"] = str(novo["user_id"] or "").strip()
    try:
        novo["intervalo_s"] = int(novo["intervalo_s"])
        novo["estabilizacao_s"] = int(novo["estabilizacao_s"])
    except (TypeError, ValueError):
        raise ErroConfig("Intervalo e tempo de estabilização precisam ser números inteiros (segundos).")
    if not 1 <= novo["intervalo_s"] <= 3600:
        raise ErroConfig("O intervalo entre varreduras deve ficar entre 1 e 3600 segundos.")
    if not 0 <= novo["estabilizacao_s"] <= 600:
        raise ErroConfig("O tempo de estabilização deve ficar entre 0 e 600 segundos.")
    try:
        novo["confianca_janela"] = int(novo["confianca_janela"])
        novo["confianca_taxa"] = float(novo["confianca_taxa"])
    except (TypeError, ValueError):
        raise ErroConfig("A janela de confiança precisa ser um número inteiro e a taxa mínima um número entre 0,5 e 1.")
    if not 3 <= novo["confianca_janela"] <= 100:
        raise ErroConfig("A janela de confiança (obras revisadas consideradas) deve ficar entre 3 e 100.")
    if not 0.5 <= novo["confianca_taxa"] <= 1.0:
        raise ErroConfig("A taxa mínima de acerto deve ficar entre 50% e 100%.")
    projetos = novo["autoconfirmar_projetos"]
    if not isinstance(projetos, list) or any(not isinstance(p, str) for p in projetos) or len(projetos) > 50:
        raise ErroConfig("Os projetos com confirmação automática precisam ser uma lista de códigos.")
    novo["autoconfirmar_projetos"] = list(dict.fromkeys(p.strip() for p in projetos if p.strip() and len(p.strip()) <= 40))
    for k in ("pasta_base", "pasta_entrada", "pasta_processados", "pasta_erros", "pasta_saida"):
        v = str(novo[k] or "").strip()
        if v and not os.path.isabs(v):
            raise ErroConfig(f"'{k}' precisa ser um caminho absoluto.")
        novo[k] = v
    p = pastas(novo)
    for n in ("processados", "erros", "saida"):
        if _dentro(p[n], p["entrada"]):
            raise ErroConfig(f"A pasta de {n} não pode ficar dentro da pasta de entrada (os resultados seriam lidos de novo).")
    if len({os.path.normcase(v) for v in p.values()}) < 4:
        raise ErroConfig("As quatro pastas (entrada, processados, erros, saída) precisam ser diferentes.")
    return novo


def salvar(cfg: dict) -> dict:
    novo = validar(cfg)
    conn = database.get_connection()
    try:
        conn.execute("INSERT OR REPLACE INTO configuracoes (chave, valor) VALUES (?, ?)", (CHAVE, json.dumps(novo, ensure_ascii=False)))
        conn.commit()
    finally:
        conn.close()
    return novo


def garantir_pastas(cfg: dict) -> dict:
    p = pastas(cfg)
    for caminho in p.values():
        os.makedirs(caminho, exist_ok=True)
    return p
