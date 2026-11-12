"""
services/rascunho_ativo.py — Avalia um RASCUNHO de ativo (cadastro rápido, TASK-144) contra a base atual, sem gravar nada.

Recebe as linhas novas (como o usuário as escreveria na planilha), tira da base as linhas antigas do(s) mesmo(s) ativo(s) no projeto, junta, preenche os COMPONENTES e
roda a MESMA validação da importação da base (`validar_base`). Só devolve os problemas NOVOS (os que a base já tinha antes do rascunho não entram). Opcionalmente simula pedidos
(`ATIVO(parâmetros)`) com a base + rascunho. Função pura: sem banco.
"""
import re

from services.ativo_parametros import nome_base, preencher_componentes, validar_base
from services.fator_formula import normalizar_fator
from services.orcamento_arred import normalizar_arred
from services.orcamento_calc import processar_calculo
from services.orcamento_lv import normalizar_lv

LIMITE_LINHAS = 500
_PREFIXO_LINHA = re.compile(r"^Linha \d+:\s*")


def _escopo(linha) -> str:
    return (linha.get("projeto") or "").strip().upper()


def _normalizar_novas(novas, projeto):
    """→ (linhas normalizadas, erros). Mesma normalização de fator/LV/ARRED da importação."""
    saida, erros = [], []
    for i, l in enumerate(novas, start=1):
        if not isinstance(l, dict):
            erros.append(f"Linha nova {i}: formato inválido.")
            continue
        r = {k: ("" if l.get(k) is None else l.get(k)) for k in ("ativo", "desc_ativo", "componente", "projeto", "mdo", "origem", "codigo", "desc_codigo", "filtro")}
        for k in r:
            r[k] = str(r[k]).strip()
        r["projeto"] = r["projeto"] or (projeto or "").strip()
        try:
            r["lv"] = normalizar_lv(l.get("lv"))
            r["arred"] = normalizar_arred(l.get("arred"))
            r["fator_i"], r["formula_i"] = normalizar_fator(l.get("formula_i") or l.get("fator_i"), "FATOR I")
            r["fator_r"], r["formula_r"] = normalizar_fator(l.get("formula_r") or l.get("fator_r"), "FATOR R")
        except ValueError as e:
            erros.append(f"Linha nova {i}: {e}")
            continue
        saida.append(r)
    return saida, erros


def _sem_ativos(base, nomes, projeto):
    """A base sem as linhas dos ativos `nomes` no escopo do projeto (linha com ATIVO em branco herda o ativo de cima)."""
    proj = (projeto or "").strip().upper()
    saida, atual = [], None
    for l in base:
        cel = (l.get("ativo") or "").strip()
        if cel:
            atual = nome_base(cel)
        if atual in nomes and _escopo(l) == proj:
            continue
        saida.append(l)
    return saida


def _chave(msg: str) -> str:
    return _PREFIXO_LINHA.sub("", msg)


def avaliar(base: list, novas: list, projeto: str, simular: list = None) -> dict:
    """{"erros": [...], "avisos": [...], "linhas": [linhas novas com COMPONENTE preenchido], "componentes_gerados": n, "simulacao": [...]|None}."""
    if not isinstance(novas, list) or len(novas) > LIMITE_LINHAS:
        return {"erros": [f"Envie até {LIMITE_LINHAS} linhas."], "avisos": [], "linhas": [], "componentes_gerados": 0, "simulacao": None}
    linhas, erros = _normalizar_novas(novas, projeto)
    if erros:
        return {"erros": erros, "avisos": [], "linhas": [], "componentes_gerados": 0, "simulacao": None}
    nomes = {nome_base(l["ativo"]) for l in linhas if l["ativo"]}
    sem = _sem_ativos(base, nomes, projeto)
    antes = validar_base(preencher_componentes([dict(x) for x in sem])[0])
    juntas, n_comp = preencher_componentes([dict(x) for x in sem] + [dict(x) for x in linhas])
    depois = validar_base([dict(x) for x in juntas])
    ja_erros, ja_avisos = {_chave(x) for x in antes["erros"]}, {_chave(x) for x in antes["avisos"]}
    novos_erros = [e for e in depois["erros"] if _chave(e) not in ja_erros]
    novos_avisos = [a for a in depois["avisos"] if _chave(a) not in ja_avisos]
    resposta = {"erros": novos_erros, "avisos": novos_avisos, "linhas": juntas[len(sem):], "componentes_gerados": n_comp, "simulacao": None}
    if simular:
        saida = []
        for s in simular[:20]:
            texto = str((s or {}).get("ativo") or "").strip()
            if not texto:
                continue
            try:
                r = processar_calculo([], [{"ativo": texto, "operacao": "I", "entidade": "0"}], projeto, juntas)
                saida.append({"ativo": texto, "resultado": r.get("resultado", []), "nao_encontrados": r.get("nao_encontrados", []), "avisos": r.get("avisos", [])})
            except Exception as e:  # noqa: BLE001 — a simulação nunca derruba a validação
                saida.append({"ativo": texto, "resultado": [], "nao_encontrados": [], "avisos": [f"Não foi possível simular: {e}"]})
        resposta["simulacao"] = saida
    return resposta
