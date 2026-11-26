"""TASK-250: Montar Orçamento rápido — a preparação da base (contexto, parâmetros, compostos) fica em cache pela impressão digital do
conteúdo da base (antes ≈1 s a cada cálculo, 3 cálculos por clique); e os avisos do Montar Orçamento ficam só com erro, ativo não
encontrado (C1-BASE) e parâmetros (C1-PARAM) — a vinculação saiu dos avisos."""
import os
import time

import config
from services import anotacao_outros as ao
from services import orcamento_calc as oc


def _base(n=1500):
    rows = [{"id": 1, "ativo": "N3", "desc_ativo": "N3", "componente": "N3_", "projeto": "PE", "mdo": "MATERIAL", "codigo": "C1", "desc_codigo": "X",
             "fator_i": 1.0, "fator_r": 1.0, "filtro": "", "origem": "", "lv": "A"}]
    for i in range(n):
        rows.append(dict(rows[0], id=i + 2, ativo=f"AT{i}(cabo=A{i % 7})", componente=f"AT{i}_", codigo=f"K{i}"))
    return rows


def test_resultado_igual_com_e_sem_cache_e_base_alterada_recalcula():
    oc.limpar_cache_preparo()
    base = _base()
    pedido = [{"entidade": "0", "operacao": "I", "ativo": "2-N3"}]
    r1 = oc.processar_calculo([], pedido, "PE", [dict(r) for r in base])
    t = time.time()
    r2 = oc.processar_calculo([], pedido, "PE", [dict(r) for r in base])          # cópia nova das linhas (como vem do banco): mesmo cache
    rapido = time.time() - t
    assert r1 == r2 and r1["resultado"][0]["total"] == 2
    oc.limpar_cache_preparo()
    t = time.time()
    oc.processar_calculo([], pedido, "PE", [dict(r) for r in base])
    assert rapido < (time.time() - t)                                              # com cache é mais rápido que sem
    alterada = [dict(r) for r in base]
    alterada[0]["fator_i"] = 3.0                                                     # uma linha mudou: a chave muda e o resultado acompanha
    assert oc.processar_calculo([], pedido, "PE", alterada)["resultado"][0]["total"] == 6


def test_contexto_da_base_em_cache_pelo_conteudo():
    base = _base(50)
    a = ao.contexto_da_base([dict(r) for r in base], "PE")
    b = ao.contexto_da_base([dict(r) for r in base], "PE")
    assert a is b
    outra = [dict(r) for r in base] + [dict(base[0], id=999, ativo="N9")]
    assert ao.contexto_da_base(outra, "PE") is not a and "N9" in ao.contexto_da_base(outra, "PE")["nomes"]
    assert ao.impressao_da_base(base) == ao.impressao_da_base([dict(r) for r in base])


def test_avisos_do_montar_orcamento_sem_vinculacao():
    js = open(os.path.join(config.STATIC_DIR, "resumo.js"), encoding="utf-8").read()
    assert "avaliarVinculacaoLocal" not in js and "VinculacaoRegras.achados(" not in js
    assert "a.severidade === 'erro' || a.regra_id === 'C1-BASE' || a.regra_id === 'C1-PARAM'" in js
    i = js.index("async function checarInconsistenciasBasicas()")
    assert "res.achados = res.achados.filter(ACHADO_QUE_PARA_O_ORCAMENTO);" in js[i:i + 400]
