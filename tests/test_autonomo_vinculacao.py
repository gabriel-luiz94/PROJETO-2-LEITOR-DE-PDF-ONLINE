"""TASK-036 — vínculo cabo<->estrutura/poste no pipeline autônomo: porte fiel do bloco de
static/resumo.js (TASK-032/033/034), só a parte automática por coordenada."""
from services.autonomo import totalizadora as tt
from services.autonomo import vinculacao as v

REGRA_TIPO4_LIVRE = {"tipo_estrutura": "4", "qtd_cabos": 1, "compatibilidade": "LIVRE"}
REGRA_TIPO4_MESMO = {"tipo_estrutura": "4", "qtd_cabos": 2, "compatibilidade": "MESMO_TIPO_FASE_OPERACAO"}


def cabo(ativo, operacao="I", x=None, y=None, vinculo=None):
    d = {"entidade": "CABO", "operacao": operacao, "ativo": ativo}
    if x is not None:
        d["_x"], d["_y"] = x, y
    if vinculo is not None:
        d["vinculoEstruturas"] = vinculo
    return d


def estrutura(ativo, x=None, y=None):
    d = {"entidade": "ESTRUTURA", "operacao": "I", "ativo": ativo}
    if x is not None:
        d["_x"], d["_y"] = x, y
    return d


# ── tipo_estrutura (bug corrigido: dígito do CÓDIGO, não da quantidade) ──────
def test_tipo_estrutura_ignora_o_digito_da_quantidade():
    """O ativo real de uma linha Outros é sempre "<qtd>-<código>" (ex.: "1-U4", nunca "U4" sozinho —
    rejeitado pela validação de contrato C1-OUT-ORFAO). Pegar o primeiro dígito da string toda
    pegaria o da quantidade; precisa isolar o código antes."""
    assert v.tipo_estrutura("1-U4") == "4"
    assert v.tipo_estrutura("2-N3") == "3"
    assert v.tipo_estrutura("10-B4") == "4"      # quantidade com 2 dígitos
    assert v.tipo_estrutura("*3-U4") == "4"      # quantidade negativa (linha viva)
    assert v.tipo_estrutura("U4") == "4"         # sem prefixo (dado sintético de teste): cai no fallback
    assert v.tipo_estrutura("") is None and v.tipo_estrutura(None) is None


# ── tentar_vincular_automaticamente ──────────────────────────────────────────
def test_vincula_cabo_mantendo_a_estrutura_mais_proxima():
    cabos = [cabo("CAA2 ABC 35 m", "M", x=0, y=0)]
    outros = [estrutura("1-U4", x=100, y=0), estrutura("1-U4", x=1, y=0)]   # a 2ª é a mais próxima
    resumo = v.tentar_vincular_automaticamente(cabos, outros, [REGRA_TIPO4_LIVRE])
    assert resumo == {"vinculados": 1, "ja_tinham": 0, "sem_candidato": 0, "sem_coordenada": 0}
    assert cabos[0]["vinculoEstruturas"] == [1]


def test_cabo_instalando_exige_duas_estruturas():
    cabos = [cabo("CAA2 ABC 35 m", "I", x=0, y=0)]
    outros = [estrutura("U4", x=1, y=0), estrutura("U4", x=2, y=0)]
    resumo = v.tentar_vincular_automaticamente(cabos, outros, [{"tipo_estrutura": "4", "qtd_cabos": 1, "compatibilidade": "LIVRE"}])
    assert resumo["vinculados"] == 1
    assert sorted(cabos[0]["vinculoEstruturas"]) == [0, 1]


def test_nunca_sobrescreve_vinculo_existente():
    cabos = [cabo("CAA2 ABC 35 m", "M", x=0, y=0, vinculo=[0])]
    outros = [estrutura("U4", x=1, y=0)]
    resumo = v.tentar_vincular_automaticamente(cabos, outros, [REGRA_TIPO4_LIVRE])
    assert resumo == {"vinculados": 0, "ja_tinham": 1, "sem_candidato": 0, "sem_coordenada": 0}
    assert cabos[0]["vinculoEstruturas"] == [0]


def test_sem_coordenada_nao_vincula():
    cabos = [cabo("CAA2 ABC 35 m", "M")]     # sem _x/_y
    outros = [estrutura("U4", x=0, y=0)]
    resumo = v.tentar_vincular_automaticamente(cabos, outros, [REGRA_TIPO4_LIVRE])
    assert resumo == {"vinculados": 0, "ja_tinham": 0, "sem_candidato": 0, "sem_coordenada": 1}
    assert "vinculoEstruturas" not in cabos[0]


def test_sem_estrutura_disponivel_nao_vincula_nada():
    assert v.tentar_vincular_automaticamente([cabo("CAA2 ABC 35 m", "M", x=0, y=0)], [], []) == \
        {"vinculados": 0, "ja_tinham": 0, "sem_candidato": 0, "sem_coordenada": 0}


def test_respeita_capacidade_da_regra():
    """A estrutura tipo 4 só aceita 1 cabo (qtd_cabos=1); o segundo cabo fica sem candidato."""
    cabos = [cabo("CAA2 ABC 35 m", "M", x=0, y=0), cabo("CAA2 ABC 10 m", "M", x=0.1, y=0)]
    outros = [estrutura("U4", x=1, y=0)]
    resumo = v.tentar_vincular_automaticamente(cabos, outros, [REGRA_TIPO4_LIVRE])
    assert resumo["vinculados"] == 1 and resumo["sem_candidato"] == 1


def test_mesmo_tipo_fase_operacao_rejeita_cabo_diferente():
    """Estrutura tipo 4 MESMO_TIPO_FASE_OPERACAO, capacidade 2: já tem um CAA2/M; um CU16/M não entra."""
    cabos = [cabo("CAA2 ABC 35 m", "M", x=0, y=0, vinculo=[0]), cabo("CU 16 ABC 10 m", "M", x=0.1, y=0)]
    outros = [estrutura("U4", x=1, y=0)]
    resumo = v.tentar_vincular_automaticamente(cabos, outros, [REGRA_TIPO4_MESMO])
    assert resumo["sem_candidato"] == 1 and "vinculoEstruturas" not in cabos[1]


def test_mesmo_tipo_fase_operacao_aceita_cabo_igual():
    """Dois cabos MANTENDO (1 estrutura cada) compartilhando a mesma estrutura tipo 4 (capacidade 2)."""
    cabos = [cabo("CAA2 ABC 35 m", "M", x=0, y=0, vinculo=[0]), cabo("CAA2 ABC 10 m", "M", x=0.1, y=0)]
    outros = [estrutura("U4", x=1, y=0)]
    resumo = v.tentar_vincular_automaticamente(cabos, outros, [REGRA_TIPO4_MESMO])
    assert resumo["vinculados"] == 1 and cabos[1]["vinculoEstruturas"] == [0]


def test_sem_regra_cadastrada_usa_a_quantidade_exigida_pela_operacao_como_capacidade():
    cabos = [cabo("CAA2 ABC 35 m", "M", x=0, y=0)]
    outros = [estrutura("N9", x=1, y=0)]   # tipo "9": sem regra cadastrada
    resumo = v.tentar_vincular_automaticamente(cabos, outros, [])
    assert resumo["vinculados"] == 1 and cabos[0]["vinculoEstruturas"] == [0]


# ── avaliar_vinculacao ───────────────────────────────────────────────────────
def test_estrutura_sem_vinculo_gera_aviso():
    achados = v.avaliar_vinculacao([], [estrutura("U4")], [REGRA_TIPO4_LIVRE])
    assert len(achados) == 1 and achados[0]["severidade"] == "aviso" and achados[0]["regra_id"] == "VINCULO-ESTRUTURA"


def test_excecao_cabo_mantendo_perdoa_zero_vinculo():
    cabos = [cabo("CAA2 ABC 35 m", "M")]   # M em algum lugar do projeto, sem vínculo com a estrutura
    achados = v.avaliar_vinculacao(cabos, [estrutura("U4")], [REGRA_TIPO4_LIVRE])
    assert achados == []


def test_excecao_retca_perdoa_zero_vinculo():
    achados = v.avaliar_vinculacao([], [estrutura("U4"), estrutura("RETCA")], [REGRA_TIPO4_LIVRE])
    assert achados == []   # só a U4 seria avaliada (RETCA não tem regra de tipo); a exceção perdoa


def test_quantidade_errada_gera_aviso():
    cabos = [cabo("CAA2 ABC 35 m", "M", vinculo=[0]), cabo("CU 16 ABC 10 m", "M", vinculo=[0])]
    achados = v.avaliar_vinculacao(cabos, [estrutura("U4")], [REGRA_TIPO4_LIVRE])  # qtd_cabos=1, tem 2
    assert len(achados) == 1 and "esperado 1" in achados[0]["mensagem"]


def test_incompatibilidade_gera_aviso():
    cabos = [cabo("CAA2 ABC 35 m", "I", vinculo=[0]), cabo("CU 16 ABC 10 m", "I", vinculo=[0])]
    achados = v.avaliar_vinculacao(cabos, [estrutura("U4")], [REGRA_TIPO4_MESMO])
    assert len(achados) == 1 and "incompat" in achados[0]["mensagem"]


def test_sem_regra_cadastrada_nao_avalia():
    achados = v.avaliar_vinculacao([], [estrutura("N9")], [REGRA_TIPO4_LIVRE])
    assert achados == []


# ── itens_vinculo_para_totalizadora ──────────────────────────────────────────
def test_gera_um_item_por_aresta_do_vinculo():
    """Ativo real de Outros é "<qtd>-<código>" — o composto usa só o código, sem a quantidade."""
    cabos = [cabo("CAA2 ABC 35 m", "I", vinculo=[0, 1])]
    outros = [estrutura("1-N4"), estrutura("2-B4")]
    itens = v.itens_vinculo_para_totalizadora(cabos, outros)
    assert [i["ativo"] for i in itens] == ["CAA2_N4", "CAA2_B4"]
    assert all(i["origem"] == "VINCULO" and i["qtd"] == 1.0 for i in itens)


def test_cabo_sem_vinculo_nao_gera_item():
    assert v.itens_vinculo_para_totalizadora([cabo("CAA2 ABC 35 m", "I")], [estrutura("N4")]) == []


def test_ativo_sem_prefixo_identificavel_nao_gera_item():
    assert v.itens_vinculo_para_totalizadora([cabo("CAA2", "I", vinculo=[0])], [estrutura("N4")]) == []


def test_indice_de_estrutura_invalido_e_ignorado():
    assert v.itens_vinculo_para_totalizadora([cabo("CAA2 ABC 35 m", "I", vinculo=[9])], [estrutura("N4")]) == []


# ── integração com montar_totalizadora (TASK-034/036) ────────────────────────
def test_composto_sem_regra_correspondente_nunca_aparece():
    itens = v.itens_vinculo_para_totalizadora([cabo("CAA2 ABC 35 m", "I", vinculo=[0])], [estrutura("N4")])
    tot = tt.montar_totalizadora([], [], [], [], "P1", itens_vinculo=itens)
    assert tot == []


def test_composto_com_regra_subst_substitui_e_nao_mantem_o_original():
    itens = v.itens_vinculo_para_totalizadora([cabo("CAA2 ABC 35 m", "I", vinculo=[0])], [estrutura("N4")])
    regra = {"origem": "VINCULO", "op_de": "", "ativo_de": "CAA2_N4", "acao": "SUBST", "op_para": "", "ativo_para": "ALCA2",
             "fator": 3, "arredondamento": "INTEIRO", "val_min": "", "val_max": ""}
    tot = tt.montar_totalizadora([], [], [regra], [], "P1", itens_vinculo=itens)
    assert [(r["ativo"], r["qtd"], r["origem"]) for r in tot] == [("ALCA2", 3.0, "VINCULO")]


def test_composto_com_regra_adicao_mantem_original_e_derivado():
    itens = v.itens_vinculo_para_totalizadora([cabo("CAA2 ABC 35 m", "I", vinculo=[0])], [estrutura("N4")])
    regra = {"origem": "VINCULO", "op_de": "", "ativo_de": "CAA2_N4", "acao": "ADICAO", "op_para": "", "ativo_para": "ALCA2",
             "fator": 3, "arredondamento": "INTEIRO", "val_min": "", "val_max": ""}
    tot = tt.montar_totalizadora([], [], [regra], [], "P1", itens_vinculo=itens)
    assert sorted((r["ativo"], r["qtd"]) for r in tot) == [("ALCA2", 3.0), ("CAA2_N4", 1.0)]
