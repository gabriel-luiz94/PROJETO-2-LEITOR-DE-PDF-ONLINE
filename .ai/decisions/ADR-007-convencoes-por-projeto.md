# ADR-007 — Convenções da concessionária são configuração do projeto, não código

**Status:** PROPOSTA (15/11/2026) — tarefas TASK-173 a TASK-179

## Contexto

O programa vai ser distribuído para outras equipes criarem as próprias bases e regras. O motor (cálculo, parâmetros, vinculação,
leitor) já é genérico, e o nome "Pernambuco" só aparece em 3 arquivos. Mas várias **convenções de Pernambuco/Neoenergia** estão
embutidas no código como regex e tabelas fixas. Um projeto de outra concessionária esbarra nelas antes de conseguir usar a vinculação,
a anotação e a exportação, mesmo com uma base perfeita.

## Levantamento (main em 15/11/2026, versão 2.3.6)

| Área | Onde | O que está fixo |
|---|---|---|
| Notação do cabo | `services/anotacao_cabos.py` (+ `static/anotacao_cabos.js`) | `[Nx]CODIGO[(neutro)][-kV]`; `RAMA`/`PB` como códigos curtos; número sozinho = cobre (`16` → `C16`); neutro com vírgula = `CAZ`; `C 3x1x95` = cabo de potência |
| Nível e seção do cabo | `services/vinculacao_regras.py:170-181, 523-537` (+ JS) | MT = tensão > 1 kV no texto; seção pelo número do código e tabela AWG (`S04`, `S1/0`) |
| Ordem do "maior" | `vinculacao_regras.py:18` | `ORDEM_MAIOR_PADRAO = ["C", "P", "A"]` (em PE o protegido é `X`) |
| Postes | `services/anotacao_outros.py:27-60` (+ JS) | tipos `DT/RC/CV/FV`, sufixos `D/R/F/FV`, compacto `121000R`, altura 7–16 m |
| Desenho | `services/dxf_geometria.py:5-55` | layers `REDE_MT`/`REDE_BT`, vermelho = MT, verde = BT (o plugin já tornou isso configurável: `plugin_nanocad/src/Nucleo/ConfigLeitura.cs`) |
| Exportação | `services/exportar_sap_pdf.py`, `services/exportar_xlsx.py`, `static/orcamento_sap.js`, `services/projeto_obra.py:28` | centro 1600 (material) / 1000 (mão de obra); cabeçalho UTD/UTEP/ID investimento; layout da "nota de orçamento" SAP (verde #00A443); formato `.rec` |
| Pontos de partida | `routers/vinculacao_config.py:22-72`, `routers/regras_leitor.py:171-231`, `services/obra_params.py:16`, `config.py:56-63` | `if projeto == "081"` / `"PERNAMBUCO"` para semente da vinculação, regras do leitor e tabela regional |
| Conferência da base | — | não existe: quem edita a base não tem como saber o que mudou nos orçamentos |
| Escrita da base | `data/manual_regras_e_ajustes.md`, `static/base_ficha_*.js` | sintaxe poderosa (compostos, curingas, parâmetros fechados), só o autor domina |

## Decisão

Tudo o que é convenção de concessionária vira **configuração do projeto**, com o comportamento de Pernambuco como **valor padrão**
(projeto sem configuração = igual a hoje). As configurações de um projeto viajam juntas num **pacote de projeto** versionado, que um
projeto novo importa de uma vez. Pernambuco (081) passa a ser só o primeiro pacote, sem `if` no código.

## Consequências

- Um projeto novo funciona importando um pacote, sem mexer em código.
- Cada convenção nova é uma tabela editável com validação, espelhada em Python e JS onde a tela usa (RULES Regra 4).
- Compatibilidade: sem configuração, o resultado é idêntico ao atual (testes atuais são o critério).

## Ordem sugerida

1. TASK-173 dicionário de cabos — maior gargalo (vinculação, anotação, nível, seção, "maior").
2. TASK-174 convenção de postes.
3. TASK-175 leitura do desenho (layers e cores).
4. TASK-176 perfil de exportação.
5. TASK-177 pacote de projeto (junta 173–176 com o que já é por projeto) e fim dos `if "081"`.
6. TASK-178 casos de referência (regressão da base).
7. TASK-179 guia e modelo de base para quem começa.
