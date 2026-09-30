# CATÁLOGO DE VALIDAÇÕES

> Lista de tudo que a validação das planilhas **Cabos** e **Outros** verifica. É a fonte de contexto
> das TASKs 011 (camada 1), 013 (camada 2), 012/014/015 (camada 3). Arquitetura: `ADR-004`.
>
> Regra 2 (RULES): **nenhuma regra de domínio é implementada ou ativada sem status `CONFIRMADA`.**
> Este arquivo é a lista para revisão linha a linha com o usuário.

## Legenda

- **Camada 1** — contrato, em código, **não editável**. **Camada 2** — regra de domínio, dados, **editável pelo admin**. **Camada 3** — IA, prompt salvo, editável pelo admin.
- **Severidade:** `erro` · `aviso` · `info`.
- **Status:** `DERIVADA DO CÓDIGO` (comportamento verificado em `orcamento_calc.py`) · `IMPLEMENTADA` (regra decidida e executada pelo motor; desligada na semente até o admin ligar) · `A CONFIRMAR` (vem de `prompt_rede_eletrica.txt §5`, que é instrução ao modelo, não especificação validada) · `CONFIRMADA`.

## Por que a camada 1 existe (achado)

O parser do backend é **tolerante e falha em silêncio** (`services/orcamento_calc.py:46-122`). Cada
tolerância abaixo é uma entrada errada que hoje vira número sem aviso — é exatamente o que a camada 1 deve expor:

| Parser faz | Consequência silenciosa |
|---|---|
| CABOS com < 3 tokens: assume fase = 1 e/ou comprimento = 1 | metragem errada no orçamento |
| CABOS com fase/comprimento não numérico: mantém o padrão 1.0 (`except: pass`) | idem |
| OUTROS com quantidade não numérica: assume `1.0` | quantidade errada |
| OUTROS com número ímpar de tokens: o último é **descartado** (`while i + 1 < len`) | ativo some do orçamento |
| Ativo vazio: linha ignorada (`continue`) | linha sem efeito |
| Ativo sem correspondência na base: vai para `nao_encontrados` | só aparece após calcular |

## Camada 1 — Contrato (código) — IMPLEMENTADA (TASK-011)

`services/validacao_planilhas.py` · `POST /api/validacao/planilhas` · testes em `tests/`.

**Achado da TASK-011:** a tabela **Cabos** da tela usa `CAA 2 ABC 35 m` (o nome pode ter espaço) e o
frontend normaliza `CAA 2` → `CAA2` (`resumo.js: calcularQtdAtivos`). O parser do backend
(`extrair_ativo_cabo`) lê `parts[0]` e só serve ao **payload da Totalizadora** (`CAA2 1 35`). Por isso:
Outros reutiliza o parser do backend (`tokenizar_outros`, extraído sem mudar o cálculo); Cabos tem leitor
próprio que **espelha a normalização do frontend** (duplicação conhecida, ARCHITECTURE §11); e
`C1-BASE` roda sobre o **payload após as regras de conversão** (ADR-003), não sobre a tabela crua,
que daria falso positivo quando uma regra troca ou gera o ativo.

| id | escopo | verifica | sev. | status |
|---|---|---|---|---|
| C1-CABO-PARTE | cabos | ativo no formato `<qtd>-<ativo>` na tabela Cabos | erro | IMPLEMENTADA |
| C1-CABO-NUM | cabos | com ≥ 2 tokens, o último (comprimento) é numérico (aceita vírgula) | erro | IMPLEMENTADA |
| C1-OUT-ORFAO | outros | número ímpar de tokens: o último seria descartado no cálculo | erro | IMPLEMENTADA |
| C1-OUT-NUM | outros | quantidade não numérica (o cálculo assumiria 1) | erro | IMPLEMENTADA |
| C1-OUT-CABO | outros | ativo terminando em `<número> m` (formato de cabo) | erro | IMPLEMENTADA |
| C1-OP | ambas | operação ∈ `I, *I, R, *R, M, *M` (vazio também é erro) | erro | IMPLEMENTADA |
| C1-QTD-ZERO | ambas | comprimento/quantidade zero (RN-17 tira do cálculo) | aviso | IMPLEMENTADA |
| C1-DUP | outros | mesmo ativo repetido na linha | aviso | IMPLEMENTADA (severidade confirmada) |
| C1-BASE | payload | ativo não resolve na base técnica (cascata RN-06, com origem e projeto) | aviso | IMPLEMENTADA |

Não checado de propósito: valores válidos de FASE (não há lista confirmada), linha standalone de cabo
(`P50` sozinho é válido, herda a fase — RN-04), e linha sem ativo (o cálculo a ignora).
`C1-BASE` cobre só o "passo 5" da cascata (`nao_encontrados`); ativo achado com COMPONENTE vazio é
descartado em silêncio pelo cálculo e não é detectado.

Exemplos: `CAA 2 ABC 35 m` ok · `3-CFU` em Cabos → C1-CABO-PARTE · `CAA 2 35 m` em Outros → C1-OUT-CABO · `3-IP RECAL` → C1-OUT-ORFAO.

## Camada 2 — Regras de domínio (editáveis pelo admin)

**Camadas por projeto (TASK-018, ADR-005):** o `DEFAULT` é o padrão de todos os projetos; um projeto só guarda o que
difere (regras próprias, ajustes de campos, regras ocultas com −, grupos redefinidos). No painel admin cada regra mostra o
selo *padrão*, *ajustada neste projeto* ou *só deste projeto*. Quando o padrão evolui, os projetos acompanham, exceto onde
sobrescreveram. Editor visual: TASK-017. O texto abaixo descreve as regras da semente (TASK-013) e continua válido.

**IMPLEMENTADA (TASK-013)** — `services/regras_dominio.py` (motor e schema), `routers/validacao_regras.py`
(`/api/validacao/regras`), tabelas `regras_dominio`/`regras_dominio_historico`, editor e painel de teste no
painel admin. Fonte das regras: `prompt_rede_eletrica.txt §5` (o arquivo não foi alterado).
**Todas entram na semente com `ativa: false`**; o admin liga uma a uma. Unidade de avaliação: a **linha da
tabela Outros** (cada linha é um poste ou conjunto de ativos soltos); só o P50 cruza Cabos e Outros.
Por padrão as regras valem só para operações **I e *I** (decisão do usuário; editável por regra).
Tipos: `requer`, `proibe`, `nao_isolado`, `texto`, `minimo_total` (parâmetros no docstring do módulo).
As três regras de P50 viraram **uma só**, `C2-P50` (`minimo_total`), porque a exigência é uma soma
(2 m por trafo mono/bi + 6 m por trafo tri + 1 m por PR15): avaliadas separadamente, cada uma compararia o
total de P50 com um pedaço da exigência e daria alerta errado.

| id | escopo | regra | sev. sugerida | status | ambiguidade a resolver |
|---|---|---|---|---|---|
| C2-CFU-SUPL | outros | linha com `CFU` exige `1-SUPL` | aviso | IMPLEMENTADA | resolvida (ver decisões abaixo) |
| C2-CFU-EF | outros | `CFU` exige elo fusível `EF…` no mesmo poste | aviso | IMPLEMENTADA | resolvida (ver decisões abaixo) |
| C2-TR-EF | outros | poste com `TR…` e `EF…` **sem** chave (`CFU` ou `CFUR`) na linha | aviso | IMPLEMENTADA | — |
| C2-TR-PR15 | outros | poste com `TR…` exige `1-PR15`, exceto se houver `RPR` | aviso | IMPLEMENTADA | resolvida (ver decisões abaixo) |
| C2-TR-PR220-MONO | outros | trafo monofásico exige ≥ `2-PR220` (dobra com duas descidas) | aviso | IMPLEMENTADA | resolvida (ver decisões abaixo) |
| C2-TR-PR220-TRI | outros | trafo trifásico exige ≥ `3-PR220` (dobra com duas descidas) | aviso | IMPLEMENTADA | resolvida (ver decisões abaixo) |
| C2-P50 | cabos+outros | P50 total (Cabos, comprimento **bruto**) ≥ 2 m por trafo TR1xx/TR2xx + 6 m por trafo TR3xx + 1 m por PR15 (Outros) | aviso | IMPLEMENTADA | — |
| C2-POSTE10-MT | outros | poste de 10 m (`DT10/…`, `CV10…`) não pode ser usado em MT | erro | IMPLEMENTADA | resolvida (ver decisões abaixo) |
| C2-ESTR-ISOL | outros | estrutura `U3/N3/R3…` não pode estar sozinha no poste: exige outra estrutura MT ou `TR…` | aviso | IMPLEMENTADA | resolvida (ver decisões abaixo) |
| C2-POSTE-FMT | outros | poste no formato `DT…`/`CV…`, não `POSTE11`, `1-DT11/300`, `DT11` | info | IMPLEMENTADA | resolvida (ver decisões abaixo) |
| C2-BT-EXT | outros | linha com estrutura `SI` (`SI\d+`) exige `RA2`, salvo se houver estrutura `S#` (`S2`, `S4`, `S44`…) | aviso | IMPLEMENTADA | — |

### Decisões do usuário (2026-09-30) — entram na TASK-013

- **Poste em MT:** o poste é de MT se a linha de Outros tiver estrutura MT (U1–U4, N1–N4, T1–T3/TE,
  R1–R4 e variantes) ou chave/trafo de MT.
- **Tipo de trafo:** `TR1xx` monofásico · `TR2xx` bifásico (segue a regra do monofásico) · `TR3xx` trifásico.
- **Duas descidas (dobra PR220):** no poste do trafo há `1-SI4` **ou** `2-SI3` → duas descidas. Só `1-SI3`
  ou outra estrutura de BT → uma descida.
- **Estruturas isoladas:** `U3, N3, R3, T3` e variantes (`U3C`, `N3IV`…).
- **P50:** validar pelo **total da planilha** (P50 na tabela Cabos, operação I/*I, contra o exigido
  por todos os trafos e PR15 de Outros, operação I/*I). Um aviso único (`linha_id` "GERAL"), sem vínculo a
  poste. Metros = **comprimento bruto** do último token (`P50 ABC 2 m` conta 2 m, sem multiplicar pelas fases).
- **Chave e elo fusível:** `CFU` e `CFUR` exigem `EF…` no mesmo poste (`CFA` fora).
- **Poste `1-DT11/300`:** severidade **info** (o parser aceita). Duplicidade de ativo: **aviso**.

- **Estrutura isolada:** "outra estrutura MT" = **qualquer** estrutura MT, mesmo outra isolada (U3 + N3 no
  mesmo poste satisfaz a regra). Trafo (`TR…`) também vale.
- **Duas descidas:** `SI3` com quantidade **≥ 2** também conta (além de `1-SI4`).
- **CFU e SUPL:** basta `1-SUPL` (qtd ≥ 1) na linha do poste, em qualquer posição; vale só para `CFU`.

- **Operações:** as regras valem só para linhas de instalação (`I`, `*I`); linha `R`/`M` não é validada.

- **Trafo e elo (`C2-TR-EF`):** só `CFU` e `CFUR` contam como a chave que libera o elo no poste do trafo. `CFA`,
  `CL` e as de reinstalação/abertura (`RCFU`, `ACFU`…) **não** liberam.
- **Extensão BT (`C2-BT-EXT`):** definição do usuário: "se houver estrutura SI (SI3, SI4, SI1) adicionar RA2, caso não
  exista estrutura tipo S# (S2, S4, S44…)". Implementada como `requer` (`SI\d+` → `RA2`, exceto `S\d+`); `SI2`
  entra pelo mesmo padrão; severidade `aviso` (o catálogo previa `info`, mas a regra é prescritiva).

Sem pendências de definição: **todas as regras do catálogo estão `IMPLEMENTADAS`** (11 na semente; o P50 reúne três
do catálogo original). Regras novas da semente não chegam sozinhas a quem já tem regras salvas: o admin usa
"Adicionar regras novas da semente" (só acrescenta o que falta, desligado). Termo sem definição no GLOSSARY: "poste associado à linha"
(a TASK-013 define poste = linha de Outros que abre com DT/CV, sem inventar além disso).

## Camada 3 — IA (prompts salvos)

**`validar-planilhas` IMPLEMENTADA (TASK-014)** — `services/validacao_ia.py` + `POST /api/validacao/ia`.
`corrigir-planilhas` **IMPLEMENTADA (TASK-015)** — ver abaixo. Escopo: o que não é determinístico. Recebe as linhas + os achados das camadas 1–2 (para não repetir).
Mecânica: linhas com ativo, em lotes de 40 (limite 400; acima disso a IA revisa só as primeiras e a resposta traz
`truncado`); o prompt salvo do projeto (fallback `DEFAULT`) é usado com temperatura do cabeçalho (0) e o modelo do
cabeçalho + reservas; resposta em JSON; **item com `linha_id` que não estava no lote é descartado**, severidade
inválida vira `info`, item sem `problema` é descartado. A IA nunca corrige e nunca bloqueia: falha de rede, cota,
prompt desativado (`ativo: false`), sem chave ou JSON quebrado viram um `status` (`erro`/`parcial`/`desativado`/
`sem_chave`/`indisponivel`) com HTTP 200, e a tela mostra as camadas 1–2 do mesmo jeito.
Cota: com a chave padrão do sistema vale o limite de mensagens por minuto por usuário (429 com mensagem clara).
Sementes em `data/validacoes/`:

| id do prompt | modo | verifica |
|---|---|---|
| `validar-planilhas` | checar | coerência semântica entre ativos de uma mesma linha/poste; ativo que parece digitação errada de outro conhecido; combinações improváveis; ativo na tabela errada por sentido (não por formato) |
| `corrigir-planilhas` | corrigir | propõe correção **por linha** para achados existentes; nunca aplica; altera só o texto do ativo (nunca a operação) |

Saída obrigatória (JSON): `{"achados":[{"linha_id","tabela","severidade","regra","problema","sugestao"}]}`
para checar, e `{"correcoes":[{"linha_id","tabela","antes","depois","motivo"}]}` para corrigir.
Toda correção passa novamente pela camada 1 antes de ser mostrada ao usuário (ADR-004, TASK-015).

### Correção assistida (TASK-015)

`services/correcao_ia.py` + `POST /api/validacao/corrigir`. Só as linhas **citadas nos achados** (e editáveis:
`CABOS-<i>`/`OUTROS-<i>`; `GERAL` e Totalizadora ficam de fora) vão à IA, junto com os achados e a sugestão da IA de
revisão. A IA só propõe; cada proposta é reconferida no servidor e **descartada com o motivo** se: a linha não foi
enviada, não muda nada (ignorando espaço e caixa), repete a linha, ou o texto proposto **ainda reprova na camada 1**
(erro de formato). O "antes" exibido é sempre o texto real da linha. Aviso da camada 1 (ex.: ativo repetido) não
descarta, mas acompanha a proposta ("Atenção"). Na tela: o usuário marca linha a linha (nada vem marcado) ou usa
"Aceitar todas"; aplicar entra no histórico como **um passo** (um desfazer restaura o lote e refazer reaplica); se a
linha mudou depois da proposta, ela é ignorada e o usuário é avisado. Depois de corrigir, o orçamento não segue
sozinho: é preciso validar/montar de novo.

## Fluxo decidido (2026-09-30)

- Validação: botão sob demanda **ou** automática ao "montar orçamento", conforme radio (preferência local, `localStorage`, padrão desligado).
- Com achados **de erro ou aviso**, o sistema **pergunta "continuar mesmo assim?"** (só achados `info` não interrompem; decisão da implementação, ajustável). Se o usuário responder **não**, entra o fluxo de **correção** (camada 3, modo corrigir), com aceite por linha.
