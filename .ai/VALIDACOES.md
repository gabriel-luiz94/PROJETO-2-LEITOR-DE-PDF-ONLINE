# CATÁLOGO DE VALIDAÇÕES

> Lista de tudo que a validação das planilhas **Cabos** e **Outros** verifica. É a fonte de contexto
> das TASKs 011 (camada 1), 013 (camada 2), 012/014/015 (camada 3). Arquitetura: `ADR-004`.
>
> Regra 2 (RULES): **nenhuma regra de domínio é implementada ou ativada sem status `CONFIRMADA`.**
> Este arquivo é a lista para revisão linha a linha com o usuário.

## Legenda

- **Camada 1** — contrato, em código, **não editável**. **Camada 2** — regra de domínio, dados, **editável pelo admin**. **Camada 3** — IA, prompt salvo, editável pelo admin.
- **Severidade:** `erro` · `aviso` · `info`.
- **Status:** `DERIVADA DO CÓDIGO` (comportamento verificado em `orcamento_calc.py`) · `A CONFIRMAR` (vem de `prompt_rede_eletrica.txt §5`, que é instrução ao modelo, não especificação validada) · `CONFIRMADA`.

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

Fonte: `prompt_rede_eletrica.txt §5`. **Todas entram na semente com `ativa: false`.**
As ambiguidades abaixo precisam de resposta do usuário antes de virar regra.

| id | escopo | regra | sev. sugerida | status | ambiguidade a resolver |
|---|---|---|---|---|---|
| C2-CFU-SUPL | outros | linha com `CFU` exige `1-SUPL` | aviso | A CONFIRMAR | "antes da chave" é ordem dentro da linha ou só presença? |
| C2-CFU-EF | outros | `CFU` exige elo fusível `EF…` no mesmo poste | aviso | A CONFIRMAR | vale também para CFA/CFUR? |
| C2-TR-EF | outros | poste com `TR…` **sem** chave não deve ter `EF…` | aviso | A CONFIRMAR | — |
| C2-TR-PR15 | outros | poste com `TR…` exige `1-PR15`, exceto se houver `RPR` | aviso | A CONFIRMAR | qual é o par com a regra de P50 (PR15 → +1 m de P50, tabela Cabos)? |
| C2-TR-PR220-MONO | outros | trafo monofásico exige ≥ `2-PR220` (dobra com duas descidas) | aviso | A CONFIRMAR | como identificar mono/bi/trifásico pelo código (TR1xx/TR2xx/TR3xx)? como saber "duas descidas"? |
| C2-TR-PR220-TRI | outros | trafo trifásico exige ≥ `3-PR220` (dobra com duas descidas) | aviso | A CONFIRMAR | idem |
| C2-P50-MONO | cabos+outros | trafo monofásico: mín. 2 m de `P50` além dos dos PR15 | aviso | A CONFIRMAR | regra cruza Cabos e Outros: a que "poste" o P50 pertence? |
| C2-P50-TRI | cabos+outros | trafo trifásico: mín. 6 m de `P50` além dos dos PR15 | aviso | A CONFIRMAR | idem |
| C2-P50-PR15 | cabos+outros | +1 m de `P50` por cada `PR15` | aviso | A CONFIRMAR | idem |
| C2-POSTE10-MT | outros | poste de 10 m (`DT10/…`, `CV10…`) não pode ser usado em MT | erro | A CONFIRMAR | **como saber que o poste é de MT?** (presença de estrutura MT? cabo de média na mesma rede?) |
| C2-ESTR-ISOL | outros | estrutura `U3/N3/R3…` não pode estar sozinha no poste: exige outra estrutura MT ou `TR…` | aviso | A CONFIRMAR | quais códigos contam como "isolada"? qual lista de "estruturas MT"? |
| C2-POSTE-FMT | outros | poste no formato `DT…`/`CV…`, não `POSTE11`, `1-DT11/300`, `DT11` | erro | A CONFIRMAR | o parser aceita `DT11/300` sem `1-`; a forma `1-DT11/300` é proibida pelo prompt mas aceita pelo parser |
| C2-BT-EXT | outros | extensão BT (SI): passante SI1/SI2, fim SI3, amarração SI4 | info | A CONFIRMAR | é regra de validação ou só orientação de geração? |

### Decisões do usuário (2026-09-30) — entram na TASK-013

- **Poste em MT:** o poste é de MT se a linha de Outros tiver estrutura MT (U1–U4, N1–N4, T1–T3/TE,
  R1–R4 e variantes) ou chave/trafo de MT.
- **Tipo de trafo:** `TR1xx` monofásico · `TR2xx` bifásico (segue a regra do monofásico) · `TR3xx` trifásico.
- **Duas descidas (dobra PR220):** no poste do trafo há `1-SI4` **ou** `2-SI3` → duas descidas. Só `1-SI3`
  ou outra estrutura de BT → uma descida.
- **Estruturas isoladas:** `U3, N3, R3, T3` e variantes (`U3C`, `N3IV`…).
- **P50:** validar pelo **total da planilha** (P50 em metros na tabela Cabos, operação I, contra o exigido
  por todos os trafos e PR15 de Outros, operação I). Um aviso único, sem vínculo a poste.
- **Chave e elo fusível:** `CFU` e `CFUR` exigem `EF…` no mesmo poste (`CFA` fora).
- **Poste `1-DT11/300`:** severidade **info** (o parser aceita). Duplicidade de ativo: **aviso**.

Ainda em aberto: quais códigos contam como "outra estrutura MT" para a regra de estrutura isolada;
se 3 ou mais `SI3` também contam como duas descidas; ordem de `SUPL` "antes da chave" (presença basta?).
Termos ainda sem definição no GLOSSARY: "poste associado à linha".

## Camada 3 — IA (prompts salvos)

Escopo: o que não é determinístico. Recebe as linhas + os achados das camadas 1–2 (para não repetir).
Sementes em `data/validacoes/`:

| id do prompt | modo | verifica |
|---|---|---|
| `validar-planilhas` | checar | coerência semântica entre ativos de uma mesma linha/poste; ativo que parece digitação errada de outro conhecido; combinações improváveis; ativo na tabela errada por sentido (não por formato) |
| `corrigir-planilhas` | corrigir | propõe correção **por linha** para achados existentes; nunca aplica |

Saída obrigatória (JSON): `{"achados":[{"linha_id","tabela","severidade","regra","problema","sugestao"}]}`
para checar, e `{"correcoes":[{"linha_id","tabela","antes","depois","motivo"}]}` para corrigir.
Toda correção passa novamente pela camada 1 antes de ser mostrada ao usuário (ADR-004, TASK-015).

## Fluxo decidido (2026-09-30)

- Validação: botão sob demanda **ou** automática ao "montar orçamento", conforme radio (preferência local, `localStorage`, padrão desligado).
- Com achados, o sistema **pergunta "continuar mesmo assim?"**. Se o usuário responder **não**, entra o fluxo de **correção** (camada 3, modo corrigir), com aceite por linha.
