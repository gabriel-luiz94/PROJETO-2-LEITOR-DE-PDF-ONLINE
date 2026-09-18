# GLOSSÁRIO

> Termos extraídos do código, do seed da base técnica (`data/tabela_seed.csv`) e do prompt de
> domínio (`prompt_rede_eletrica.txt`).
>
> **Nenhuma definição foi inventada.** Termos cuja definição não pôde ser estabelecida com segurança
> estão marcados com `[DEFINIÇÃO NECESSITA CONFIRMAÇÃO DO USUÁRIO]`.
>
> ⚠️ O `prompt_rede_eletrica.txt` é uma instrução para o modelo de linguagem, não código executado.
> Definições que vêm só dele descrevem o que a IA foi instruída a fazer, não o que o sistema valida.

---

## 1. Conceitos estruturais do sistema

### ATIVO
**Definição:** item físico da rede identificado no projeto e usado como chave de entrada do cálculo
(ex.: `DT11/300`, `CAA 2 ABC 35 m`, `1-CFU`). Também é o nome da coluna `ativo` na base técnica.
**Onde é utilizado:** coluna `ativo` em `tabela_orcamento`/`tabela_orcamento_master`; campo `ativo`
nas tabelas Cabos/Outros/Totalizadora; primeiro passo da cascata de resolução.
**Observações:** o formato de escrita depende da ORIGEM (ver RN-03). Um mesmo ativo pode ter várias
linhas na base (uma por código consumido).

### DESC ATIVO / `desc_ativo`
**Definição:** descrição textual do ativo na base técnica (ex.: `POSTE CONCRETO DUPLO T 9X 150`).
**Onde é utilizado:** coluna da base técnica; passos 3 e 4 da cascata de resolução buscam por ela
(match exato e depois parcial); exibida como descrição na Totalizadora.

### COMPONENTE
**Definição:** agrupador da base técnica. Um ativo aponta para um componente, e o componente reúne
**todos** os códigos consumidos por aquele item (mão de obra + materiais).
**Onde é utilizado:** coluna `componente`; `orcamento_calc.py:206-231` — achado o ativo, toma-se o
componente do primeiro match e expandem-se todos os códigos daquele componente.
**Observações:** é o mecanismo central de expansão 1→N do orçamento. No seed, apenas a primeira
linha de um componente traz o `ATIVO` preenchido; as demais repetem `DESC ATIVO` e `COMPONENTE`
com o `ATIVO` vazio.

### CÓDIGO / `codigo`
**Definição:** código de serviço ou material da concessionária. É a saída final do orçamento.
**Onde é utilizado:** coluna `codigo`; chave de agrupamento `(codigo, mdo)` na soma; coluna do
resultado; reconhecido nos PDFs de REC pelo padrão de 6 dígitos (`upload.py:93`).

### DESC CODIGO / `desc_codigo`
**Definição:** descrição do código (ex.: `Poste Limpo (sem mat. ou equip.) 12 >=P<= 600`).
**Onde é utilizado:** coluna da base; campo do resultado; terceiro critério de ordenação (A–Z).

### FATOR I / `fator_i` (fator instalando)
**Definição:** multiplicador aplicado à quantidade quando a operação é de **instalação** (`I`/`*I`).
**Onde é utilizado:** `orcamento_calc.py:255` — `soma_i += qtd × fator_i`.
**Observações:** no passo 5 da cascata (ativo resolvido como código direto) é **forçado para 1.0**.

### FATOR R / `fator_r` (fator removendo)
**Definição:** multiplicador aplicado à quantidade quando a operação **não** é de instalação.
**Onde é utilizado:** `orcamento_calc.py:257` — `soma_r += qtd × fator_r`.
**Observações:** vale para `R`, `*R`, `M`, `*M` e qualquer outro valor — o código testa apenas se a
operação é `I` ou `*I`; tudo o mais cai no ramo de remoção.

### FATOR (genérico)
**Definição:** o termo aparece com **dois sentidos distintos** no sistema:
1. na base técnica → `fator_i`/`fator_r` (acima);
2. na Tabela de Regras de Conversão → coluna `FATOR`, multiplicador da quantidade aplicado pelo
   motor de regras antes do cálculo (`resumo.js:2968`).
**Observações:** não confundir. São camadas diferentes do pipeline.

### MDO
**Definição:** classificação da linha do orçamento. Valores observados: `MAO-DE-OBRA` e `MATERIAL`.
**Onde é utilizado:** coluna `mdo`; parte da chave de agrupamento; define a ordenação do resultado
(`get_mdo_rank`: vazio/`-`/`MAO-DE-OBRA` = 1, `MATERIAL` = 2, demais = 3).
**Observações:** quando ausente na linha, o código busca um fallback pelo `codigo` e, em último
caso, usa `"-"` — nunca fica vazio.

### FILTRO / `filtro`
**Definição:** coluna livre da base técnica, repassada ao resultado e usada na busca textual.
**Onde é utilizado:** `orcamento.py:search`, campo `filtro` do resultado.
**Observações:** o propósito de negócio do conteúdo dessa coluna
`[DEFINIÇÃO NECESSITA CONFIRMAÇÃO DO USUÁRIO]` — no seed ela está majoritariamente vazia.

### PROJETO
**Definição:** concessionária/praça à qual a linha da base técnica se aplica (`RONDONIA`, `PARAIBA`).
Cadastrado na tabela `projetos` com um `codigo` (`027`, `229`).
**Onde é utilizado:** coluna `projeto` da base; prioridade de match (RN-08); chave das regras de
conversão (`regras_conversao_<projeto_codigo>`).
**Observações:** aceita múltiplos valores separados por `/` (ex.: `PARAIBA/BAHIA`). Linha com
`projeto` vazio é genérica e vale para qualquer projeto.

### ORIGEM
**Definição:** procedência do registro na camada unificada. Dois valores: `CABOS` e `OUTROS`.
**Onde é utilizado:**
- atribuída no parse (`orcamento_calc.py:96` e `:121`);
- coluna `origem` da base técnica — restringe quais linhas podem atender a qual origem
  (`_matches_origem`: linha com `origem` vazia serve para qualquer uma);
- critério de match das regras de conversão;
- coluna editável na Tabela Totalizadora.
**Observações:** determina também o formato de escrita do ativo (RN-03). A coluna `origem` existe no
SQLite mas **não existe no schema versionado do Supabase** (ver ARCHITECTURE.md §4.2).

### CAMADA UNIFICADA
**Definição:** consolidação das tabelas Cabos e Outros em uma lista única antes do motor de regras,
preservando `origem` e um identificador de agrupamento `baseId` (`TOT-n`).
**Onde é utilizado:** `resumo.js:syncTotalizadora` — variável `rawItems`, que vira a Tabela
Totalizadora.
**Observações:** **o termo "camada unificada" não aparece no código.** É o nome conceitual da etapa;
no código ela é `rawItems` → `tableStates.totalizadora`.

### CONTEXTO / REGRAS
**Definição:** conjunto das etapas que interpretam e transformam registros antes do cálculo.
No código, são duas camadas distintas:
1. **classificação automática** — `computeRowLogic()`: cor, layer e texto → entidade/operação/ativo;
2. **Tabela de Regras de Conversão** — motor que casa `origem`+`op_de`+`ativo_de` e aplica
   `ADIÇÃO`/`SUBST`, `fator`, arredondamento e limites `val_min`/`val_max`.
**Onde é utilizado:** `resumo.js:181` (1) e `resumo.js:2918` (2).

### ENTIDADE
**Definição:** classificação do ativo. Valores fixos (`resumo.js:15`):
`0`, `CABO`, `CHAVE`, `TRAFO`, `ESTRUTURA`, `APOIO`, `IP`, `POSTE`, `RAMAIS`, `CERCA`.
**Onde é utilizado:** separa as tabelas (`CABO` → Cabos; demais ≠ `0` e ≠ `RAMAIS` → Outros);
vira a coluna `obs` na Totalizadora; enviada no payload do cálculo.
**Observações:** `0` significa "não classificado". `IP`, `APOIO`, `CERCA` e `RAMAIS` forçam a
operação para `I` (`resumo.js:297`). A entidade **não participa do cálculo** no backend —
`processar_calculo` usa apenas `ativo`, `operacao` e a origem.

### OPERAÇÃO
**Definição:** o que será feito com o ativo. Valores (`resumo.js:16`): `I`, `*I`, `R`, `*R`, `M`, `*M`.
- `I` = instalar · `R` = remover · `M` = manter
- `*` = **linha viva** (energizado / sem desligamento)
**Onde é utilizado:** derivada da cor e do layer; critério de match das regras; determina se a soma
vai para `soma_i` ou `soma_r`; coluna do resultado.
**Observações:** no backend, apenas `I` e `*I` contam como instalação; todo o resto (inclusive `M`)
soma em `soma_r` com `fator_r`. O tratamento de `M` como remoção
`[DEFINIÇÃO NECESSITA CONFIRMAÇÃO DO USUÁRIO]` — pode ser intencional ou um efeito colateral.

### LINHA VIVA
**Definição:** intervenção com a rede energizada, sem desligamento. Marcada pelo prefixo `*` na
operação e pelo layer `01_LV` (e `01_RETENS_LV`) no DXF.
**Onde é utilizado:** `resumo.js:304` (layer `01_LV` → prefixa `*`); `orcamento_calc.py:111`
(no parse de OUTROS, `*` é convertido em `-`, representando quantidade negativa);
`prompt_rede_eletrica.txt:8-10`.
**Observações:** há um botão "LINHA VIVA" em `resumo.html:592` marcado como **"Função futura"** —
não implementado. Existe também o ativo `GLV` (grampo de linha viva) na base, que é outra coisa.

### REC
**Definição:** registro da obra orçada, identificado por `numero_obra` e salvo como JSON em
`historico_rec`. Também é o nome do arquivo exportado (`.rec`).
**Onde é utilizado:** `routers/recs.py`, `POST /api/importar-rec-pdf`, `resultado_orcamento.html`.
**Observações:** o significado da sigla `[DEFINIÇÃO NECESSITA CONFIRMAÇÃO DO USUÁRIO]`.

### OBRA
**Definição:** projeto em edição, salvo com `id`, `nome`, `data` e `dados_json` por usuário.
**Onde é utilizado:** tabela `obras`, `routers/obras.py`.
**Observações:** distinto de REC — obra é o trabalho em andamento; REC é o orçamento fechado.

### TABELA MASTER
**Definição:** base técnica oficial mantida pelo admin, espelhada entre Supabase
(`tabela_orcamento_master`) e SQLite local.
**Onde é utilizado:** `sync_service.sync_tabela_master()`, `admin.py`.
**Observações:** tem precedência sobre o seed, mas perde para as edições locais do usuário (RN-11).

### QTD ATIVOS / `qtdAtivos`
**Definição:** número de ativos que uma linha de cabo representa — normalmente o número de fases.
**Onde é utilizado:** `resumo.js:calcularQtdAtivos`; multiplicador no cálculo
(`orcamento_calc.py:81-94`); número de repetições da linha na camada unificada (`resumo.js:2861`).
**Observações:** ver RN-04 para as exceções (`M…`, `CAZ`, `CAA2`, `CA4`) e a herança de fase.

---

## 2. Ativos de rede elétrica

> Fontes: `prompt_rede_eletrica.txt` (lista de códigos válidos e interpretações) e
> `data/tabela_seed.csv` (descrições reais da concessionária).

### 2.1 Postes

#### POSTE
**Definição:** suporte físico da rede. É uma das ENTIDADES do sistema.
**Onde é utilizado:** classificado quando o texto contém `DT` ou `CV` **e** `/`
(`resumo.js:269`, `1022`); modal "Gerador de Códigos — POSTES E ESTRUTURAS" (`resumo.js:2085`).
**Observações:** postes entram pela tabela **OUTROS**, não por uma origem própria. Na notação de
ativos, o poste inicia a linha **sem quantidade** (`prompt:13-14`).

#### DT
**Definição:** poste de concreto **Duplo T**. Notação `DT<altura>/<resistência>`
(ex.: `DT11/300` = 11 m, 300 daN). Seed: `POSTE DISTR CONCR DT 11M 300DAN …`.
**Onde é utilizado:** `resumo.js:269`, `orcamento_calc.py:106`, seed.
**Observações:** padrão preferencial quando o acesso é normal (`prompt:15`). Postes de 10 m não
podem ser usados em MT — regra do prompt, **não validada em código**.

#### CV
**Definição:** poste circular/alternativo ao DT, usado quando o acesso é difícil
("bandolado", "sem acesso", "local remoto" — `prompt:16`). Notação `CV11/600`.
**Onde é utilizado:** mesmos pontos do DT.
**Observações:** o que exatamente a sigla significa `[DEFINIÇÃO NECESSITA CONFIRMAÇÃO DO USUÁRIO]`.

### 2.2 Cabos e condutores

#### CABO
**Definição:** entidade dos condutores. Formato do ativo: `[ATIVO] [FASE] [COMPRIMENTO] m`.
**Onde é utilizado:** `tableStates.cabos`, origem `CABOS`.

#### CA / CAA / CAL / CAZ / CU / P / M…
**Definição (pelo seed e pelo prompt):**
- `CA` — cabo de alumínio nu (seed: `CABO ALUM CA NU 1/0 AWG 1F POPPY` para `CA10`)
- `CAA` — cabo de alumínio com alma de aço `[DEFINIÇÃO NECESSITA CONFIRMAÇÃO DO USUÁRIO]`
- `CAL`, `CAZ` — prefixos reconhecidos como cabo (`resumo.js:275`); descrição exata
  `[DEFINIÇÃO NECESSITA CONFIRMAÇÃO DO USUÁRIO]`
- `CU` — cabo de cobre (`CU2`, `CU4`, `CU10`, `CU40`, `CU70`, `CU300`)
- `P` seguido de bitola (`P16`, `P25`, `P35`, `P50`, `P70`, `P95`, `P120`, `P150`, `P185`, `P240`);
  variantes `DP` (ex.: `P50DP`)
- `M` — cabo multiplex, notação `M<fases>x1x<seção>+<neutro>` (ex.: `M1x1x16+16`, `M3x1x185+120`)
**Onde é utilizado:** `calcularQtdAtivos`, `autoClassifyEntidade`, seed, prompt.
**Observações:** `M…` e `CAZ…` sempre contam como **1** ativo, independentemente da fase (RN-04).

#### FASE
**Definição:** identificação das fases do trecho de cabo (ex.: `A`, `AB`, `ABC`). O **comprimento**
da string é o número de fases e, por padrão, a quantidade de ativos.
**Onde é utilizado:** `calcularQtdAtivos`, `extrairFase`, formato `[ATIVO] [FASE] [COMPRIMENTO]`.

#### Linha standalone
**Definição:** linha de cabo que contém só o ativo (ou ativo + comprimento), sem fase. Herda a fase
da **linha seguinte** da tabela.
**Onde é utilizado:** `isStandaloneLine`, `recalcAllQtdAtivos` (percorre de baixo para cima).

### 2.3 Transformadores

#### TRAFO / TR
**Definição:** transformador de distribuição. Entidade `TRAFO`. Notação `TR<fases><potência>`
(seed: `TR345M` = `TRANSF DISTR AER OMI 3F 13,8KV 220/127V 45KVA`; `TR115M` = `1F/N … 15KVA`).
**Onde é utilizado:** `resumo.js:288-292`; `processAtivoFormula` converte `TR-1-`/`TR-2-`/`TR-3-`
em `1-TR1`/`1-TR2`/`1-TR3`; lista de códigos válidos no prompt (`TR105`…`TR3300A`).
**Observações:** o primeiro dígito indica o número de fases (1, 2 ou 3) e os seguintes a potência
em kVA — inferido da lista e das descrições do seed; confirmar com o usuário para casos com sufixo
`A`. A notação **`TR+1`, `TR+2`, `TR+3` não existe no código nem no prompt.**

#### RTR1 / RTR2 / RTR3
**Definição:** **reinstalação** de transformador (1, 2 ou 3 — número de fases).
**Onde é utilizado:** gerado automaticamente por `computeRowLogic` quando o texto casa `TR-<n>` com
cor preta em layer de retensionamento (`resumo.js:244-250`); `autoClassifyEntidade:1015`;
prompt (`RTR1, RTR2, RTR3` — "Reinstalação de Trafos").

#### ATR1 / ATR2 / ATR3
**Definição:** **abertura** de trafos (prompt:47-48).
**Onde é utilizado:** apenas na lista de ativos válidos do prompt; sem tratamento especial no código.

### 2.4 Chaves, elos e proteção

#### CHAVE
**Definição:** entidade das chaves de manobra. Classificada quando o ativo contém `-CF` ou `-EF`
(`resumo.js:287`).

#### CFU / CFA / CFUR / CL
**Definição (prompt:35-36):** chaves e acessórios de manobra. `CFU` = chave fusível unipolar
(interpretação de "chave" no prompt:86). `CFA`, `CFUR`, `CL`
`[DEFINIÇÃO NECESSITA CONFIRMAÇÃO DO USUÁRIO]`.
**Onde é utilizado:** `processAtivoFormula` converte `-100A-` em `-CFU …-EF` e `3CF-400A` em
`3-CFA`, `3CL-300A` em `3-CFU 3-CL`.
**Observações:** regra do prompt — toda `CFU` deve vir acompanhada de `1-SUPL`. **Não validada em código.**

#### RCFU / RCFA
**Definição:** reinstalação de chave (prompt:38-39).
**Onde é utilizado:** gerado por `computeRowLogic` quando `<n>-100A` aparece em preto em layer de
retensionamento (`resumo.js:254-260`).

#### ACFU / ACFA
**Definição:** abertura de chaves (prompt:45-46). Sem tratamento especial no código.

#### EF / Elo fusível
**Definição:** elo fusível. Notação `EF<valor>` (`EF05H`, `EF1H`, `EF3H`, `EF6K`, `EF10K`, `EF40K`…).
**Onde é utilizado:** lista `elosFusivel` em `resumo.js:207`
(`0,5H`,`1H`,`2H`,`3H`,`5H`,`6K`,`8K`,`10K`,`12K`,`15K`,`25K`,`30K`,`40K`); quando um desses
valores aparece isolado em vermelho/cinza, o código procura um `<1|3>-100A` nas 10 linhas vizinhas
da **mesma página** e monta `<qtd>-EF<valor>`.
**Observações:** os sufixos `H` e `K` são classes de elo; o significado exato
`[DEFINIÇÃO NECESSITA CONFIRMAÇÃO DO USUÁRIO]`.

#### PR15 / PR220 / PRMT / PRBT / PR220T / PRBTT / RPR
**Definição:** para-raios e proteção (prompt:74-75). `PR15` = para-raio de MT;
`PR220` = para-raio de BT; `RPR` = reinstalação de para-raio.
**Observações:** regras do prompt (todo trafo leva `1-PR15` + 2 ou 3 `PR220`; `RPR` substitui o
`PR15`; cada `PR15` adiciona 1 m de `P50`) **não são validadas em código**.

### 2.5 Estruturas

#### ESTRUTURA
**Definição:** entidade dos arranjos de montagem no poste. Classificada por padrão de texto
`<n>-<código alfanumérico>` (`resumo.js:299-302`, `1040-1042`).

#### U1, U2, U3, U4
**Definição (prompt:55):** estruturas **monofásicas de MT**.
**Definição do número (prompt:101-105):** `1` ou `2` = passante; `3` = fim de rede ou derivação;
`4` = amarração.
**Onde é utilizado:** lista de ativos válidos do prompt; dados da base técnica.
**Observações:** **não há lógica de código específica para U1–U4** — são resolvidos pela cascata
genérica. Regra do prompt: estruturas do tipo `3` não podem existir sozinhas no poste.

#### U23, U32
**Definição:** listadas em `prompt_rede_eletrica.txt:60` sob "Outras", junto com `U1C`–`U4C`,
`N1IV`–`N4IV`, `P1`, `P3`, `P4`, `B1`–`B4`, `S0`–`S5`, `S11`–`S55`, `CE1`–`CEJ2BE`.
**Onde é utilizado:** apenas essa menção. Sem tratamento em código.
**Observações:** `[DEFINIÇÃO NECESSITA CONFIRMAÇÃO DO USUÁRIO]` — o prompt não explica o que
distingue `U23` de `U32`, nem se a numeração composta indica transição entre os tipos 2 e 3.
**Não assuma** que `U23` e `U32` são equivalentes ou simétricos.

#### N1–N4, T1–T3/TE, R1–R4, SI1/SI3/SI4
**Definição (prompt:56-59):**
- `N1`–`N4` — trifásicas normais de MT
- `T1`, `T2`, `T3`, `TE` — trifásicas tipo T de MT
- `R1`–`R4` — bifásicas de MT
- `SI1`, `SI3`, `SI4` — BT multiplex (`SI1`/`SI2` passante, `SI3` fim/derivação, `SI4` amarração)
**Observações:** mesma convenção numérica de U1–U4. Sem tratamento específico em código.

### 2.6 Apoio, aterramento e serviços

#### APOIO
**Definição:** entidade dos serviços de apoio civil. Marcadores: `-ROCO`, `-RECAL`, `-BASE`,
`-CAVA`, `PODA` (`resumo.js:235`, `1011`).
**Observações:** existe detecção de **conflito de apoio** — se o texto contém `APOIOS`, `LARGURA`,
`BASE`+`CALÇADA`, ou mais de um marcador, a classificação como APOIO é **descartada**
(`hasApoioConflict`).

#### ROCO / RECAL / BASE / CAVA / PODA
**Definição pelo que gera cada um (`processAtivoFormula`):**
- `ROCO` — roçado; `"<n> METROS"` → `<n>-ROCO`
- `RECAL` — recomposição de calçada; texto com `CALÇADA`/`RECAL`/`REC. CAL` → `<qtd>-RECAL`
- `BASE` — concretagem de base; texto com `CONC` + `BASE` → `<qtd>-BASE`
- `CAVA` — abertura de cava; texto com `COMPRESSOR` ou `CAVA` → `<qtd>-CAVA`
- `PODA` — poda de árvore; variantes `PODA P/M/G` e `PODAS` são todas normalizadas para `PODA`
**Observações:** `MANILHAR`, `MD/LEVE`, `PODAU`, `PODAR` aparecem na lista do prompt sem definição.

#### TERRA1 / TERRA3
**Definição:** aterramento. Interpretação do prompt: quando não especificado, usar `TERRA3`
(3 hastes) — `prompt:90`. Logo, o dígito indica o número de hastes.
**Onde é utilizado:** lista do prompt; dados da base.

#### IP
**Definição:** iluminação pública. É uma das ENTIDADES e também um código de ativo.
**Onde é utilizado:** `processAtivoFormula` gera `<qtd>-IP`; entidade `IP` força operação `I`;
entra também no modal RAMAIS.

#### RAMAIS
**Definição:** entidade dos ramais de ligação. Detectada pelo padrão `\sRS\s+[MT]\s` no texto
original (`resumo.js:264`, `1017`).
**Onde é utilizado:** as linhas `RAMAIS` **não vão para a tabela Outros** — são desviadas para o
modal RAMAIS (`window._ramaisData`), junto com `IP` e `APOIO` cujo texto case
`REC.*CAL[CÇ]ADA` ou `CONC.*BASE`.
**Observações:** o que `RS M` e `RS T` significam `[DEFINIÇÃO NECESSITA CONFIRMAÇÃO DO USUÁRIO]`
(provavelmente ramal de serviço monofásico/trifásico — **confirmar**).

#### RABICHO
**Definição:** aparece **apenas** como descrição de código no seed
(`RABICHO DE LIGAÇÃO LE (SDBT) POR UN`, `RABICHO DE LIGAÇÃO LM (SDBT) POR UN` — códigos 64072/64073).
**Onde é utilizado:** dados da base técnica; **nenhuma lógica de código**.
**Observações:** `[DEFINIÇÃO NECESSITA CONFIRMAÇÃO DO USUÁRIO]`. As siglas `LE`, `LM` e `SDBT`
também não estão definidas em lugar nenhum do repositório.

#### CERCA
**Definição:** entidade classificada quando o texto contém `FIOS` (`resumo.js:286`, `1034`).
`processAtivoFormula` converte `"<texto> N FIOS"` em `1-<texto>F`. Prompt lista `CERCA (3F–16F)`.

#### AF / AF2 / AFASTADOR
**Definição:** afastador. `processAtivoFormula` retorna `1-AF` para qualquer texto contendo
`AFASTADOR`; `1-AF` é classificado como `ESTRUTURA`.

#### FLY / RFLY
**Definição:** identificados pelo texto contendo `FLY` em vermelho (`resumo.js:228-232`):
- `REF` + `FLY` → `1-RFLY`, operação `I`
- `DESF` + `FLY` → `1-FLY`, operação `R`
- `INST` + `FLY` → `1-FLY`, operação `I`
**Observações:** o que é um "FLY" `[DEFINIÇÃO NECESSITA CONFIRMAÇÃO DO USUÁRIO]`
(possivelmente *flying tap*/derivação aérea — **confirmar**).

#### SUPL / 1SUPL / GLV / ISOL / ISOLP / KITP / KITINT / ESTAI / BANDOL
**Definição:** listados no prompt como aterramento e acessórios (`prompt:78`). `GLV` aparece no seed
como `GRAMPO DE LINHA VIVA`. Os demais `[DEFINIÇÃO NECESSITA CONFIRMAÇÃO DO USUÁRIO]`.
**Observações:** `SUPL` é exigido junto de toda `CFU` pela regra do prompt (não validada em código).

---

## 3. Termos solicitados que NÃO existem no projeto

Verificado por busca em todo o repositório (código, prompt, seed, HTML):

| Termo | Situação |
|---|---|
| **K1** | Não encontrado em nenhum arquivo. `[DEFINIÇÃO NECESSITA CONFIRMAÇÃO DO USUÁRIO]` |
| **L1**, **L2** | Não existem como ativos. As ocorrências de "L1"/"L2" no repositório são partes de outras strings (`CL1/4/5` em descrições de fio, `CL2` = classe de poste/trafo no seed). `[DEFINIÇÃO NECESSITA CONFIRMAÇÃO DO USUÁRIO]` |
| **TR+1, TR+2, TR+3** | Notação não existe. O que existe é `TR1/TR2/TR3`, `RTR1–3`, `ATR1–3`. `[DEFINIÇÃO NECESSITA CONFIRMAÇÃO DO USUÁRIO]` |
| **rabicho** | Existe só como texto de descrição de código no seed (ver acima), sem lógica associada |
| **CAMADA UNIFICADA** | Conceito existe (`rawItems`), o **termo** não aparece no código |

Se algum desses termos for uma regra de negócio real ainda não implementada, ele deve virar uma
tarefa em `.ai/tasks/`, não uma suposição na documentação.

---

## 4. Termos de infraestrutura

### APP_MODE
**Definição:** modo de operação: `desktop` (local/`.exe`) ou `server` (nuvem).
**Onde é utilizado:** `config.py:84`, CORS em `app.py`, `connectivity_monitor`, `update.py`.

### IS_FROZEN
**Definição:** verdadeiro quando rodando como executável PyInstaller (`sys.frozen`).
**Onde é utilizado:** resolução de caminhos; **trava de segurança** de `/extract-local`.

### SEED
**Definição:** carga inicial da base técnica a partir de `data/tabela_seed.csv` (2.328 linhas,
delimitador `;`, cabeçalho `ATIVO;DESC ATIVO;COMPONENTE;PROJETO;MDO;CODIGO;DESC CODIGO;FATOR I;FATOR R;FILTRO`).
**Onde é utilizado:** `database.py:init_db` — só roda se `tabela_orcamento` estiver vazia.

### SYNC MASTER
**Definição:** download paginado (1000/lote) da `tabela_orcamento_master` do Supabase para o SQLite,
em transação que apaga e reinsere tudo.
**Onde é utilizado:** `sync_service.sync_tabela_master`, disparado por `GET /api/orcamento/dados` e
`POST /api/health/sync-master`.
