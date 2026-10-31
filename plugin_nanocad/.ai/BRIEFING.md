# Briefing — Plugin nanoCAD (C#) para o Leitor de Projetos Online Pro

> Status: **ideação aprovada só em mockups. Nada foi implementado.** Este briefing serve para iniciar o chat do plugin. Antes de codar, confirmar com o usuário os "pontos em aberto" no fim.

## 1. Contexto
- Projeto: `gabriel-luiz94/PROJETO-2-LEITOR-DE-PDF-ONLINE` (FastAPI + SQLite/Supabase + JS vanilla). Aplicação online que lê projetos de rede elétrica (DXF/PDF), vincula cabos e estruturas e gera orçamento.
- Fluxo do repositório: plano → implementar com testes e docs (`.ai/STATE.md`, `.ai/CHANGELOG.md`, `.ai/tasks/TASK-NNN-*.md`, `data/manual_regras_e_ajustes.md`) → PR só quando o usuário pedir (merge_method "merge", push normal, sem force). Respostas em português.
- Hoje o leitor roda na aplicação (JS ⇄ Python em paridade: `static/vinculacao_regras.js` ⇄ `services/vinculacao_regras.py`, `static/ativo_parametros.js` ⇄ `services/ativo_parametros.py`, `static/anotacao_outros.js`/`anotacao_cabos.js` ⇄ `services/*.py`, `services/orcamento_calc.py::_preparar_parametros`).
- Seed de regras de Pernambuco (081): `data/regras_leitor_*_pernambuco_seed.json`, `data/vinculacao_pernambuco_seed.json`.

## 2. Objetivo do plugin
Plugin **C# (.NET)** para **nanoCAD 26** (e **nanoCAD 5**, a confirmar) que roda dentro do desenho e produz uma **obra totalmente parametrizada** (`.obra.json`) para a aplicação online ou o modo autônomo calcular o orçamento. **O cálculo não roda no CAD** (por ora, com espaço para o futuro).

## 3. Fluxo
1. Ao abrir o plugin (comando `ORCAMENTO`), diálogo com **Construir do zero**, **Ler projeto** e **Continuar obra salva**; escolha das regras do projeto (ex.: Pernambuco 081).
2. **Construir do zero:** o usuário seleciona postes + anotações e vãos + anotações. Para projetos sem padrão de layers/cores.
3. **Ler projeto:** leitor instantâneo por **layer + cor** (a cor vem antes, a layer confirma, conforme regras do projeto) monta postes, vãos e vínculos. Sobra só **conferência e ajustes** (filtros Ok/Conferir/Erro).
4. **Numeração:** cada poste (P1…) e cada vão (V1…) tem número, exibido no desenho.
5. **Vínculo vão⇄poste sempre visível** (linhas/marcas), também usado internamente com as estruturas.
6. **Pop-up** ao clicar num vão ou poste: **entrada** (lida do desenho), **saída** (para a obra) e **ajustes** (operação etc.).
7. **Ver vínculos** (tabela: vão, anotação, operação, poste A/estrutura, poste B/estrutura, estado) e **Ver ativos com parâmetros preenchidos** (ex.: `N4(9400D, P, 3XA35P_7.94_-13.8KV, ?)`, `?` = pendente).
8. **Salvar / Atualizar / Reler.** Reler = **reset do zero** (lê tudo de novo).
9. **Persistência nos dois lugares:** dentro do DWG e em arquivo `<nome do desenho>.obra.json` na mesma pasta.
10. **Sem camada de marcas.** O plugin pode **alterar a layer dos textos** (ex.: ao mudar a operação, 01_INSTALANDO → 01_REMOVER), só com opção marcada no pop-up.
11. **Cálculo:** fora do CAD. Aplicação online (Carregar Obra) ou **modo autônomo** (vigia a pasta, acha o `.obra.json`) calcula; deve ser adicionado **motor de cálculo na aplicação** que salva o resultado (`.rec`/`.xlsx`/relatório de pendências) **na pasta do desenho**.

Definições: **vão = linha + anotação associada**.

## 4. Contrato de dados (prioridade nº 1)
O `.obra.json` precisa ser aceito por `/api/obras/importar` (mesmo formato de exportar/importar em `/api/obras/...`). Conferir no código antes de fixar o esquema. Campos relevantes já usados pelo vínculo:
- unidades de estrutura (id = linha; `linha+k/100` para a 2ª+ estrutura da mesma linha), `vinculoEstruturas`, `cabosFicticios`, `cabosFicticiosSlots`, `estruturasFicticias` (cabo), `participaVinculo` (operação R/`*R` fora do vínculo), `textoOriginal`, `_x/_y/_geo/_geoTipo/_poste`;
- parâmetros `cabo` (descrição completa), `tipocabo` (VALOR DE CABO), `cabo2`; `parametros` por regra de estrutura.
- Vínculo só entre cabos e estruturas com operação M/I (e variantes `*M`/`*I`); R e `*R` não entram.
**Exigir teste de contrato:** gerar `.obra.json` de exemplo (sem CAD) → importar via API → conferir.

## 5. Arquitetura sugerida (a decidir)
- Núcleo **independente do CAD** (leitura/regras/serialização) separado da camada nanoCAD (UI, seleção, gráficos), para testar fora do CAD.
- Onde roda a lógica de leitura: (a) servidor (plugin envia geometria e recebe a obra), (b) motor JS embutido reaproveitando `static/*.js`, (c) reescrita em C#. Recomendação inicial: manter **uma só fonte de regras** e evitar terceira cópia da lógica — avaliar (a) ou (b) antes de (c).
- Linhas de vínculo como **gráficos transitórios** (não gravam no DWG); dados da obra gravados no DWG (extension dictionary/XData) e no `.obra.json`.

## 6. Riscos
- **nanoCAD 5**: sem confirmação de API .NET. Plano B enxuto (script/COM/LISP) que só extrai para o mesmo `.obra.json`; conferência na aplicação.
- APIs do nanoCAD (.NET, C++ NRX, COM, LISP; SDK pelo nanoCAD Developers' Club) — **verificar documentação/SDK reais** antes de prometer recursos.
- O ambiente de desenvolvimento (Linux/nuvem) **não roda nanoCAD**: testes no CAD ficam com o usuário, em Windows. Entregar em passos pequenos testáveis por ele e manter o máximo testável sem CAD.
- Paridade das regras: qualquer mudança nas regras do leitor precisa continuar coerente entre JS e Python.

## 7. Pontos em aberto (perguntar ao usuário)
1. Onde roda a lógica de leitura (servidor / motor JS embutido / C#)?
2. nanoCAD 5 aceita plugin .NET? Qual build/SDK exato do 26 e do 5 o usuário tem?
3. Como o modo autônomo detecta o arquivo (pasta vigiada configurável?) e onde grava o resultado.
4. Seed de vínculo de Pernambuco: qual **Valor de cabo** usa o protegido (X ou P) e se a **U4** tem 1 ou 2 cabos.
5. Formato exato de armazenamento dentro do DWG.

## 8. Mockups (referência visual, só ilustrativos)
`A_inicio, B_construir, C_ler, D_vinculos, E_ativos, F_popup_layer, G_persistencia` (PNG, enviados ao usuário na sessão anterior; geradores em scratchpad `mock/gen3.py`).

## 9. Regra de ouro
Não implementar nada sem aprovação explícita; primeiro propor plano por etapas (contrato → núcleo testável → UI no CAD → motor/autônomo) e esperar o "pode implementar".
