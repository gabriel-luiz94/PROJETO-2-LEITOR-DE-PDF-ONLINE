# REQUISITOS DO PLUGIN

> Requisitos numerados, tirados do briefing ([`BRIEFING.md`](BRIEFING.md)) e dos mockups ([`../referencias/mockups/`](../referencias/mockups/)).
> Situação: ✅ implementado (testado sem CAD) · 🟡 implementado, falta conferir no nanoCAD · ⏳ pendente · ❓ em aberto (decisão do usuário).
> Ao mudar um requisito, atualize aqui **e** em `STATE.md`.

## RQ — Fluxo e telas

| ID | Requisito | Situação | Onde |
|---|---|---|---|
| RQ-01 | Comando `ORCAMENTO` abre a tela com **Construir do zero**, **Ler projeto** e **Continuar obra salva**; escolha das regras do projeto (mockup A) | 🟡 | `Plugin/Telas.cs › TelaInicio`, `Comandos.Orcamento` |
| RQ-02 | **Construir do zero**: selecionar postes (bloco + anotação) e vãos (linha + anotação); cada um ganha número; o plugin sugere os postes das pontas e a anotação, o usuário confirma ou troca (mockup B) | 🟡 | `ORC_POSTE`, `ORC_VAO`, `Nucleo/Construtor.cs` |
| RQ-03 | **Ler projeto**: leitor por **layer + cor** (a cor decide, a layer confirma) monta postes, vãos e vínculos; sobra a conferência com filtros Ok/Conferir/Erro (mockup C) | ✅ núcleo · 🟡 CAD | `Nucleo/LeitorProjeto.cs`, `Plugin/PainelObra.cs` |
| RQ-04 | Numeração: cada poste (P1…) e vão (V1…) tem número **exibido no desenho** | ✅ núcleo · 🟡 CAD | `LeitorProjeto.Numerar` (segue a rede), `Plugin/Marcas.cs` |
| RQ-05 | Vínculo vão ⇄ poste **sempre visível** (linhas/marcas) e usado internamente com as estruturas | ✅ núcleo · 🟡 CAD | `Nucleo/Vinculador.cs`, `Marcas` (tracejado ciano) |
| RQ-06 | **Pop-up** ao clicar num vão/poste: entrada (lida do desenho), saída (para a obra), ajustes (operação…) (mockup F) | 🟡 | `Plugin/Telas.cs › TelaElemento`, `ORC_EDITAR` |
| RQ-07 | **Ver vínculos** (vão, anotação, operação, poste A/estrutura, poste B/estrutura, estado) (mockup D) | ✅ núcleo · 🟡 CAD | `Relatorios.Vinculos`, `ORC_VINCULOS` |
| RQ-08 | **Ver ativos com parâmetros** (`N4(9400D, P, 3XA35P_7.94_-13.8KV, ?)`, `?` = pendente, em destaque) (mockup E) | ✅ núcleo · 🟡 CAD | `Relatorios.Ativos`, `ORC_ATIVOS` |
| RQ-09 | **Salvar / Atualizar / Reler**; Reler = do zero (descarta ajustes); Atualizar mantém os ajustes | ✅ núcleo · 🟡 CAD | `ORC_SALVAR`, `ORC_ATUALIZAR`, `ORC_RELER`, `Construtor.ReaplicarAjustes` |
| RQ-10 | Persistência nos **dois lugares**: dentro do DWG e em `<desenho>.obra.json` na mesma pasta (mockup G) | ✅ formato · 🟡 CAD | `Plugin/PersistenciaDwg.cs`, `ObraArquivo` |
| RQ-11 | **Sem camada de marcas** (gráficos transitórios); o plugin pode **mudar a layer do texto** ao trocar a operação, só com a opção marcada no pop-up (marcada por padrão) | 🟡 | `Marcas` (TransientManager), `Comandos.TrocarLayerDoTexto`, `RegrasProjeto.LayerDaOperacao` |
| RQ-12 | Vão = **linha + anotação associada** | ✅ | `LeitorProjeto.CortarEmVaos` / `AssociarTextos` |

## RQ — Dados e contrato

| ID | Requisito | Situação | Onde |
|---|---|---|---|
| RQ-20 | O `.obra.json` é aceito por `POST /api/obras/importar` (mesmo formato de exportar/importar) | ✅ | `ObraArquivo.Envelope`; `tests/test_plugin_nanocad_contrato.py` |
| RQ-21 | O arquivo leva o vínculo: unidades de estrutura (id = linha; linha + k/100 para a 2ª+), `vinculoEstruturas`, `textoOriginal`, `_x/_y/_geo/_geoTipo/_poste` — e a importação **não descarta** esses campos | ✅ | `ObraArquivo.Dados`; aplicação: `services/obras_arquivo.py › CAMPOS_VINCULO` (TASK-103 do principal) |
| RQ-22 | Parâmetros `cabo` (descrição completa), `tipocabo` (VALOR DE CABO), `cabo2`; `parametros` por regra de estrutura | ✅ | `RegrasVinculo.Parametros`; conferido contra `parametros_por_estrutura` da aplicação |
| RQ-23 | Vínculo só entre cabos e estruturas com operação I/M (e `*I`/`*M`); R e `*R` ficam fora | ✅ | `RegrasVinculo.ParticipaVinculo`, `Vinculador` |
| RQ-24 | **Teste de contrato**: `.obra.json` de exemplo gerado sem CAD → importado via API → conferido | ✅ | `exemplos/Projeto_Rua_das_Flores.obra.json`, `tests/test_plugin_nanocad_contrato.py` |
| RQ-25 | `cabosFicticios`, `cabosFicticiosSlots`, `estruturasFicticias` (fictícios) | ⏳ | a importação já aceita; o plugin ainda não cria fictícios |

## RQ — Plataforma

| ID | Requisito | Situação | Onde |
|---|---|---|---|
| RQ-30 | DLL pronta para o APPLOAD do **nanoCAD 26** nesta pasta | ✅ (compila contra a API real) · 🟡 carregar no CAD | `appload/nanoCAD26/OrcamentoNanoCAD.dll` |
| RQ-31 | **nanoCAD 5** (.NET, confirmado pelo usuário): DLL gerada nesta pasta | 🟡 | `build_nanocad5.bat` → `appload/nanoCAD5/` (compila no Windows do usuário) |
| RQ-32 | Núcleo independente do CAD, testável fora dele | ✅ | `src/Nucleo/`, `tests/Nucleo.Testes` |
| RQ-33 | Uma só fonte de regras (sem terceira cópia das REGRAS) | ✅ regras · ⚠️ motor portado | ADR-001 |

## RQ — Ligação com a aplicação (TASK-003)

| ID | Requisito | Situação | Onde |
|---|---|---|---|
| RQ-50 | Conectar à aplicação (localhost ou online) com login; senha não guardada, token renovado sozinho | ✅ núcleo · 🟡 CAD | `ClienteAplicacao`, `Conexao`, `TelaConexao`, `ORC_CONECTAR` |
| RQ-51 | Usar as regras do leitor e de vínculo **gravadas na aplicação** (com cópia guardada para abrir sem ela) | ✅ · 🟡 CAD | `GET /api/plugin/regras`, `Conexao.RegrasEmCache` |
| RQ-52 | Mudança na aplicação (regras, vínculo, conversão, base) **aparece no plugin em poucos segundos**, sem reler o DWG | ✅ · 🟡 CAD | `GET /api/plugin/versao`, `Sincronizador` (intervalo padrão 5 s) |
| RQ-53 | Apontar **estruturas/cabos/postes não cadastrados na base de orçamento** (Conferir), pelo mesmo caminho do cálculo | ✅ · 🟡 CAD | `POST /api/plugin/conferir`, `ConferenciaBase.Aplicar` |
| RQ-54 | Sem conexão: segue com o que tem e tenta de novo; sessão vencida pede login | ✅ · 🟡 CAD | `Sincronizador.Aplicar` |
| RQ-55 | `ORC_SINCRONIZAR`: baixar agora e conferir, com a lista do que está fora da base | 🟡 | `Comandos.CmdSincronizar` |

## RQ — Cálculo (fora do CAD)

| ID | Requisito | Situação |
|---|---|---|
| RQ-40 | Modo autônomo da aplicação acha o `.obra.json` (pasta vigiada) e calcula | ⏳ |
| RQ-41 | Motor de cálculo na aplicação grava o resultado (`.rec` / `.xlsx` / relatório de pendências) **na pasta do desenho** | ⏳ |
| RQ-42 | Futuro: o plugin abre/mostra o resultado sem calcular | ⏳ |

## Pontos em aberto (do briefing)

| # | Ponto | Situação |
|---|---|---|
| 1 | Onde roda a lógica de leitura | **decidido**: no CAD, em C#, regras como dados + paridade testada (ADR-001). Revisável. |
| 2 | nanoCAD 5 aceita .NET? Builds/SDK exatos | 5 aceita .NET (usuário). 26: API `hostmgd`/`hostdbmgd` net6.0 (pacote ModPlus). ❓ build exato do 5 e se o 26 carrega a DLL sem ajuste. |
| 3 | Como o modo autônomo detecta o arquivo e onde grava o resultado | ❓ (RQ-40/41) |
| 4 | Seed de vínculo de Pernambuco: VALOR DE CABO do protegido (X ou P) e se a U4 tem 1 ou 2 cabos | ❓ hoje o seed diz **P** e U4 com **2** cabos; o plugin usa o seed como está |
| 5 | Formato dentro do DWG | **decidido**: dicionário de objetos nomeados `ORCAMENTO_LEITOR` → Xrecord `OBRA`, mesmo JSON do arquivo (ADR-003) |
