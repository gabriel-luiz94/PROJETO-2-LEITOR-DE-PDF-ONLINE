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

## RQ — Correções do primeiro uso (TASK-004 a 010)

| ID | Requisito | Situação | Onde |
|---|---|---|---|
| RQ-60 | Diagnóstico do que foi lido (arquivo + tabela) e log, para desenhos reais virarem teste | ✅ · 🟡 CAD | `ORC_DIAGNOSTICO`, `Diagnostico`, `Aplicacao.Log` |
| RQ-61 | Leitura e filtros iguais aos do Leitor da aplicação no desenho inteiro; textos descartados visíveis (Ignorados) | ✅ | `MontadorDesenho`, `LeitorDxfTests`, filtro Ignorados |
| RQ-62 | Anotação ligada ao bloco/vão pelas linhas guias (LEADER, MLEADER, linha curta), medindo pelo contorno do texto | ✅ · 🟡 CAD | `LeitorProjeto.Ancoras` |
| RQ-63 | Layers da rede, bloco/layer de poste e guias configuráveis por projeto | ✅ · 🟡 CAD | `ConfigLeitura`, `TelaLeitura`, `ORC_LEITURA` |
| RQ-64 | Seleção manual funcionando (poste bloco/círculo, anotação TEXT/MTEXT/MLEADER) | 🟡 | `ORC_POSTE`, `ORC_VAO`, `Visual.Executar` |
| RQ-65 | Voltar ao menu inicial; fechar a obra | 🟡 | ⌂ Início, `ORC_FECHAR` |
| RQ-66 | Visual na paleta da aplicação | 🟡 | `Plugin/Tema.cs` |
| RQ-68 | Várias anotações no mesmo vão/poste (recondutagem, troca de poste) | ✅ · 🟡 CAD | irmãos V7b/P3b (`LeitorProjeto.CriarIrmaos`) |
| RQ-69 | Sugerir a configuração de leitura a partir do desenho (layers de rede, blocos de poste) | ✅ · 🟡 CAD | `Diagnostico.SugestoesDeRede/DeBlocos`, ORC_LER |
| RQ-70 | Painel com Cabos separados de Postes/estruturas (como o Resumo da aplicação), ícones fáceis de ver | ✅ · 🟡 CAD | `PainelObra`, `Relatorios.Cabos/Outros`, `Icones` |
| RQ-71 | Ler mostra só o que as regras do leitor classificaram; Construir é livre (sem filtro das regras) | ✅ | `Relatorios.Visivel` |
| RQ-72 | Atualização automática do plugin (baixa da aplicação, confere, vale ao reabrir; voltar versão) | ✅ · 🟡 CAD | `Atualizacao`, `AtualizacaoPlugin`, `/api/plugin/atualizacao` |
| RQ-73 | DLL pronta para o nanoCAD 5 (.NET 4), degradando o que a API dele não tiver | ✅ build · 🟡 CAD | `src/nanoCAD5`, `src/ApiNanoCAD5`, `Compat` |
| RQ-74 | Tela inicial com todos os projetos da aplicação; regras sempre do projeto escolhido | ✅ · 🟡 CAD | `/api/plugin/projetos`, `RegrasFrescas` |
| RQ-75 | Configuração de leitura guardada na aplicação por projeto (qualquer computador) | ✅ · 🟡 CAD | `/api/plugin/leitura`, `RegrasDe` |
| RQ-76 | Instalado junto com o programa principal | ✅ · 🟡 Windows | `services/plugin_instalador.py`, botão "Plugin nanoCAD" |
| RQ-77 | Editar o ativo no plugin com o autocompletar da Base de Orçamento (como na aplicação) | ✅ · 🟡 CAD | `CampoAtivo`, `Autocompletar`, `/api/plugin/catalogo` |
| RQ-78 | Excluir qualquer linha da obra (a lida pode ser restaurada) | ✅ · 🟡 CAD | `Construtor.Excluir/Restaurar` |
| RQ-79 | Adicionar linha sem apontar no desenho (Cabos / Outros) | ✅ · 🟡 CAD | `ORC_NOVO_CABO`, `ORC_NOVO_OUTROS` |
| RQ-67 | Abrir a aplicação já com a obra do plugin carregada | ✅ API · 🟡 CAD | `ORC_ABRIR_APP`, `/api/plugin/obra`, Resumo `?obra=` |

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
