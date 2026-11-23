# CHANGELOG DO PLUGIN

> Alterações relevantes do `plugin_nanocad/` (arquitetura, regras, contrato, comportamento). Mais recente primeiro.

---

## 2026-11-23 — 0.15.0

Anotação de poste longe do bloco e sem guia casa 1:1 com o bloco livre mais perto (até 25 × a altura, Conferir). A que continua sem bloco leva `_posteAprox` para o mapa do OrçaCore. DLLs do nanoCAD 26 e 5 geradas com `compilar.sh`. [TASK-037](tasks/TASK-037-23-11-2026.md).

## 2026-11-22 — 0.14.0

Montagem manual com polígono, limpeza e editor de anotações. Postes agrupam pelo layer, cor e operação. A área de revisão tem envio incremental à obra aberta no OrçaCore. [TASK-036](tasks/TASK-036-22-11-2026.md).

## 2026-11-20 — 0.13.1

Ler projeto, reler e diagnosticar usam o DXF ou DWG ativo. DWGs e desenhos com alterações em memória são exportados automaticamente para um DXF temporário e seguem para o mesmo Leitor do OrçaCore. [TASK-035](tasks/TASK-035-20-11-2026.md).

## 2026-11-20 — 0.13.0

A leitura automática usa o mesmo extrator DXF e o motor JavaScript do OrçaCore. O plugin preserva as linhas e sua ordem, associando a geometria e os vínculos depois da classificação. [TASK-034](tasks/TASK-034-20-11-2026.md).

## 2026-11-20 — 0.12.4

A tela inicial abre pelo comando `ORCACORE` ou pelo atalho `OC`. `ORCAMENTO` continua disponível. [TASK-033](tasks/TASK-033-20-11-2026.md).

## 2026-11-20 — 0.12.3

A seleção de projeto usa o mesmo código para carregar as regras e ler o desenho. Paraíba (027) usa o pacote salvo no OrçaCore, sua última cópia baixada ou o seed do leitor de 027. O log indica a origem e `ORC_DIAGNOSTICO` registra o código, as contagens e a procedência da geometria. [TASK-032](tasks/TASK-032-20-11-2026.md).

## 2026-11-20 — 0.12.2

O painel abre mesmo durante o preenchimento inicial, quando o WinForms ainda está sem aba selecionada. A correção trata o erro em `ORCAMENTO` ao entrar em Construir. [TASK-031](tasks/TASK-031-20-11-2026.md).

## 2026-11-19 — 0.12.0

Criação e edição de MTEXT, movimentação de grupo, reapontamento de guias associativas e auditoria das layers de operação ficam disponíveis no menu Anotações. O modo Análise usa marcas temporárias e filtra os MTEXT na viewport quando a API oferece o recurso.

## 2026-11-19 — 0.11.0

Painel de auditoria com busca e filtros, seleção de candidatos, destaque de origem e destino, cores temporárias por operação e prévia de lote. A saída Cabos e Outros continua usando a leitura e as regras do projeto.

## 2026-11-19 — 0.10.4

DLL para nanoCAD 5 Free compilada em x86, CLR 2 e .NET Framework 3.5. A versão 26 continua em .NET 6.

## 2026-11-19 — integração nanoCAD

TASK-025 e TASK-026: plugin 0.10.3, N3 removido na tabela Outros e teste de paridade incluindo chamadas múltiplas. A abertura no Windows depende também da correção no OrçaCore (TASK-197). Testar com a nota real nas duas versões do nanoCAD.

## 2026-11-19 — 0.10.2 para avaliação no CAD

- TASK-022: leitura das cores e do modo do projeto, curingas nas layers, cópia local e edição do modo no CAD.
- TASK-023: distribuição por poste e por prioridade de regra, reserva de vínculos manuais e comparação do exemplo com o motor da aplicação.
- TASK-024: menus Projeto, Conferir e Configurações, origem das regras no cabeçalho, ações de entrega visíveis e Atualizar como ação principal.

## 2026-11-15 — TASK-021: "+ Cabo" lê o texto como cabo e pega o comprimento da mesma linha ou da vizinhança (0.10.1)

`Construtor.LerCaboDeTexto`: em "+ Cabo" (e quando as regras leem cabo em "+ Outros"), o texto escolhido é SEMPRE lido como cabo — rótulo pelo que as
regras reconhecem como CABO ou pelo processamento — e o comprimento vem, nesta ordem: da mesma linha ("… 35 m"), de um texto de comprimento escolhido
junto, ou da VIZINHANÇA no desenho (`ComprimentoVizinho`: o texto "35 m"/"35,5M"/"40 metros" mais perto do rótulo, até 8 alturas, ao lado ou embaixo
antes de em cima), sem precisar selecioná-lo. Os textos do desenho só são lidos quando falta o comprimento. A tela mostra de onde veio o comprimento
ou avisa que não achou. "+ Vão" (`ORC_VAO`) usa a mesma leitura. Testes: `CaboDeTextoTests` (13).

---

## 2026-11-13 — TASK-020: integração com o projeto de trabalho do OrçaCore; sem gravação no banco (0.10.0)

- **Salvar** (`ORC_SALVAR`) grava `<desenho>.obra` na pasta do desenho no formato do 📁 Projeto do OrçaCore (`Nucleo/ProjetoArquivo`): envelope +
  `cad` + `projeto_trabalho` (SAP, resultado e o que o programa guardou ficam; nome e número de obra vêm da obra). `cad.assinaturaLinhas` = impressão
  das linhas gravadas; se o arquivo tiver outras linhas (editadas no OrçaCore) ou for um projeto só do programa, pergunta antes de gravar por cima.
  Gravação atômica. A cópia no DWG continua. O `.obra.json` antigo ainda é lido no Continuar (o `.obra` vale primeiro).
- **Nome e número de obra** (`ORC_PROJETO`, `TelaProjeto`, botão no painel; perguntado no primeiro Salvar): `Obra.NumeroObra`; o Continuar lê o que
  o programa guardou em `projeto_trabalho`.
- **🌐 Abrir na aplicação** (`ORC_ABRIR_APP`): salva e entrega o CAMINHO ao OrçaCore deste computador (`ProgramaLocal.EntregarProjeto`; fechado:
  `--abrir-projeto "<caminho>"`). **Removido** o envio para o banco (`ClienteAplicacao.EnviarObra`, `POST /api/plugin/obra`). Sem o programa: o
  projeto fica na pasta do desenho.
- Nome: **OrçaCore para nanoCAD** (o programa mudou de nome na TASK-156 do principal).
- Testes: `ProjetoArquivoTests` (6), `ProgramaLocalTests` (+1); exemplo `exemplos/Projeto_Rua_das_Flores.obra` aberto pela aplicação
  (`tests/test_plugin_projeto_163.py`).

---

## 2026-11-12 — TASK-019: cabo lido de texto escolhido — comprimento da linha de baixo na linha do cabo, sem vírgula (0.9.2)

Quando o texto escolhido (+ Cabo / + Outros e + Vão) está numa layer que não é de operação, as regras não CLASSIFICAM, mas o PROCESSAMENTO de
Pernambuco já monta "3XA35P(7,94)-13,8KV 35 m" (regra "Cabo: label sozinho, comprimento na linha de baixo", janela de 2 linhas). O plugin
juntava os textos crus com ", " ("3XA35P…, 35 M"). Agora usa o resultado do processamento (`Construtor.AtivoDosTextos`): o comprimento entra na
linha do cabo e a linha só de comprimento não entra; textos que o processamento não juntou seguem separados por ", ". Sem classificação, um cabo
com comprimento vai para a Tabela de Cabos. Testes em `LinhaDeTextoTests` (layer de operação, layer 0, mesmo MTEXT, ANOTACOES, + Vão).

---

## 2026-11-12 — TASK-018: "+ Cabo / + Outros" lendo um TEXT/MTEXT escolhido no desenho (0.9.1)

Na tela de linha nova, "Selecionar texto no desenho (TEXT/MTEXT)…": a tela sai da frente, o usuário escolhe o(s) texto(s) e as regras do leitor
leem (`Construtor.LerTextoParaLinha`): ativo, operação e a tabela (CABO → Cabos; outra entidade → Outros). A tela volta preenchida para conferir
(autocompletar continua valendo) e escolher Cabos/Outros. A linha fica ligada ao texto (`LigarTexto`: localizar, marcas; ativo mudado antes de
adicionar = correção sobre o lido). ↻ Atualizar: a linha volta, a não ser que a leitura passe a pegar aquele texto (não fica em dobro).
Testes: `LinhaDeTextoTests` (4).

---

## 2026-11-12 — TASK-017: editar o ativo com autocompletar, excluir qualquer linha, linhas digitadas (0.9.0)

- **Editar o ativo** no pop-up (`ORC_EDITAR` / duplo clique): `CampoAtivo` com o autocompletar da aplicação (`Nucleo/Autocompletar`: sugestões da
  Base de Orçamento do projeto com descrição e apelidos, dica dos parâmetros dentro de "(", valores do parâmetro atual, `)` automático — porte de
  `static/ativo_dica.js`). Catálogo de `/api/plugin/catalogo` (`Plugin/Catalogos`, 10 min na memória + cópia em `cache/`). `Construtor.TrocarAtivo`:
  guarda o do desenho (`AtivoDoDesenho`), o vão relê cabo/metros; "Voltar ao lido" desfaz. O desenho não é alterado.
- **Excluir qualquer linha** (`Construtor.Excluir`): a lida fica `Excluido` (fora da obra, do vínculo, do .obra.json, das marcas; listada em Ignorados,
  duplo clique → "Restaurar à obra"); a montada à mão é apagada.
- **+ Cabo / + Outros** (`ORC_NOVO_CABO`/`ORC_NOVO_OUTROS`, `TelaNovaLinha`): linha digitada (L1, L2…) para a Tabela de Cabos ou de Outros.
- Correção, exclusão e linhas digitadas sobrevivem ao ↻ Atualizar (`ReaplicarAjustes`; a correção vale enquanto o texto do desenho for o mesmo) e ao
  Salvar/Continuar (seção `cad`: `ativoDoDesenho`, `excluido`).
- Testes: `EdicaoTests` (6), `AutocompletarTests` (3).

---

## 2026-11-12 — TASK-016: "Abrir na aplicação" no GridReader deste computador (0.8.0)

`ORC_ABRIR_APP` / 🌐: procura o GridReader aberto (porta de `programa.json`, 8000 e a do ORC_CONECTAR local; `GET /api/plugin/desktop`) e entrega a
obra (`POST /api/plugin/receber`, sem login: só do próprio computador) — a janela dele vem para a frente com a obra. Fechado: grava a obra em
`%APPDATA%\OrcamentoNanoCAD\obras_para_abrir\` e abre o .exe com `--abrir-obra "<arquivo>"`. Sem o programa instalado: o caminho antigo
(aplicação do ORC_CONECTAR no navegador). `Nucleo/ProgramaLocal`, `Plugin/ProgramaNoComputador`; testes `ProgramaLocalTests`.

---

## 2026-11-11 — TASK-014 e TASK-015: nanoCAD 5 pronto; projetos e leitura da aplicação (0.7.0)

**TASK-014 (nanoCAD 5):** `src/nanoCAD5/OrcamentoNanoCAD5.csproj` (.NET 4) compila o MESMO código contra `src/ApiNanoCAD5` (hostmgd/hostdbmgd
gerados por `ferramentas/GerarApiNanoCAD5` com o Roslyn: só o que o plugin usa da API). `Plugin/Compat.cs`: recurso que a API do nanoCAD não
tiver (MLEADER, transitórios, UserData, vista, janela não modal, comando em andamento, blocos dinâmicos, contorno) é desligado sem derrubar o
resto; extrator com um leitor por tipo de entidade; marcas com reserva em vetores; zoom pelo comando ZOOM. Versão em `src/Versao.props`.

**TASK-015 (aplicação):** tela inicial lista todos os projetos da aplicação (`/api/plugin/projetos`, cache); regras sempre baixadas do projeto
escolhido/digitado (sem a aplicação: última cópia ou as embutidas); configuração de leitura guardada na aplicação (`/api/plugin/leitura`) e
vinda em `/api/plugin/regras` (vale em qualquer computador; a local fica como reserva). A atualização automática grava `versao.json` ao lado da DLL.

---

## 2026-11-10 — TASK-013: atualização automática (0.6.0)

**Tipo:** funcionalidade · `Nucleo/Atualizacao`, `Plugin/AtualizacaoPlugin`, `ClienteAplicacao` (AtualizarPlugin), `TelaConexao`, `compilar.sh` (versao.json)

20 s depois de abrir o nanoCAD e a cada 4 h, o plugin pergunta à aplicação (`/api/plugin/atualizacao`, TASK-130 do principal) se há versão
mais nova; baixa só de HTTPS ou localhost, confere tamanho/SHA-256/"MZ" e troca a DLL ao lado da atual (a em uso vira `.anterior`).
Vale ao reabrir o nanoCAD. Arquivo travado → PowerShell oculto troca quando o nanoCAD fechar. `ORC_ATUALIZAR_PLUGIN`, `ORC_VOLTAR_VERSAO`
(a versão desfeita não volta a ser baixada). Liga/desliga em `ORC_CONECTAR`.

---

## 2026-11-03 — TASK-012: painel Cabos / Postes e estruturas, filtro do leitor, ícones (0.5.0)

**Tipo:** melhoria · `PainelObra`, `Relatorios`, `Marcas`, `Tema`, `Icones` (novo)

Painel em duas listas (CABOS e POSTES E ESTRUTURAS), as mesmas linhas de Cabos/Outros da aplicação. Na leitura só aparece o que
as regras do leitor classificaram (sem anotação = oculto no painel, nas marcas e na tabela de vínculos; a contagem vai no rodapé);
em Construir e no que foi apontado à mão, tudo aparece. Ícones desenhados: estado, rede MT/BT, operação (+ I, − R, ↻ M), poste/solto;
barra de ícones com dica; filtros com ícone.

---

## 2026-11-02 — TASK-011: leitura do desenho real (0.4.0)

**Tipo:** correção · `LeitorProjeto`, `ConfigLeitura`, `Diagnostico`, `Vinculador`, `Marcas`, `Telas`

Com o primeiro diagnóstico real: guias na mesma layer do texto até 6 alturas e em pedaços; várias anotações por vão/poste
(irmãos V7b/P3b); MT/BT da anotação respeitado; redes em pedaços juntadas; blocos no mesmo lugar = um poste; rede sem poste
descartada; 2ª passada por proximidade (Conferir); placa do poste; sugestões de layers de rede e blocos de poste; "=NOME" exato.

---

## 2026-11-02 — TASK-004 a 010: correções do primeiro uso no nanoCAD (0.3.0)

**Tipo:** correção + feature · núcleo (`DesenhoBruto`, `ConfigLeitura`, `Diagnostico`, `LeitorProjeto`), CAD (`ExtratorDesenho`, `Tema`,
`PainelObra`, `Telas`, `Comandos`) · aplicação: `routers/plugin_cad.py`, `static/resumo.js`, `login.js`, `auth_fetch.js` (TASK-106 do principal)

- Leitura em duas etapas: DWG → desenho **bruto** (CAD) → interpretação no núcleo (testável com DXF; paridade texto a texto com o Leitor da aplicação).
- **MLEADER** lido (texto + linhas como guias); LEADER sempre guia; guia toca o **contorno** do texto.
- **Configurar leitura** por projeto (layers da rede, bloco/layer de poste, círculo como poste, layers de guia, raios).
- **Ignorados** no painel (textos que as regras descartaram, com motivo); `ORC_DIAGNOSTICO` (arquivo + tabela) e log.
- Botões do painel disparam com Enter e foco no desenho; seleção manual aceita bloco/círculo, MLEADER, polilinha 3D.
- Navegação (⌂ Início, ✕ Fechar obra, voltar ao painel); visual na paleta da aplicação.
- **🌐 Abrir na aplicação** (`ORC_ABRIR_APP`): grava a obra do desenho na aplicação e abre o Resumo com ela carregada.

---

## 2026-10-31 — TASK-003: ligação com a aplicação (0.2.0)

**Tipo:** feature · `src/Nucleo/ClienteAplicacao.cs`, `src/Plugin/Sincronizador.cs`, telas/comandos · aplicação: `routers/plugin_cad.py` (TASK-104 do principal)

Com a aplicação aberta (localhost/online), o plugin usa as regras gravadas nela e confere a obra com a Base de Orçamento;
mudanças lá aparecem em poucos segundos. Estrutura/cabo fora da base vira Conferir. Comandos `ORC_CONECTAR` e `ORC_SINCRONIZAR`.

---

## 2026-10-31 — TASK-001: subproduto plugin nanoCAD (0.1.0)

**Tipo:** novo subproduto · `plugin_nanocad/` · aplicação: `services/obras_arquivo.py` (TASK-103 do principal)

- Pasta própria com `.ai` (requisitos, estado, arquitetura, regras, glossário, decisões, tarefas), mockups em `referencias/`.
- Núcleo C# sem CAD: motor do leitor portado (paridade com o JS), anotações (paridade com o Python), Ler projeto
  (redes → vãos → associação → numeração seguindo a rede), vínculo automático/manual com parâmetros iguais aos da
  aplicação, conferência, Construir do zero, ajustes, `.obra.json`.
- Camada nanoCAD 26: 18 comandos (`ORCAMENTO`, `ORC_*`), telas (início, painel, pop-up, tabelas), marcas transitórias,
  obra gravada no DWG e em `<desenho>.obra.json`, troca opcional da layer da anotação.
- DLL pronta em `appload/nanoCAD26/`; `build_nanocad5.bat` para o nanoCAD 5.
- Contrato: a importação de obra da aplicação passa a manter os campos de vínculo/geometria; teste de contrato em `tests/test_plugin_nanocad_contrato.py`.
