# CHANGELOG

> Histórico de alterações **relevantes** do projeto, do ponto de vista de contexto para IA:
> arquitetura, regras de negócio, fluxo de dados, estrutura de banco e comportamento importante.
>
> Não registre aqui alterações triviais (formatação, renomeação local, ajuste de texto).
> Formato: mais recente primeiro.

---

## 2026-11-26 — Removendo: verde / laranja / vermelho

Substitui o azul da TASK-252. Linhas R e *R no «Colorir ativos por parâmetros»: **verde** = parâmetros ok; **laranja** = falta (ou é inválido) um parâmetro que decide material ou mão de obra (o parâmetro muda a RETIRADA: há linhas da base restritas por ele com fator R ≠ 0 ou fórmula R e ≥ 2 valores/máscaras entre elas; parâmetro que só escolhe itens de instalação ou só entra na fórmula de quantidade não conta — ex.: `N3` removendo fica verde, `AVULSO_TRAFO(trafo=…)` laranja se o trafo não é válido); **vermelho** = ativo não cadastrado na base. Sem aviso/validação novos (TASK-247 mantida). [TASK-253](tasks/TASK-253-26-11-2026.md).

## 2026-11-26 — Removendo (R, *R): destaque de parâmetro determinante e de ativo fora da base

No «Colorir ativos por parâmetros», as linhas removendo voltam a ser destacadas, mas só com duas marcas: **vermelho** (o ativo não existe na base — Cabos e Outros) e **azul** (`par-determinante`: o ativo removido tem parâmetro que decide QUAL item sai). Sem verde/laranja nelas (parâmetro faltando continua sem aviso em R/*R, TASK-247). [TASK-252](tasks/TASK-252-26-11-2026.md).

## 2026-11-26 — Cabos: notação com seção AWG (`3xA1/0(1/0)+1/0`)

`services/anotacao_cabos.py` e `static/anotacao_cabos.js`: o `+N` do condutor extra aceita seção AWG fracionária (`+1/0`, `+4/0`) e o neutro/extra AWG (`(1/0)`, `+1/0`) mapeia para a mesma família do cabo (`A1/0` → `A1/0`, `A4/0` → `A4/0`); em cabo que não é AWG, o valor segue literal. Sem regressão em `3xC16(16)+16`, `3xA35P(7,94)` e `+20kV`. [TASK-251](tasks/TASK-251-26-11-2026.md).

## 2026-11-26 — Montar Orçamento rápido e avisos sem vinculação

O Montar Orçamento demorava ~2–3 s mesmo com uma linha: cada `processar_calculo` refazia a preparação da base inteira (contexto, parâmetros, compostos; ≈1 s em ~4 mil linhas) e o clique calculava 3 vezes (validação de cabos e de outros + Resultado). Agora `contexto_da_base` e a preparação ficam em cache pela impressão digital do conteúdo da base (+ projeto + convenção de postes) e o servidor prepara cada projeto ao iniciar: clique → Resultado ≈0,2 s. Os avisos do Montar Orçamento ficam só com erro, ativo não encontrado (C1-BASE) e parâmetros (C1-PARAM); os achados de vinculação saíram da validação. [TASK-250](tasks/TASK-250-26-11-2026.md).

## 2026-11-25 — Plugin 0.16.6: Abrir no OrçaCore sem travar o nanoCAD

A conversa do plugin com o OrçaCore deste computador (Abrir no OrçaCore) saiu da thread do nanoCAD; chamadas a 127.0.0.1 sem proxy e sem keep-alive; tempos no log do plugin e passos da janela no log do OrçaCore (`services/janela_desktop.py`). [TASK-039 do plugin](../plugin_nanocad/.ai/tasks/TASK-039-25-11-2026.md).

## 2026-11-25 — Plugin 0.16.5: nanoCAD 5 sem o «API Logger» no Salvar

Os botões do painel do plugin consultavam `Document.CommandInProgress`, que o nanoCAD 5 declara e não implementa (abria o «API Logger» e o Salvar da montagem manual não gravava). No build do nanoCAD 5 a consulta saiu. [TASK-039 do plugin](../plugin_nanocad/.ai/tasks/TASK-039-25-11-2026.md).

## 2026-11-24 — Mapa: parâmetros dos ativos com Ctrl+Enter

As linhas do ponto no Mapa (editar linha, + Linha, Editar ponto/estrutura, Nova estrutura/poste) abrem o mesmo balão de parâmetros das tabelas principais: Ctrl+Enter, duplo clique ou mouse parado; ↑↓/Tab/Enter/Del/Esc como na tabela. `AtivoBalao` ganhou `seletorExtra`/`preparar` (pontas por campo em `input._ativoBalao`) e desenha o balão dentro do `<dialog>` modal. [TASK-249](tasks/TASK-249-24-11-2026.md).

## 2026-11-24 — Cava em rocha (modal ao lado do MAPA)

Nova parte `cava_rocha` do pacote (pacote 081 traz as 121 localidades da aba CAVA_ROCHA da ICP). Botão **⛏ CAVA EM ROCHA** na barra de modais, logo depois do MAPA, só no projeto que tem a parte: calcula SDEMU3007II = mín(F; parte inteira de MARRED(F × índice; 0,1)), com F = estais em solo + postes de 9 m ou mais instalados no orçamento do Resumo (+ extra digitado), e lança `<qtd>-CAVA_ROCHA` em Outros (substitui a anterior; Ctrl+Z desfaz). `services/cava_rocha.py` e `static/cava_rocha.js` com a mesma conta; rota `GET /api/orcamento/cava_rocha` (com `ativo_na_base`). [TASK-248](tasks/TASK-248-24-11-2026.md).

## 2026-11-24 — Resultado: removendo sem aviso de parâmetro faltando

O aviso «faltou informar CABO … — calculado só o que não depende disso» do Resultado do orçamento não aparece mais para itens R e *R (o cálculo não muda). Complemento da [TASK-247](tasks/TASK-247-24-11-2026.md).

## 2026-11-24 — Removendo e fictícia fora dos avisos; cor dos parâmetros na hora

Avisos de vinculação só para I, *I, M, *M (R, *R e F não avisam); C1-PARAM não avisa linhas R/*R (o C1-BASE continua); «Colorir ativos por parâmetros» não pinta removendo e atualiza na hora ao gravar pelo balão ou trocar a operação; cabo `CAZ7,94` não fica mais vermelho quando a base tem o literal `[CAZ7,94]`. [TASK-247](tasks/TASK-247-24-11-2026.md).

## 2026-11-24 — Plugin 0.16.4: cabo pelo trecho clicado e metragem por linha

Modo mapa do plugin, Selecionar cabo: sem a pergunta «Inteira/Trecho» — polilinha vale o trecho clicado (vértice a vértice), LINE vale inteira. Janela com Medido na linha e Comprimento usado (padrão: medido arredondado) e metragem própria em cada linha do cabo («PB70 32 m», «PB70/32 m»). [TASK-246](tasks/TASK-246-24-11-2026.md).

## 2026-11-24 — Transporte de material (TRANSP / TRANSP_KM)

Nova parte `transporte` do pacote de projeto (o pacote 081 a traz). Com ela, o menu **Mais** do Resumo ganha «Transporte de material…»: calcula o frete dos materiais pesados igual à aba TRANSPORT da ICP (151 casos conferidos, Python e JS) a partir da instalação do orçamento do Resumo e da UTD da obra, e lança `TRANSP(faixa=…)`/`TRANSP_KM(faixa=…)` em Outros. [TASK-245](tasks/TASK-245-24-11-2026.md).

## 2026-11-24 — Estrutura com quantidade no Mapa

«2 FLABIT» passa a ser duas estruturas no vínculo/Mapa (a 2ª entra no fim da linha, sem mexer nos vínculos existentes) e aparece como «FLABIT · 2ª» no painel do ponto. [TASK-244](tasks/TASK-244-24-11-2026.md).

## 2026-11-24 — Pré-preenchimento do ativo do cabo

Nos formulários de cabo do Mapa, o ativo se completa enquanto se digita: «3» → «3x», «A35P(» → «7,94)-13,8kV», «-» → «13,8kV» (aprendendo com os cabos da obra); Tab aceita. [TASK-243](tasks/TASK-243-24-11-2026.md).

## 2026-11-24 — Editar cabo e poste no Mapa pelos formulários novos

«Editar» do cabo abre o formulário do «+ Cabo» (F/I/M/R, comprimento separado) e «Editar» da estrutura abre a linha em «Linhas do ponto»; o cartão antigo saiu. [TASK-242](tasks/TASK-242-24-11-2026.md).

## 2026-11-24 — Mapa como janela e em outra janela do Windows

«Vincular cabos» passa a ser o **Mapa** (abre na aba Mapa) e abre como as janelas de Resultado e Base: área principal, rodapé, maximizar/restaurar e «Abrir em nova janela» — o Resumo em modo janela mostra só o Mapa numa janela do Windows, sincronizado nos dois sentidos com a principal. Esc no Mapa só deseleciona. [TASK-241](tasks/TASK-241-24-11-2026.md).

## 2026-11-24 — Plugin 0.16.3: cor do cabo, coletados na tela e ações do item

Modo mapa do plugin: tipo do cabo pela cor da linha; pontos e cabos coletados marcados na tela enquanto o painel está aberto; no painel, clicar num item enquadra e destaca o item com o objeto base, e há Editar, Excluir e Definir poste/cabo. [TASK-038 do plugin](../plugin_nanocad/.ai/tasks/TASK-038-23-11-2026.md).

## 2026-11-24 — Cabos lado a lado no mapa

Cabos entre o mesmo par de pontos ficam lado a lado (juntos pelo par, não pela geometria), cada um do seu lado qualquer que seja o sentido, com no mínimo 9 px entre eles em qualquer zoom. [TASK-240](tasks/TASK-240-24-11-2026.md).

## 2026-11-24 — Plugin 0.16.2: anotação de várias linhas sem ativo dividido

Selecionar poste/cabo no plugin: o parágrafo de continuação do mesmo MTEXT entra no campo de cima e a metragem já juntada pela regra não vira campo F separado (que, trocado para I, dividia o ativo). [TASK-038 do plugin](../plugin_nanocad/.ai/tasks/TASK-038-23-11-2026.md).

## 2026-11-24 — Cabos sem geometria no DXF (moldura, tipo, X) e nome do ponto no cartão

Leitura do DXF: a anotação emoldurada procura a linha-guia pela borda da moldura; o cabo que diz a tensão só vai para a rede do seu tipo; marcas de X não são rede. Pacote 081 com ANOTACOES nas layers de MT. O cartão do cabo usa o nome do ponto do mapa. Na NOTA_9101578729: 87 de 94 cabos com geometria (antes 50), nenhum no tipo errado. [TASK-239](tasks/TASK-239-24-11-2026.md).

## 2026-11-24 — Plugin 0.16.1: pontos e cabos do mapa no painel

No plugin, os pontos e cabos de Selecionar poste/cabo aparecem nas abas do painel e podem ser ajustados ou excluídos (`ORC_MAPA_ITEM`). [TASK-038 do plugin](../plugin_nanocad/.ai/tasks/TASK-038-23-11-2026.md).

## 2026-11-24 — Um ponto só por grupo no mapa

O grupo «@nome» é um ponto só no mapa, com todas as operações: as linhas sem coordenada própria entram no ponto do grupo (antes, cada uma virava um «P1»). O painel do ponto lista todas as estruturas, com I/\*I/M/\*M em destaque e R/\*R apagadas, fora do vínculo. [TASK-238](tasks/TASK-238-24-11-2026.md).

## 2026-11-23 — Plugin: montagem manual no formato do mapa

O modo manual do nanoCAD (plugin 0.16.0) entrega cada poste como o ponto «@nome - comentário» com uma linha por campo e a sua operação (F, I, \*I, M, \*M, R, \*R), e cada cabo (linha inteira ou trecho) com a metragem e a geometria, sem vínculo. Com as regras de um projeto, o texto fora delas entra como F. Chega pelo envio incremental de anotações; o vínculo é feito no mapa. [TASK-237](tasks/TASK-237-23-11-2026.md).

## 2026-11-23 — PDF com coordenadas no mapa

O PDF vetorial monta o mapa como o DXF: rede pela cor do traço, postes nas pontas da rede, associação texto→rede/poste igual à do DXF e escala escolhida entre 1:500 e 1:1000 pelas metragens dos cabos (`services/pdf_geometria.py`). Coordenadas do mapa em metros; `_x/_y` seguem em pontos da página. [TASK-236](tasks/TASK-236-23-11-2026.md).

## 2026-11-23 — Novo conceito do ponto no mapa

O mapa monta a rede com + Ponto, + Cabo e Traçar cabo. O ponto é «@nome - comentário» com linhas e operação por linha; F (fictício) congela a linha fora do orçamento mas mantém a estrutura no vínculo, e o poste da linha F preenche o parâmetro POSTE do ponto. Edição da linha no painel mantém os vínculos quando a nova estrutura aceita os cabos; clique no cabo e na estrutura vincula. Pontos no formato antigo são convertidos sob confirmação (linhas «#» juntas por «; »). [TASK-235](tasks/TASK-235-23-11-2026.md).

## 2026-11-23 — 2ª ponta do cabo M só à mão

O vínculo automático liga só as `pontas` do cabo (M = 1); o usuário pode ligar à mão até `pontasMax` (M = 2, coluna «Máx. à mão» da Tabela de Cabos). No painel do poste a ponta extra aparece como «Opcional». [TASK-234](tasks/TASK-234-23-11-2026.md).

## 2026-11-23 — Montagem rápida da rede no mapa

Ponto sem bloco no CAD aparece na posição da anotação (`_posteAprox` ou `_x/_y`), marcado como aproximado. O mapa mostra ao fundo a rede e os blocos do CAD que ainda não viraram linha (`services/fundo_cad.py`) e transforma um bloco em ponto com um clique. Painel de pendências com progresso e correção a um clique, ímã nas pontas do cabo, ferramenta «Traçar vão» com o comprimento medido, seleção por caixa, mover em grupo, atalhos e lista em lotes. O plugin casa 1:1 a anotação de poste que sobrou com o bloco livre mais perto (fonte, sem build). [TASK-233](tasks/TASK-233-23-11-2026.md).

## 2026-11-22 — Edição pelos cartões do mapa

Os cartões editam cabos, postes e estruturas nas próprias tabelas da obra. No cabo, a edição inclui a anotação com metragem, a operação, o número de fases, a tensão e o nível. [TASK-223](tasks/TASK-223-22-11-2026.md).

## 2026-11-22 — Arraste, seleção e âncoras no mapa

O mapa permite arrastar pontos georreferenciados, filtrar a rede, localizar as pontas pelo cabo selecionado e ancorá-las em postes fictícios. As âncoras são referências visuais excluídas do orçamento. [TASK-222](tasks/TASK-222-22-11-2026.md).

## 2026-11-22 — Conexões de cabo no mapa

A edição do traçado e a revisão de conexões detectadas registram as pontas da rede. O painel de cada ponto aplica as regras nas ligações com as estruturas. Os campos do mapa sugerem ativos da base do projeto. [TASK-221](tasks/TASK-221-22-11-2026.md).

## 2026-11-22 — Montagem da obra no mapa

O Resumo permite criar postes, estruturas, cabos reais e cabos fictícios pelo mapa. Os pontos avulsos podem receber pontas e ser associados a postes. O mapa explica bloqueios de vinculação com as regras do projeto. [TASK-220](tasks/TASK-220-22-11-2026.md).

## 2026-11-22 — Mapa editável da rede

Postes sem coordenadas começam em grade, com as pontas livres dos cabos na lateral. O mapa e o detalhe do poste permitem editar os vínculos e aplicar sugestões revisadas. As posições visuais são salvas com a obra. [TASK-219](tasks/TASK-219-22-11-2026.md).

## 2026-11-22 — Montagem manual do nanoCAD

O Leitor DXF aceita região no fluxo manual do plugin. O Resumo acrescenta as anotações revisadas à obra aberta e ajusta os vínculos do lote aos índices das tabelas de destino. [TASK-218](tasks/TASK-218-22-11-2026.md).

## 2026-11-21 — Segurança, prazo e DeepSeek no Assistente

O chat protege a chave OpenAI compartilhada e valida os destinos alternativos. O DeepSeek passa a usar o modelo `deepseek-flash`, com teste do ciclo completo de ferramenta. O prazo esgotado gera erro e encerramento do chat. O servidor exige uma chave JWT própria, e o desktop gera a sua no perfil local. [TASK-217](tasks/TASK-217-21-11-2026.md).

## 2026-11-21 — Arraste nativo e abertura mais leve

O Windows controla o movimento das janelas destacadas pela barra de título nativa. O botão de reencaixe continua na visualização. PDF, DXF e Excel carregam suas bibliotecas ao primeiro uso. [TASK-216](tasks/TASK-216-21-11-2026.md).

## 2026-11-21 — Resultado e Base em janelas nativas

Resultado e Base voltam a abrir em tela inteira no Resumo. A barra tem controles por ícones, e minimizar mantém a tela em uma aba no rodapé. No desktop, é possível destacar a visualização, movê-la para outro monitor e retornar à janela principal. Montar Orçamento atualiza o Resultado onde ele estiver aberto. [TASK-215](tasks/TASK-215-21-11-2026.md).

## 2026-11-21 — Largura das grades em obras grandes

As linhas pendentes de renderização e os separadores de grupo acompanham as colunas visíveis de Cabos e Demais Ativos. A tabela preenche a área disponível após restaurar a obra e após alternar as colunas opcionais. [TASK-214](tasks/TASK-214-21-11-2026.md).

## 2026-11-20 — Janelas auxiliares no Resumo

O Resumo revela Cabos e Outros ao receber a obra do nanoCAD e adapta as tabelas à área disponível. **Ir à linha** recolhe a comparação e mantém seus filtros. Vinculação, Base de Orçamento e Resultado podem ficar abertas como janelas flutuantes. **Montar Orçamento** atualiza o Resultado na mesma janela. [TASK-213](tasks/TASK-213-20-11-2026.md).

## 2026-11-20 — Motor do Leitor DXF no plugin nanoCAD

A rota autenticada do plugin usa o extrator e o motor JavaScript existentes do Leitor. O plugin preserva a classificação, a operação, o ativo e a ordem das linhas. [TASK-212](tasks/TASK-212-20-11-2026.md).

## 2026-11-20 — Comparação de obras no Resumo

Ler Projeto, Carregar Obra e Importar JSON oferecem Comparar. A janela abre nas divergências de Cabos ou de Postes e estruturas, indica correspondências incertas, mostra a diferença de metragem e permite localizar a linha da obra aberta. Ler Projeto também oferece Subtrair quando já há conteúdo. [TASK-211](tasks/TASK-211-20-11-2026.md).

## 2026-11-20 — Modal de vinculação e comando do nanoCAD

A análise de equipamentos que leem cabos do poste usa o índice espacial. O modal mantém o aviso e os controles de saída se uma aba falhar. No nanoCAD, `ORCACORE` e `OC` abrem a tela inicial. [TASK-210](tasks/TASK-210-20-11-2026.md) · [TASK-033 do plugin](../plugin_nanocad/.ai/tasks/TASK-033-20-11-2026.md).

## 2026-11-20 — Abertura de projetos grandes do nanoCAD

O Resumo carrega a obra entregue pelo nanoCAD pela API local autenticada. A vinculação automática e o mapa indexam os postes por região para localizar as pontas dos vãos. [TASK-209](tasks/TASK-209-20-11-2026.md).

## 2026-11-20 — Instalador Windows por usuário

A distribuição passa a gerar `OrcaCore-Setup.exe` com o mesmo binário e versão do portátil. O instalador cria o atalho e o registro de desinstalação para a conta atual. Após uma autoatualização, a versão exibida pelo Windows acompanha o executável que abriu. [TASK-206](tasks/TASK-206-20-11-2026.md).

## 2026-11-20 — Assistente de IA no Resumo e no mapa

Planos preservam a operação existente, respeitam linhas protegidas e remapeiam vínculos quando uma linha é removida. A prévia usa a Totalizadora do Resumo e compara quantidades de orçamento antes da aplicação. Os atalhos passam pelo mesmo cartão de confirmação. O mapa permite consultar o poste ou vão selecionado. Skills publicadas têm revisão por código de projeto e o teste do provedor reserva verifica uma ferramenta. [TASK-204](tasks/TASK-204-20-11-2026.md) · [TASK-205](tasks/TASK-205-20-11-2026.md).

## 2026-11-20 — Grafo de vinculação

O mapa mostra postes como nós e vãos como arestas. A janela do poste lista os cabos e as estruturas e permite aplicar ou desfazer vínculos pelas regras do projeto.

## 2026-11-19 — Layers das notas de Pernambuco

O pacote 081 reconhece as layers `POSTE_*` de instalação, remoção e reaplicação e suas variantes de linha viva. Seus dados alimentam o leitor da aplicação e o plugin.

## 2026-11-19 — Auditoria da leitura no nanoCAD

O plugin mostra a origem da anotação, a metragem lida, a geometria e a saída da obra. A configuração por projeto aceita o campo opcional `auditoria_tolerancia_m` de 0 a 1000. Zero desativa o aviso. A revisão e as escolhas de destino são armazenadas na seção `cad` da obra.

## 2026-11-19 — Abrir no OrçaCore

A entrega ao Resumo aberto usa a tela atual. A leitura de cabeçalho SAP vazio dispensa o acesso ao perfil do projeto. O instalador verifica o conteúdo da DLL ao encontrar a mesma versão.

## 2026-11-19 — integração nanoCAD

TASK-197 e TASK-198: entrega do .obra à janela desktop, regras do 081 para N3 removido e extração de chamadas MULTILEADER no DXF. Testar com a nota real nas duas versões do nanoCAD.

## 2026-11-18 — Ocultar parâmetros: com a linha em foco continua só o nome; "(" abre o balão

- Com «Ocultar parâmetros ao sair da linha» ligado, clicar na linha **não** mostra os parâmetros na linha: continua `PR`. Ao **digitar `(`** (ou levar o cursor para depois dele: End, →, clique) os parâmetros **preenchidos** aparecem no **balão** (o mesmo do mouse: `cabo = XLPE · manual`…), ancorado no campo; some ao sair do campo ou ao voltar ao nome. API: `window.mostrarBalaoParametros(el)`/`esconderBalaoParametros(el)` (`resumo.js`); o balão de outra linha nunca fica na tela. O texto do campo e os cálculos não mudam.

## 2026-11-18 — Correção: balão de parâmetros e cores voltam a funcionar

- A declaração de `_cacheAssin` (resumo.js) tinha sido apagada numa edição da TASK-185; o `ReferenceError` era engolido por `try/catch` e o balão «N1 · completo (3 de 3) … manual» e a coloração por parâmetros ficavam vazios. Declaração restaurada no topo do arquivo e teste de regressão (`test_cache_de_assinaturas_declarado_no_topo_do_resumo`).

## 2026-11-23 — TASK-232: instalador privado com dados locais

- A compilação privada empacota o executável, o `.env` do destinatário e uma cópia SQLite sem sessões. O Setup instala os arquivos ao lado do executável e preserva os que já existem. A abertura é local, com acesso direto ao Supabase pelo fluxo desktop. [TASK-232](tasks/TASK-232-23-11-2026.md).

## 2026-11-23 — TASK-231: revisão dos vínculos no ponto e edição do poste

- O painel mostra cabos e estruturas em duas colunas, desenha os vínculos, destaca os cabos de uma estrutura selecionada e oferece **Só pendências**.
- **Editar poste** reutiliza o cadastro preenchido e mantém a posição e os cabos associados. A montagem aplica as regras do projeto para preencher vínculos inequívocos nas chegadas confirmadas. As disputas seguem para revisão. [TASK-231](tasks/TASK-231-23-11-2026.md).

## 2026-11-23 — TASK-230: comentário opcional nos pontos do mapa

- Criar, agrupar ou copiar um ponto sem texto de comentário gera o título e as linhas de ativos. Ao limpar o comentário e salvar no painel, a linha é removida e os vínculos acompanham os índices das linhas. [TASK-230](tasks/TASK-230-23-11-2026.md).

## 2026-11-23 — TASK-230: script do pacote de entrega (instalador + dados prontos para outra pessoa)

- `scripts/montar_pacote_entrega.py`: copia o banco com a API de backup (consistente com o programa aberto/WAL), **higieniza** (apaga usuários, sessões, auditoria, sync, chaves salvas em `configuracoes`, obras/histórico REC e históricos de regras; `VACUUM` para apagar de verdade; sem `-wal/-shm`), cria a conta de administrador do usuário e torna a base que a conta de origem enxerga a base padrão (edições pessoais > mestre > seed — a mesma ordem de `get_merged_orcamento`), gera `.env` limpo (nuvem/IA só com `--com-nuvem`/`--com-chaves-ia`; a senha do usuário não entra no `.env`), `Instalar.bat`, `LEIA-ME.txt`, `SHA256SUMS.txt`, `ACESSO_*.txt` (senha, fora da pasta e do zip) e `--zip`.
- **Conferência obrigatória:** roda o programa sobre uma cópia do pacote (init do banco, `/api/health`, login certo/errado, `/api/orcamento/dados`, `/api/projetos`); se falhar, apaga a entrega parcial e sai com erro. Recusa também: e-mail/senha inválidos, base vazia, base de mais de um dono sem `--base-de`, banco inválido.
- Testes: `tests/test_pacote_entrega_230.py` (inclui varredura de bytes: nada do dono sobra no pacote). README: seção «Entregar o OrçaCore a outra pessoa».

## 2026-11-23 — TASK-229: pontos fora das coordenadas do CAD — verificação e «Posições do CAD»

- **Verificação:** a colocação inicial dos pontos pelas coordenadas do CAD (`_poste`) foi comparada entre a versão 2.7.8 (antes dos grupos/Rearranjar) e a atual com a mesma obra (12 pontos, ruas tortas, 2 estruturas sem coordenada): **idêntica, ponto a ponto**. A importação do `.obra` do plugin também preserva `_poste`/`_geo`/`_geoTipo`. Não há regressão no desenho inicial.
- **Causas prováveis de o ponto não aparecer na coordenada:** (1) posição **ajustada à mão** gravada na linha (`_mapaPoste`: arrastar, criar/copiar ponto, «Agrupar pontos» e **Rearranjar** gravam) — vale no lugar do CAD ao reabrir o projeto; (2) a escolha **1:500** fica lembrada e a rede maior que a janela mostra só um pedaço.
- **Novo:** botão **Posições do CAD** (descarta as posições ajustadas à mão — campo do modo atual — e os «puxões» dos cabos; um passo de Ctrl+Z); o aviso do mapa diz quantos pontos têm posição ajustada e, em 1:500, quando a rede não cabe na janela.
- Arquivos: `static/vinculacao_ui.js`, `tests/test_mapa_posicoes_cad_229.py`.

## 2026-11-23 — TASK-228: «Rearranjar» deixava os pontos longe dos cabos

- **Causa:** a ponta de um cabo só se liga a um poste a menos de 5 m (`RAIO_POSTE_PADRAO`). Ponta mais longe fica «solta» e era desenhada na coordenada fixa do CAD. O «Rearranjar» (e qualquer arraste) movia os pontos, mas o cabo de ponta solta ficava para trás; e como o ponto não tinha vão ligado, o rearranjo o tratava como **isolado** e o jogava numa grade ao lado → pontos agrupados longe dos cabos.
- **Correção:** cada ponta solta é ancorada visualmente no ponto mais perto (até 25 m do CAD; `d.ancoras`, `pontaAncorada`): a ponta anda junto com o ponto (mesmo deslocamento) e conta como ligação para o layout do «Rearranjar». A ligação de verdade (vínculo/topologia) continua exigindo 5 m — nada muda no orçamento. O alerta «⚠ A linha do desenho termina a mais de 5 m de um poste» (TASK-227) segue marcando esses cabos.
- Arquivos: `static/vinculacao_ui.js`, `tests/test_mapa_ancora_pontas_228.py`.

## 2026-11-22 — TASK-227: nível MT/BT igual na lista e no mapa; diagnóstico dos cabos que não aparecem

- **Causa do MT na lista × BT no mapa:** a lista (coluna Nível / `extraCabo`) deriva o nível do **rótulo** (13,8kV ou dicionário), e o mapa (`gruposVisuaisDoVao`) dava prioridade ao **tipo da linha do desenho** (`_geoTipo`, cor+layer da rede MT/BT que o CAD associou ao rótulo). Quando o CAD ligava um rótulo MT a uma linha BT (redes paralelas na mesma rua, raio de busca) o mapa mostrava BT.
- **Correção:** `nivelDoCabo` (vinculacao_regras.js) é a conta única: nível digitado na linha > tensão no rótulo > nível padrão do dicionário > tipo da linha do desenho > BT. O desenho só decide quando o rótulo não diz; `extraCabo` (lista e regras de vínculo) usa o desenho como último recurso em vez de assumir BT.
- **Diagnóstico no mapa (`diagnosticoDoCabo`):** o cartão do cabo mostra «⚠ Nível MT pelo rótulo, mas a linha do desenho é BT» (e a aresta fica tracejada) quando há divergência — é sinal de associação errada no CAD — e o motivo quando o cabo não aparece: sem linha do desenho e sem ponta ligada (não é desenhado: `semVao`), sem linha do desenho (só a ponta ligada), ou linha cuja ponta fica a mais de 5 m de um poste. Resumo no aviso do mapa e novo filtro **«A conferir (nível/geometria)»**.
- Não alterado: o `.obra` do plugin (C#) segue gravando só `_geo`/`_geoTipo` (`nivel` não é gravado) — a associação rótulo↔vão acontece no CAD.
- Arquivos: `static/vinculacao_regras.js`, `static/vinculacao_ui.js`, `tests/test_nivel_geometria_mapa_227.py`.

## 2026-11-22 — TASK-226: posicionar o ponto novo com o mouse; copiar pontos e cabos no mapa

- **Posicionar:** depois de confirmar «+ Poste», «+ Estrutura» (avulsa) ou «+ Poste fictício», o mapa entra em modo de posicionamento: o ponto acompanha o mouse, **o clique fixa** o lugar (grava como os demais, `_mapaPoste`/`_mapaPoste500`) e **Esc mantém** o lugar sugerido (o modal não fecha). Durante o posicionamento os pontos e vãos ficam transparentes ao mouse (`pointer-events:none`) e dá para andar/zoom pelo fundo; um clique logo depois de arrastar o fundo não fixa.
- **Copiar ponto:** botão no painel do ponto — cria um grupo novo (nome automático, mesmo comentário) com cópias das linhas do ativo, **sem vínculos com cabos, sem coordenadas do CAD e sem posição**; a estrutura de um poste acompanha a cópia (`_mapaPosteRef` remapeado para o id novo); poste fictício continua congelado/bloqueado. Depois a cópia pede para ser posicionada.
- **Copiar cabo:** botão «Copiar» no cartão do cabo — copia o texto/nível/fases para o fim de Cabos, **sem pontas ligadas**, e já seleciona a ponta livre para ligar clicando num ponto.
- Ambos são um passo de Ctrl+Z (`copiarPontoNoMapa`/`copiarCaboNoMapa` em `static/resumo.js`; lista de campos que não vão para a cópia: `CAMPOS_FORA_DA_COPIA`).
- Arquivos: `static/vinculacao_ui.js`, `static/resumo.js`, `tests/test_mapa_copiar_posicionar_226.py`.

## 2026-11-22 — TASK-225: pontos do mapa como grupos, posições estáveis e «Rearranjar»

- **Bug:** adicionar um poste no mapa reorganizava os pontos. Causa: o encaixe «Ajustar» dependia do conjunto de pontos (a largura útil mudava quando surgia o primeiro ponto sem CAD e a grade dos pontos sem coordenada era refeita). Agora, **antes de qualquer inclusão** (poste, estrutura, poste fictício, «Agrupar pontos») todos os pontos têm a posição gravada (`_mapaPoste`, ou o deslocamento em metros em 1:500) e o ponto novo cai num lugar livre perto do centro da janela.
- **Tudo é ponto** (poste, estrutura avulsa, poste fictício): cada ponto é um **grupo em Demais Ativos** (TASK-195): título `@nome`, uma linha de comentário `#` e a(s) linha(s) do ativo. Criar um ponto pelo mapa insere essas 3 linhas (nome vazio = número automático `P<n>` que não repete os `P<linha>` dos pontos ainda sem grupo; nome repetido é recusado). Estrutura de um poste entra no fim do grupo desse poste.
- **Painel do ponto:** nome (ou automático), comentário, **Congelar/Bloquear grupo** (age em todas as linhas, como no Resumo; não vale para poste fictício, que já é congelado/bloqueado) e a lista «Linhas nas tabelas» com a linha do grupo, o comentário, o ativo em Demais Ativos e os cabos ligados. Ponto que ainda não é grupo: «Criar o grupo do ponto». Rótulo do mapa e «Ponto <nome>» usam o nome do grupo.
- **Agrupar pontos (N):** transforma de uma vez os pontos existentes em grupos `@P<linha>` (mantém o nome que o ponto já tinha no mapa; um passo de Ctrl+Z; os vínculos por índice acompanham a reordenação).
- **Rearranjar:** novo botão (módulo `static/mapa_layout.js`, sem sorteio — mesmo mapa, mesmo resultado): molas nos vãos + repulsão por parte conectada, depois grade de passo igual ao vão (alinhado, nenhum ponto em cima de outro), partes lado a lado, gira 90° se a rede couber maior, vãos retos. Grupos **bloqueados** não saem do lugar. Em 1:500 o botão devolve os pontos à posição do CAD (a escala manda). Ctrl+Z desfaz.
- `services/obras_arquivo.py` passa a guardar `grupoDecl`, `gid`, `grupoDe`, `grupoRecolhido` (antes os grupos da TASK-195 se perdiam no arquivo `.obra`). Correção junto: em 1:500 sem coordenadas do CAD o arraste usava o modo errado (agora vale «Ajustar»).
- Interpretação das «três linhas»: ao criar o ponto entram 3 linhas em Demais Ativos (grupo `@`, comentário `#`, ativo) e o painel mostra essas linhas e as dos cabos ligados.
- Arquivos: `static/mapa_layout.js` (+ rota `/static/mapa_layout.js` sem cache), `static/grupos_linhas.js`, `static/resumo.js`, `static/vinculacao_ui.js`, `static/resumo.html`, `services/obras_arquivo.py`, `app.py`, `tests/test_pontos_mapa_225.py`.

## 2026-11-22 — TASK-224: mapa de vinculação com escala 1:500 e edição mais leve

- **Escala:** seletor «Ajustar à tela / 1:500» no mapa (lembra a escolha). O desenho do CAD está em metros (vãos de 20–50 m); em 1:500, 1 m = 2 mm = 7,559 px (96 dpi; aproximado, depende do monitor) e a vista abre centrada em 1 px = 1 unidade. Indicador «Escala ≈ 1:N» com barra de escala em metros, nos dois modos (muda com o zoom). Em 1:500 o arraste manual é guardado como **deslocamento em metros** (`_mapaPoste500`, `_mapaCabo500`), separado do arraste do «Ajustar» (`_mapaPoste`/`_mapaCabo`, fração da tela): trocar de escala não mistura as duas disposições. Os campos novos passam por desfazer, salvar obra, duplicar linha e `services/obras_arquivo.py` (|n| ≤ 100000).
- **Fluidez (redes grandes):** (1) **recorte por janela** — acima de 400 itens só pontos/vãos perto da janela ficam no DOM (`display:none` ainda custava layout/estilo a cada quadro; por isso saem do DOM, em duas camadas `vz-camada-vaos`/`vz-camada-nos`); (2) **pan/zoom uma vez por quadro** (antes cada evento do mouse/roda mexia no viewBox) e a âncora do zoom usa a vista em memória; (3) **rótulos somem enquanto mexe** em rede grande (o texto SVG era o que mais pesava); (4) arrastar um poste/cabo só redesenha os vãos dele (índices por ponto e por cabo, em vez de varrer todos).
- **Medição** (2400 postes, 4700 cabos, Chromium headless com raster por software, ms por passo): arrastar poste 52 → 10 · mover o mapa 90 → 11 · zoom 114 → 13 em 1:500; em «Ajustar» (rede inteira visível, sem o que recortar) o zoom foi de 114 → 64 e o pan de 90 → 79. Rede pequena (96 postes): sem diferença visível (abaixo do limite de recorte).
- **Achados não corrigidos:** `pintar()` refaz tudo a cada alteração (~0,9 s em 2400 postes; JS ~0,3 s, o resto é o navegador) e a lista lateral cria um cartão por cabo (4700 nós); `convencao_postes.compilar` serializa a configuração a cada chamada (~220 ms por repintada nesse porte). Candidatos a uma próxima rodada.
- Arquivos: `static/mapa_escala.js` (contas puras, `/static/mapa_escala.js` sem cache), `static/vinculacao_ui.js`, `static/resumo.js`, `static/resumo.html`, `services/obras_arquivo.py`, `app.py`, `tests/test_mapa_escala_224.py`.

## 2026-11-20 — TASK-207: Pernambuco classifica estrutura solta como ESTRUTURA

- Antes: no Pernambuco (081) só havia regras de POSTE e CABO; `N3`, `B2`… soltos ficavam com entidade `0`. Nova regra (ordem 30) no seed `data/regras_leitor_classificacao_pernambuco_seed.json`: estrutura (`N`, `U`, `T`, `R`, `B`, `S`, `CE`, `SI`, `RA`, `BI` + número, `ISOL`), sozinha ou em lista, com quantidade e `(parâmetros)` → `ESTRUTURA`. Poste (`11300D, N3`) e cabo seguem antes; `CON`, `PR220`, `FLABIT` continuam `0`.
- O hash do seed anterior entrou em `seeds_antigos.classificacao` do pacote 081: projetos com as regras **intocadas** atualizam sozinhos; quem editou as regras precisa importar/acrescentar a regra à mão. Na vinculação `ESTRUTURA` e `0` já valiam igual (sem mudança). Teste: `tests/test_estrutura_solta_pe_197.py`.

## 2026-11-19 — TASK-196: congelar e bloquear grupos de linhas

- Congelar/bloquear (menu do botão direito e botões ❄ 🔒) na linha do grupo («@») age em **todos os membros**; o menu passa a dizer «Congelar/Descongelar/Bloquear/Desbloquear grupo». A declaração continua sempre congelada (não desce disso); o bloqueio vale também para ela. `expandirGrupos` em `static/resumo.js`; teste em `tests/test_grupos_linhas_195.py`.

## 2026-11-19 — TASK-195: grupos de linhas («@P1») em Outros e Cabos, mover linhas e menu de contexto que não corta

- **Declaração:** linha cujo Ativo começa com `@` (ex.: `@P1 Rua das Flores`) é um grupo (`#` continua sendo comentário). Fica **congelada** (fora de totalizadora/orçamento/validação/Resumo da rede), com `grupoDecl`+`gid`; os **membros** (`grupoDe`) ficam sempre colados logo abaixo. Digitar `@` marca ao sair da célula (ou no Enter); apagar o `@` desfaz o grupo e solta os membros.
- **Enter** numa linha de grupo (declaração ou membro) cria um membro no **fim** do grupo; **Shift+Enter** cria linha **solta imediatamente abaixo do grupo**. Em linha comum nada muda (Enter e Shift+Enter continuam iguais).
- **Menu do botão direito:** novo 1º item **Grupo ▸** (as vistas trocam o conteúdo do mesmo menu): *Adicionar ao grupo ▸ [lista de grupos]*, *Criar grupo…* (pede o nome; vale para uma ou várias linhas selecionadas), *Tirar do grupo*, *Recolher/Expandir o grupo*, *Desfazer o grupo*, *Mover para cima/baixo*.
- **Recolher/expandir (+/−):** botão no número da linha da declaração com a contagem de linhas; reaproveita o mecanismo de linhas ocultas do filtro (navegação por teclado, seleção e cópia já ignoram oculta). Alt+←/→ também. O estado é salvo.
- **Mover:** Alt+↑/↓ (linhas selecionadas ou a atual; numa declaração, o bloco todo). Item que passa da borda sai do grupo; linha solta que desce sobre a declaração entra nele; grupo recolhido é atravessado de uma vez. Em Outros os vínculos cabo↔estrutura (por índice) são remapeados.
- **Integração:** `deepClone` leva os campos (desfazer/refazer/salvar obra); ordenar por coluna mantém os membros colados; excluir a declaração solta os membros; duplicar a declaração não gera 2º grupo; importar/subtrair obra gera ids novos de grupo; congelar/comentar não se aplicam à declaração.
- **Menu cortado (correção):** o menu de contexto, o menu de ordenar e o «Exibir» agora são medidos depois de abertos e **viram para cima/esquerda** quando não cabem, com rolagem se for mais alto que a tela (antes a altura era fixa em `innerHeight-140`).
- Sem nenhuma linha `@` nada muda: bateria de regressão de 44 passos (Enter, Tab, #, congelar, bloquear, duplicar, ordenar, vincular, Ctrl+Z…) idêntica à anterior, exceto o novo item no menu.
- Arquivos: `static/grupos_linhas.js` (lógica pura, `/static/grupos_linhas.js` sem cache), `static/resumo.js`, `static/resumo.html`, `app.py`, `tests/test_grupos_linhas_195.py`.

## 2026-11-19 — TASK-194: balão de parâmetros grava também quando o ativo tem quantidade («6 CON»)

- **Sintoma:** com «CON» o balão preenchia os parâmetros; com «6 CON» (ou `6-CON`, `*2 CON`, `11300D, 2 N4`) as setas navegavam, o Enter «escolhia» e nada era gravado. **Causa:** a leitura dos parâmetros procurava a dica pelo nome SEM a quantidade (`it.base`, «CON») e a gravação (`gravar` em `resumo.js`) pelo nome COM ela (`it.nome`, «6 CON», que não existe na base): a dica vinha vazia e a função saía em silêncio.
- **Correção:** a gravação virou `AtivoItens.gravarParametro` (pura, `static/ativo_itens.js`): pergunta a base pelo nome sem quantidade e mantém a quantidade escrita no texto (`6 CON(cabo=A70,…)`, `*2 CON(tipo=COMP)`). Também: vírgula entre dígitos (`0,5 CON`) deixou de ser tratada como separador de itens em `AtivoItens.tokens`.
- Verificado no navegador com a base de Pernambuco (`CON`, `6 CON`, `6-CON`, `*2 CON`, `11300D, 2 N4`, `3 CON, 2 CON`). Testes: `tests/test_balao_quantidade_194.py`.
- **Ainda em aberto (outro comportamento):** parâmetros declarados abertos na base (`cabo=%`, `cabo2=%`) só oferecem «✎ digitar outro valor…»; o Enter com o campo vazio não grava e segue para o próximo.

## 2026-11-19 — TASK-193: «Lembrar de mim» realmente entra direto no desktop e escolher arquivo não trava a tela

- **Login:** o desktop (`app.py`) abria SEMPRE em `/login?new_session=1`, e `login.html` apagava token e «lembrar» ao ver esse parâmetro: a sessão de 30 dias (commit 70d5994) era destruída a cada abertura. Agora a janela abre na tela inicial (`inicio_url`; com obra do nanoCAD, direto na tela da obra) e `login.html` encaminha para `next`/`/` quando há sessão lembrada. Sem sessão, token vencido ou «lembrar» desmarcado, as telas e o `auth_fetch.js` mandam ao login como antes (uma vez, sem laço). `?new_session=1` digitado à mão ainda limpa a sessão.
- **Arquivo escolhido (desktop):** a caixa «Abrir» fazia o Python ler o arquivo todo, virar base64 (+33%) e cruzar a ponte JSON do WebView2, e o JS decodificava antes de a tela reagir (segundos sem sinal em DXF/PDF grande). Agora `abrir_arquivos` devolve só `{id, nome, tamanho}` (sem ler o arquivo), a janela mostra «Abrindo <nome> (N MB)…» na hora e os bytes vêm por `GET /api/desktop/arquivo/<id>` (binário, id aleatório de 10 min, só arquivos escolhidos na caixa nativa, exige login). `services/arquivos_desktop.py`, `routers/desktop_arquivos.py`, `static/desktop_arquivos.js`.
- Verificado no navegador: aviso em ~1 ms e 40 MB no campo em ~240 ms (a ponte real do WebView2 não dá para medir fora do Windows; o ganho principal é tirá-la do caminho dos bytes). Testes: `tests/test_desktop_abrir_193.py`.

## 2026-11-19 — TASK-192: Resumo da rede conta corretamente em Pernambuco (poste compacto e linha viva)

- **Sintoma:** na nota real de Pernambuco (DXF 9101578729, 74 cabos, 35 linhas de Outros) o resumo mostrava **0 postes, 282 m de média, 276 m de baixa, 0 equipamentos**. As regras fixas serviam à Paraíba.
- **Causas:** (1) a base de Pernambuco cadastra o poste com o nome oficial compacto (`12600D=12600D(fases=TRI), DT12/600`), o item chegava como `12600D` e a regra `^(DT|RC|CV|FV)` não o reconhecia (postes = 0; os `TR` do poste também sumiam das linhas `*I`); (2) só a operação `I` exata contava: as linhas `*I` (linha viva), metade da nota, ficavam fora de postes, equipamentos e cabos.
- **Correção:** a regra saiu de `resumo.js` para `static/resumo_rede.js` (sem DOM, testável em Node). Poste = lista fixa (Paraíba) OU a **convenção de postes do projeto** (`AnotacaoOutros.ehPoste`), com ou sem parâmetros; «instalando» = `I` ou `*I` (R, *R, M, *M e vazia não contam). Equipamento (TR/CFU/CFA) e as regras de média/baixa não mudaram. Resultado na mesma nota: **14 postes, 443 m, 340 m, 1 equipamento**, iguais a uma contagem independente feita nos dados.
- **Atenção (todos os projetos):** `*I` agora conta como instalando também na Paraíba. Testes: `tests/test_resumo_rede_192.py` (Pernambuco com base, sem base, parâmetros, falsos positivos, linha viva, Paraíba no formato antigo, cabos). `tests/test_resumo_rede_116.py` agora lê o módulo novo.

## 2026-11-19 — TASK-191 (ajuste 3): assistente rápido e sem travar em «Gerando»

- **Nunca mais «Gerando» para sempre:** cada chamada ao provedor tem tempo máximo (45 s; 20 s se há reserva configurado; `ErroTempo` troca de modelo/provedor sem repetir) e a pergunta inteira tem teto de 120 s. A tela mostra «Pensando · N s» (e o nome da etapa: «Consultando a base», «Calculando o orçamento»…), avisa quando demora (15 s) e o botão de enviar vira **■ Parar** (aborta o pedido).
- **Menos idas ao modelo por comando:** plano limpo (sem aviso nem rejeição) ENCERRA a pergunta com um resumo local («Proponho N alterações… clique em Aplicar») em vez de mais uma volta só para o modelo resumir; a skill `formato_ativos` e a tela atual (até 60 linhas, com ids) já vão no prompt (sem `carregar_skill`/`ler_tabela`); o prompt manda fazer o mínimo de passos e chamar ferramentas em paralelo; `propor_alteracoes` já confere a base (sem `validar_linhas` antes). Gemini 3 usa pensamento baixo (`thinking_level=low`, com recuo automático se o modelo não aceitar). As ferramentas rodam em outra thread (cálculo grande não trava o servidor).
- **Atalho sem IA ampliado:** «adicione 2 CON, 3 PR220 e 1 CFU», «remova 1,5 IT; 2x CON», «tire 2-CON» (vários itens, verbos adicione/coloque/insira/remova/tire/retire/exclua); frase que não seja só «quantidade ativo» (ex.: «adicionar 2 postes de 11 metros») vai para a IA em vez de virar um ativo errado.
- Testes: `tests/test_assistente_191.py` (terminais, prazo, tempo esgotado, thinking, atalho, prompt) e `tests/test_assistente_front_191.py`; verificado no navegador (contador, Parar, atalho com 3 itens em ~0,8 s).

## 2026-11-19 — TASK-191 (ajuste 2): provedor reserva automático no assistente

- **Configuração** (Configuração de IA → «Provedor reserva»): ativar, provedor (OpenAI-compatível / Gemini / Claude), URL base, modelo e chave, com atalhos DeepSeek, Qwen (Alibaba · Singapura), OpenRouter, Ollama, Gemini e Claude. DeepSeek/Qwen/OpenRouter/Ollama entram como «OpenAI-compatível». Gravado em `configuracoes` (`ai_reserva_*`); a chave nunca volta para a tela (`GET /api/assistente/reserva` só diz `tem_chave`). `PUT /reserva` e `POST /reserva/testar` (chamada mínima ao reserva salvo): só admin.
- **Comportamento** (`services/assistente/agente.py`, `AdaptadorComReserva`): qualquer falha do principal (503, cota, chave, rede) passa ao reserva na MESMA pergunta e o chat mostra um aviso («Usando o provedor reserva…»). Com reserva, o principal desiste mais cedo (2 tentativas por modelo em vez de 3). Falha pausa o principal por 90 s (as perguntas seguintes vão direto ao reserva, sem esperar); depois volta a ser tentado. Se os dois falharem, a mensagem traz os dois motivos. Sem chave no principal mas com reserva ativo, o reserva atende (antes: 401). O atalho «adicionar 2 CON» agora roda ANTES de resolver a chave (funciona sem chave nenhuma).
- **Correção:** `static/auth_fetch.js` tratava QUALQUER 401 de `/api/` como sessão expirada; o chat e a lista de modelos respondem 401 quando falta a chave da IA («API Key não encontrada…»), o que mandava o usuário sem chave do Gemini para o login ao abrir a Configuração de IA. Agora 401 com `detail` começando por «API Key» não derruba a sessão.
- Testes: `tests/test_assistente_191.py` (reserva, pausa com relógio falso, rotas, validações, sem expor a chave) e `tests/test_assistente_front_191.py` (atalhos, corpo do PUT, `auth_fetch`). Verificado no navegador: atalho DeepSeek, salvar com e sem chave, aviso da troca no chat.

## 2026-11-19 — TASK-191 (ajuste): assistente tenta de novo quando o provedor está sobrecarregado (503)

- Erro `503 UNAVAILABLE … high demand` do Gemini derrubava a pergunta: o fallback trocava de modelo sem repetir o mesmo antes. Agora cada modelo é tentado até 3 vezes (esperas de 1,5 s e 4 s) quando a falha é de momento (503/502/504, indisponível, sobrecarga, timeout, conexão); cota esgotada ou modelo inexistente trocam direto, sem esperar. Se tudo falhar, a tela mostra «O provedor de IA está sobrecarregado agora… repita ou mude o provedor/modelo em Configuração de IA» em vez do JSON cru. `services/assistente/agente.py` (`com_fallback`, `mensagem_amigavel`); testes em `tests/test_assistente_191.py`.

## 2026-11-19 — TASK-191: assistente de IA com ferramentas e skills (consulta a base, monta, totaliza e orça)

- **Antes:** chat de uma volta só, com o catálogo de ativos colado num prompt estático; a IA respondia uma tabela markdown que o front lia por regex. **Agora:** laço de agente com **ferramentas** (`services/assistente/agente.py`: Gemini, OpenAI e Claude por adaptadores; fallback de modelo; até 8 passos por pergunta) e **skills** (arquivos `.md`).
- **Ferramentas** (`services/assistente/ferramentas.py`): `buscar_ativo`/`detalhar_ativo` (base real do projeto — `catalogo.py`), `ler_tabela`, `validar_linhas`, `propor_alteracoes`, `totalizar`, `calcular_orcamento` (Totalizadora do projeto → `processar_calculo`, com `com_plano` para simular), `listar_obras`/`consultar_obra`/`propor_obra`, `ajustar_tela`, `listar_skills`/`carregar_skill`. **Nada altera a tabela no servidor**: a tela manda uma cópia das linhas a cada pedido (`id` = posição) e a resposta traz um **plano**; o navegador mostra o cartão (＋ ✎ − com avisos), aplica só no clique em «Aplicar» (um passo de histórico: Ctrl+Z/«Desfazer»), confere que a linha ainda é a mesma e ignora o que mudou. O atalho «adicionar 2 CON» continua sem IA e aplica na hora.
- **Skills** (`services/assistente/skills.py`, `data/skills/_geral|<projeto>/*.md`): pacote (versionado) + pasta do usuário (`skills/` ao lado do .exe; vence o pacote); o índice (nome + `descricao:`) vai no prompt e o conteúdo só com `carregar_skill`. Vêm `formato_ativos`, `montar_projeto`, `orcamento` e `pernambuco/projeto`; o botão «Gerar catálogo de ativos» (admin) grava a skill `ativos` do projeto lendo a Base de Orçamento. Janela 📚 no painel do assistente: listar, editar, criar, apagar.
- **API** (`routers/assistente.py`): `POST /api/assistente/chat` (NDJSON: ferramenta · texto · plano · acao_ui · erro · fim), `GET/PUT/DELETE /api/assistente/skills[/escopo/nome]`, `POST /api/assistente/skills/gerar-catalogo` (gravar/apagar/gerar: só admin). `ai_chat.resolver_provedor()` foi extraído do chat antigo (`/api/gemini/chat` segue igual) e é reaproveitado.
- **Formato correto de Outros** nas skills/prompt: compacto (`11300D, N1, 2 N3`, quantidade antes) ou antigo (`DT11/300 1-N1`), nunca misturados — a mistura é apontada pela própria Totalizadora.
- Testes: `tests/test_assistente_191.py` (skills, catálogo, ferramentas, adaptadores com SDK falso, laço, rotas) e `tests/test_assistente_front_191.py` (Node). Verificado no navegador (Playwright) com a resposta simulada: plano → Aplicar → Desfazer, atalho rápido, Descartar, janela de Skills (catálogo, nova skill).

## 2026-11-19 — TASK-190: cabo com condutor extra «+16» (`3XC16(16)+16 35 M`) é reconhecido e vira 5 ativos na Totalizadora

- A notação do label (`anotacao_cabos`) só conhecia `+TENSÃOkV`; `3xC16(16)+16` dava «rótulo não reconhecido» e o cabo não entrava na Totalizadora. Agora `+N` (sem kV) é **mais um condutor da mesma família com a seção N** (`neutroDe`): `3xC16(16)+16 35 m` = 3 fases C16 + neutro C16 + 1 C16 = **5 linhas** `C16 35` na Totalizadora/payload. `+20kV` segue como tensão. JS (`static/anotacao_cabos.js`) e Python (`services/anotacao_cabos.py`) em paridade; a coluna «Qtd ativos» do Cabos continua contando as fases (3), como em `3xA35P(7,94)`. Testes em `tests/test_anotacao_cabos.py`.

## 2026-11-19 — TASK-189: correção — Resumo com muitas linhas ao abrir e aviso ao fechar o programa

- **Botões do Resumo sem resposta e só Cabos aparecendo** depois de ler um DXF e abrir a aba: o primeiro `renderTable` (com mais de 60 linhas já gravadas) repinta a seleção em intervalo ao montar as linhas (`_terminouHidratar` → `intervaloGrade`), e a `const grade` ainda não tinha sido declarada (era mais abaixo no arquivo): «Cannot access 'grade' before initialization» derrubava a inicialização — nenhum botão era ligado e a tabela Outros nem era desenhada. A TASK-185 só aparecia com linhas coladas depois do carregamento; agora `grade` fica no topo, com os demais estados do 1º desenho. Reproduzido no navegador com as 662 linhas do DXF da nota 9101578729 (Cabos 74 / Outros 49).
- **Aviso ao fechar** (`ConnectionResetError: [WinError 10054]` em `trabalhador_js.py`): o trabalhador do leitor (QuickJS) lia o socket enquanto o programa principal fechava; agora `main()` encerra quieto nos erros de conexão.
- Testes: `tests/test_resumo_carrega_com_dados_189.py`.

## 2026-11-19 — TASK-188: Leitor lê o comprimento do cabo POR POSIÇÃO no desenho

- **Problema (conferido no DXF da nota 9101578729):** a vizinhança das regras do Leitor olha os N itens seguintes **da lista**, e a lista do DXF sai **ordenada por altura (Y)**; o comprimento de um cabo caía a 26–37 itens do rótulo (janela 16) ou antes dele. Resultado: cabo sem comprimento (rótulos riscados `2xC25-13,8kV`/`1xS04-13,8kV` do print) ou, pior, com o comprimento de OUTRO cabo — de 40 rótulos, 22 pegavam um texto a 14–134 unidades de distância enquanto o certo estava a 2–7 unidades logo abaixo.
- **Novo na vizinhança (`regras_leitor_engine.js`, `_vizinhoPorPosicao`):** `geometrica: true` + `raio` (unidades do desenho, padrão 12) + `posicao` (`ABAIXO` padrão, `ACIMA`, `QUALQUER`): usa `_x/_y` dos itens e escolhe o texto da regex **mais perto** (empate: o mais próximo na lista; só da mesma página). Com coordenadas e nada dentro do raio a regra não casa (não cai para a ordem da lista); sem coordenadas (PDF) usa a janela/ordem de antes. Validação em `routers/regras_leitor.py` e campos no editor de regras (Por posição no desenho · Raio · Onde).
- **Rótulo com neutro** (`3xC16(16)+16`, `3xC35(35)+35`): a regex dos cabos aceita `+N` antes do `-13,8kV`.
- **Regras do projeto Pernambuco (081)** ajustadas e entregues em JSON (regra «Cabo: label sozinho» por posição, raio 12; `+N`). No DXF da nota: cabos reconhecidos 65 → 74, todos com comprimento; os dois riscados e o de cima passam a ler `27 m`. Testes: `tests/test_leitura_por_posicao_188.py`.

## 2026-11-18 — TASK-187: ajustes do uso do balão, sugestões, setas e menu de contexto

- **Balão aberto pelo mouse:** só abre com o mouse parado (~1,5 s) na linha **selecionada** (campo com o foco); em linha não selecionada nada abre. Aberto assim, já tem o teclado: setas, Tab, Enter e Esc funcionam com o mouse em cima; mexeu no teclado, ele deixa de fechar quando o mouse sai.
- **«sem parâmetros» indevido:** o status de cada item agora casa com o item interpretado **na mesma posição** (`_statusParametrosDaLinha(..., alinhado)`); o casamento por nome falhava com postes (`11300D` → `DT11/300`), apelidos e conversões. Sem o mesmo número de itens, cai para o nome. Origem dos valores vem do status do próprio item.
- **Sugestão de ativos:** ao **digitar** o nome exato («PR») a sugestão aparece (`sugerir(..., incluirExato)`); ao só clicar num campo já preenchido continua não abrindo à toa; Tab com a sugestão igual ao texto segue para a próxima célula. A sugestão fica escondida enquanto o balão está aberto por classe (`body.ab-aberto`), não por `style` (que a deixava invisível para sempre).
- **Seta ↓** da última linha preenchida vai para a primeira linha em branco.
- **Menu de contexto** (botão direito ou tecla de menu): a célula em foco é guardada e **volta a ter o foco** depois da ação ou do Esc (Tab continua dali); clicar em outro lugar mantém o foco onde o clique foi.

## 2026-11-18 — Correção: ativo com quantidade (`6 CON`, `6-CON`) é reconhecido no balão

- O balão procurava na base o texto inteiro do item («6 CON»), por isso mostrava «ativo não encontrado na base». Cada item agora tem `base` (o nome **sem a quantidade**, `AtivoItens.semQuantidade`: `6 CON`/`6-CON` → `CON`, `*1,5-IT_2` → `IT_2`; postes/códigos como `12600D` ficam como estão) usada na busca da dica, no status e nas origens; `nome` continua como foi escrito. Em linhas com vírgula o formato é `6 CON` (o antigo `6-CON` continua sendo aviso da tabela, mas o balão já edita os parâmetros dele).

## 2026-11-18 — Correção: o balão volta a ter a linha «✎ digitar outro valor…» em todo parâmetro

- No balão de parâmetros (TASK-186) a linha de digitar só aparecia em parâmetros «abertos»; em listas fechadas (ex.: `tipocabo`) o usuário não conseguia escrever. Agora **todo** parâmetro tem a linha «✎ digitar outro valor…» (a última), uma **letra digitada** com o parâmetro marcado cai direto nela (Enter grava, Esc volta) e, se o valor atual não está na lista, ele já vem escrito ali.

## 2026-11-18 — TASK-186: «Ocultar parâmetros» com campo só de nomes, balão interativo dos ativos/parâmetros e Tab na tabela Outros

- **Campo do ativo (Outros) com «Ocultar parâmetros» ligado:** mostra e edita só os **nomes** (`12600D, N1, N4`) — cursor, escrever, selecionar, colar, copiar normais; `row.ativo` guarda o texto cheio. Os **parâmetros ficam presos a cada item** (`static/ativo_itens.js`: `itens`, `compacto`, `aplicarEdicao` por diferença de texto + posição do cursor): apagou ou renomeou = item novo, sem parâmetros; item repetido (`N1, N1`) guarda os SEUS. Colar texto com parênteses leva os parâmetros junto. Desligada: texto inteiro, como antes. Acabou o rótulo sobreposto/campo transparente (`ativo-compacto`); sem selinho de contagem.
- **Balão interativo (`static/ativo_balao.js`)**, abre com **Ctrl+Enter**, duplo clique ou mouse parado ~1,5 s na linha; lista os ativos da linha (✓ completo, ⚠ falta, – sem parâmetros) com o que já está preenchido; **↑↓/Tab** navegam, **Enter**/clique entram nos parâmetros do ativo; lá **Tab/Shift+Tab** andam entre parâmetros, **↑↓** escolhem o valor (ou «✎ digitar outro valor…»), **Enter** grava e vai ao próximo (no último, ao próximo ativo com pendência; sem pendência, fecha), **Del** limpa, **Esc** volta à lista e outro **Esc** fecha devolvendo o foco ao campo. Grava pelo mesmo caminho do blur (`_gravarCheio` → `confirmarMudanca`), um passo do Ctrl+Z. Com a opção ligada, **`(`** não escreve: abre direto os parâmetros do ativo sob o cursor (`AtivoBalao.abrirNoCursor`). O balão só-leitura antigo continua em Cabos/Totalizadora.
- **Tab em Outros:** Operação → Ativo → **Operação da linha de baixo** (última linha preenchida: cria linha nova); Shift+Tab volta; Enter no ativo continua criando linha. Tab/Enter na Operação sempre vão ao ativo da mesma linha.
- Alternar a opção em Configurações refaz a tabela Outros. `AtivoDica.contextoAtual()` exportado. Testes: `tests/test_ativo_itens_186.py`.

## 2026-11-18 — TASK-185b: tabelas de 1000+ linhas, teclado e achados na aba Cabos ⇄ Postes, `layer_em` com curinga, teste de desempenho

- **Tabelas muito grandes (> 300 linhas):** em segundo plano só se monta ~2500 px em volta da área visível (`LIMIAR_PARCIAL`, `hidratarJanela(type, 2500)`); o resto monta ao rolar. Com 1000+1000 linhas: Enter 67 ms, excluir 34 ms, desfazer 146 ms, rolagem 16 ms/quadro.
- **Aba «Cabos ⇄ Postes»:** teclado (↑↓ escolhem o cabo, → vai aos postes, ↑↓ escolhem o poste, Enter liga/desliga, ← volta) e painel **«aviso(s) do algoritmo por poste»** (incompleta, sobra, ponta solta, sem geometria) sob a lista.
- **Modal:** `vincular` ~40 ms (assinatura dos dados calculada uma vez por ação).
- **Leitor:** `layer_em` aceita **curinga `*`** (`POSTE_*_LV`, `*_LV`), sem diferenciar maiúsculas; sem `*` continua exato. Vale para as tabelas de Processamento e Classificação (`layerCasa` em `static/regras_leitor_engine.js`).
- **Desempenho:** `scripts/perf_resumo.js <url> <token> [linhas]` mede o Resumo e o modal com limites e sai com código 1 se passar; no pytest, `test_desempenho_com_450_linhas_em_cada_tabela` roda se `PERF_URL` e `PERF_TOKEN` estiverem definidos.

## 2026-11-18 — Hidratação das linhas só quando necessária (rajadas de ações)

- Antes, clique/tecla/menu forçavam a montagem de TODAS as linhas ainda vazias (`lazy-row`) a cada ação; em sequência rápida o custo voltava a ~120 ms. Agora só **colar/copiar/cortar/arrastar e Ctrl+A** montam tudo; as demais ações trabalham sobre os dados e as classes da linha (`row-selected`), e a seleção em intervalo é repintada quando as linhas são montadas. Rajada de Enter/excluir/duplicar/desfazer com 450+450 linhas: 28–64 ms por ação (antes 100–330 ms).

## 2026-11-18 — TASK-185: tabelas e painel de vínculo fluidos com 400+ linhas; aba «Cabos ⇄ Postes»

- **Modal de vinculação (aba Vínculo) reescrito:** lista **virtualizada** (`criarPainel` em `static/vinculacao_ui.js`): só as linhas visíveis (+ margem) existem no DOM; cada linha tem altura fixa (cabo 30 px, avulso 24 px, **poste 30 px — uma linha por poste**, com as estruturas dele lado a lado como chips `N1 1/2`), então a posição das linhas de ligação sai por aritmética (nada de `getBoundingClientRect`) e o SVG cobre só a janela visível. A **análise** (`VR.analisar`), os índices de ocupantes/`infoCabo`/`infoEstrutura` e a interpretação de cada linha de Outros (`metaAtivo`) são **memoizados** por uma assinatura dos dados (`ctx()`); selecionar um cabo só repinta a janela (~7 ms) e vincular refaz os itens (~45 ms com 450+450 linhas; antes ~300 ms). O fundo do modal ficou **opaco**: o translúcido refazia a mistura com a página a cada quadro da rolagem (rolagem 60 fps).
- **Nova aba «Cabos ⇄ Postes»:** o usuário só liga cada cabo ao(s) poste(s) a que ele chega (clique no cabo e no poste, ou o contrário; clique na linha desfaz) e o **algoritmo por poste** (o mesmo da obra do plugin, `VR.vincularPorPoste`) escolhe a estrutura de cada poste. Novo campo da linha de cabo **`vinculoPostes`** (linhas de Outros; `VR.postesAtribuidos`/`postes_atribuidos` em Python): vale **no lugar da geometria** (`_geo`/`_poste`); vínculo manual, cabo congelado e cabo sem geometria seguem como antes. Opção «Vincular as estruturas automaticamente» (ligada) e botão «Vincular estruturas agora». Persistido no `.obra` (`CAMPOS_VINCULO`), reajustado ao excluir/ordenar linhas de Outros, copiado no undo/redo.
- **Tabelas do Resumo com 400+ linhas:** acima de 60 linhas, `renderTable` monta de imediato só as linhas perto da área visível; as demais entram como linha vazia de altura fixa (`tr.lazy-row`, `tr._montar`) e são **hidratadas** aos poucos em segundo plano (`agendarHidratacao`), ao rolar (`hidratarJanela`) ou **antes de qualquer ação do usuário** que possa ler as células (mousedown/click/contextmenu/copy/cut/paste/teclas que não são caractere → `hidratarTudo`). Criar/excluir/duplicar/desfazer: 120–330 ms → ~40–60 ms com 450+450 linhas. Também: ícone da lixeira sem filhos no DOM (CSS mask), `atualizarResumoRedeUI` agendado (`agendarResumoRede`), cálculo de parâmetros por linha fora do desenho, e **sem `backdrop-filter`** nos painéis das tabelas (era o custo da rolagem: 35 → 16 ms/quadro).
- Testes: `tests/test_cabos_postes_185.py` (regras Python = JS, persistência, técnicas de desempenho presentes no código).

## 2026-11-18 — Navegação Operação ↔ Ativo, linha "digitar" nos parâmetros e preenchimento rápido

- **Operação → Ativo:** Enter no campo Operação (digitada ou escolhida na lista) leva o cursor ao **ativo da mesma linha** (Cabos, Demais Ativos, linha em branco e Totalizadora). **←** no **começo** do ativo volta para a operação (abre a lista); no meio do texto a seta continua movendo o cursor.
- **Caixa de parâmetros:** sempre há uma linha **"✎ digitar outro valor…"** (a última): chegue nela com ↑↓ (↑ vai direto), escreva qualquer valor (a linha passa a mostrar `✎ usar "X"`); Enter/Tab mantêm o valor e avançam.
- **Velocidade (investigada com a base real de ~3,9 mil linhas e o profiler do navegador):** cada tecla ao preencher parâmetros custava **100–350 ms** porque o contexto da base inteira (`AnotacaoOutros.contextoDaBase` + `assinaturasDoProjeto`, que chamam `analisarCelulaBase` e `compilar` da convenção de postes linha a linha) era refeito **várias vezes por tecla**: a caixa pede a origem dos valores (`cfg.origens` → `_contextoParametros()`, `_parametrosVinculoPorLinha()`, `_defineDoVinculo()`) e cada um refazia `_ctxOutros()` por linha. Agora `_ctxOutros()` e as assinaturas são guardados e só refeitos quando a base, o projeto ou a convenção de postes mudam; depois de carregar a base eles são montados em segundo plano (`aquecerCaches`, `AtivoDica.aquecer`). Medido: **7–45 ms por tecla** (antes 87–349 ms), também com 80 linhas na tabela.
- Testes: `tests/test_navegacao_e_velocidade.py`.

---

## 2026-11-18 — Menu pelo teclado, Duplicar linha, Enter cria linha e sugestões só da base

- **Menu do botão direito:** ↑↓ (Home/End) navegam, Enter executa, Esc fecha.
- **Duplicar linha** (menu): copia a linha clicada ou as selecionadas e coloca a cópia **logo abaixo** (em bloco, na mesma ordem); a cópia é uma linha comum (sem bloqueio e sem o vínculo/geometria do original; congelada/comentário são mantidos); um passo de Ctrl+Z (`duplicarLinhas`).
- **Enter** no campo Ativo (Cabos e Demais Ativos) **cria uma linha nova logo abaixo** e leva o cursor para ela; numa linha ainda em branco só desce.
- **Sugestões ao digitar** deixam de listar as linhas já digitadas na tabela: Cabos sugere só **ativos da base** (nome/apelido do projeto, origem CABOS) enquanto se digita o nome; Demais Ativos mantém as sugestões de ativos e a caixa de parâmetros.
- Testes: `tests/test_menu_enter_duplicar.py`.

---

## 2026-11-18 — Lista no estilo VS Code (Operação e parâmetros) e botões ❄ 💬 🔒

- **Operação:** lista no estilo do VS Code (`static/ide_lista.js`) no lugar do `<datalist>` nativo, que só mostrava as opções que casavam com o valor atual (com "I" só apareciam `I` e `*I`). Ao entrar no campo aparecem **todas** as operações (a atual selecionada, com o detalhe instalar/remover/manter e "linha viva"); ↑↓ navegam, Enter/Tab/clique escolhem, digitar filtra, Esc cancela.
- **Parâmetros (TASK-181, revisão):** assinatura `ATIVO(param, …)` com o parâmetro da vez em destaque e lista de valores como sugestões; **↑↓** navegam pelos valores; **Enter/Tab** aceitam o valor realçado, escrevem pelo nome e **avançam para o próximo parâmetro** (a caixa continua aberta; no último fecha com o cursor depois do `)`); Shift+Tab volta; Esc fecha mantendo o texto (Ctrl+Z desfaz); digitar filtra a lista. (Antes: ↑↓ trocavam o parâmetro, ←→ o valor e Enter fechava.)
- **Botões ❄ 💬 🔒** no cabeçalho das tabelas Cabos e Demais Ativos: agem nas linhas selecionadas (clique no número) ou na linha do cursor; a seleção continua marcada depois de aplicar.
- A opção **Ocultar parâmetros ao sair da linha** fica em ⚙ Configurações › Aparência › Avançado (agora já aberto por padrão).
- Testes: `tests/test_ide_lista.py`.

---

## 2026-11-18 — TASK-182 C (operação digitada) e TASK-181 (seletor de parâmetros)

- **Operação digitada (182 C):** a coluna Operação (Cabos, Outros, linha em branco e Totalizadora) é um campo de texto: maiúsculas e sem espaços (`*r` → `*R`); só aceita `I, *I, R, *R, M, *M` (`OPERACOES` é a fonte única); inválida não grava, a célula fica vermelha com "Operação inválida: use I, *I, R, *R, M ou *M" e o valor volta; Esc cancela; a mudança válida entra no histórico. `criarCampoOperacao`/`validarOperacaoTot` em `static/resumo.js`.
- **Seletor de parâmetros (181)** em `static/ativo_dica.js` (parte pura testada em Node, `tests/test_ativo_dica_181.py`): `(` depois de ativo com parâmetros já escreve `()` com o cursor no meio — **corrige o bug da vírgula** em `RL, PR, ATRINT`; o valor entra **pelo nome** (`PR(CABO2=S04)`; a tabela escreve em maiúsculas); gravar de novo troca só o valor; texto posicional (`CE_4(12600D,A35P,A35P)`) é lido como o cálculo (`resolver_pedido`) e reescrito por nome ao alterar; `PR()` vira `PR`; a caixa mostra o valor e a origem (manual, vínculo, da obra, padrão, automático) com o mesmo status do balão (`cfg.origens`). `planoDeInsercao` (TASK-149) segue exportado, mas a tela não o usa.
- Testes: `tests/test_operacao_digitada_182.py`, `tests/test_ativo_dica_181.py`.

---

## 2026-11-18 — TASK-182 (A e B): marcas das linhas gravadas no projeto e "#" = comentário

- **Causa real das marcas perdidas:** o `deepClone` que vale em `static/resumo.js` (o segundo, declarado mais abaixo, sobrescreve o primeiro) copiava só campos conhecidos e **descartava `congelada`, `bloqueada` e `comentario`** — desfazer/refazer não as restaurava e o Salvar/abrir do projeto também as perdia. Agora o clone preserva as três marcas (regressão das linhas congeladas de 17/11).
- **A:** `pushHistory`, desfazer e refazer chamam `GridRProjeto.marcarSujo(false)`; abrir/carregar um projeto (`restoreObraSnapshot`) não o marca alterado.
- **B:** ativo (Cabos e Outros) que começa com `#` vira comentário ao sair do campo, ao criar a linha, ao colar e ao abrir um projeto; tirar o `#` devolve a linha ao normal; "Tirar comentário" pelo menu numa linha com `#` também tira o `#`. Linha bloqueada não muda.
- Testes: `tests/test_linhas_congeladas.py`.

---

## 2026-11-16 — TASK-173: dicionário de cabos por projeto (notação, nível, seção, ordem do "maior")

- Lista `dicionario_cabos` na configuração da vinculação (a primeira linha que casa vale): `padrao`, `familia`, `valor_cabo`, `secao` (número ou `$1`, com tabela `awg`), `nivel`, `fases_padrao`, `apelidos`, `neutro` (`CAZ{valor}`), `notacao_solta`, `ativa`. Sem dicionário = regras de Pernambuco como antes.
- `anotacao_cabos.interpretar(texto, dic)`, `derivar_cabo/extra_cabo(…, dic)`, `secao_cabo(codigo, dic)` e `info_cabo` (`valor_cabo` do dicionário quando a Tabela de Cabos não tem; `secao` no info) consultam o dicionário; `ordem_maior` sem valor próprio usa a ordem dos valores do dicionário. Paridade Python = JS testada.
- `notacao_solta`: sem ela, um código do dicionário só vira notação quando o texto traz Nx, tensão ou parênteses (o formato antigo `CAA2 ABC 35 m` não muda).
- Aba "Dicionário de cabos" no modal Vinculação (editor de tabela + "Testar um texto"). O dicionário do 081 (`data/pacotes/081/dicionario_cabos.json`) dá o mesmo resultado do código fixo (testado).

---

## 2026-11-16 — TASK-174: convenção de postes por projeto

- Tabela `postes` na configuração da vinculação: `tipos` (`codigo`, `sufixo`, `nome`), `formatos` (formas de escrever o poste; a 1ª é o nome oficial; `{tipo}`, `{sufixo}`, `{altura}`, `{esforco}`), `altura` [mín, máx], `esforco_digitos` [mín, máx] e `compacto`. Sem tabela = Pernambuco (DT/RC/CV/FV, `11600D`), idêntico ao que era fixo.
- `anotacao_outros` (Python e JS) lê a convenção por `ctx.convencao_postes` (`contexto_da_base(..., convencao_postes)`); a poste vira o parâmetro `poste` das estruturas da linha como antes (a ambiguidade de altura/esforço continua resolvida pela base).
- `/api/orcamento/calcular` e o plugin resolvem a convenção do projeto (`postes_do_projeto`); a tela do Resumo usa a da configuração carregada. Aba "Postes" no modal Vinculação com "Testar um nome".
- Limite: os modos autônomos (`services/autonomo`) ainda usam o padrão de Pernambuco.

---

## 2026-11-16 — TASK-175: leitura do desenho (DXF) configurável por projeto: layers e cores de rede

- A configuração de leitura do projeto (`plugin_leitura_config`, a mesma do plugin do nanoCAD: `layers_rede_mt`, `layers_rede_bt`, `blocos_poste`, `layers_poste`, `layers_guia`) agora também vale na extração do DXF no servidor: `classificar_rede`, postes e guias seguem a configuração; sem ela, o padrão de Pernambuco (idêntico ao anterior, testado).
- Campos novos opcionais: `cor_mt`/`cor_bt` (ACI, `#RRGGBB`, faixa RGB ou lista) e `regra` (`COR_MANDA` padrão | `LAYER_MANDA` | `SO_LAYER`; `cor_confirma: false` do plugin = `SO_LAYER`). Layer: "contém", `=NOME` exato ou curinga (`MT_*`).
- `/upload` e `/extract-local` aceitam `?projeto=<código>` (a tela envia o projeto selecionado). Aba "Desenho" no modal Vinculação com "Testar" (layer + cor → MT/BT/descartada).
- Limite: o código do plugin (C#) não foi alterado; ao salvar a leitura pelo plugin, `cor_mt`, `cor_bt` e `regra` ainda não são reenviados por ele (ConfigLeitura.cs só conhece os campos antigos).

---

## 2026-11-16 — TASK-176: perfil de exportação por projeto (cabeçalho, centros, layout e arquivos)

- `exportacao` na configuração da vinculação: `cabecalho` (campos `chave`, `rotulo` do PDF, `tela`, `coluna` esq/meio/id, `obra`, `lembrar`), `centro` (`MATERIAL`/`MAO-DE-OBRA`), `pdf` (`cor_titulo`, `rodape`, `rotulo_ordem`, `rotulo_id`), `xlsx.colunas` (ordem) e `arquivos` (`pdf_sap`, `xlsx`, `rec`, `csv`). Sem perfil = Neoenergia: PDF, XLSX e visão SAP idênticos aos de antes (comparados com a baseline em `tests/dados/exportacao_padrao_176.json`).
- `gerar_pdf/gerar_pdf_unico/gerar_xlsx` aceitam `perfil`; `/api/orcamento/exportar-xlsx` e `/exportar-sap-pdf` usam o `perfil` do corpo (pré-visualização) ou o salvo para o `projeto`. A caixa "Informações do relatório SAP" (Resumo e Resultado), o CENTRO da visão SAP e os botões oferecidos seguem o perfil.
- O `.obra` grava as chaves do cabeçalho do perfil (`sap_limpo(sap, projeto)`); campo desconhecido continua descartado; arquivo antigo abre igual. `OBRA.*` = chaves marcadas `obra` do perfil (+ REGIAO e CONTEXTO); o validador da base aceita `OBRA.CONTRATO` nesse projeto.
- Limites: o formato do `.rec` e do CSV segue fixo (o perfil só liga/desliga a oferta); a tabela de `OBRA.REGIAO` por projeto fica para a TASK-177; fonte do PDF não é configurável.

---

## 2026-11-16 — TASK-177: pacote de projeto: exportar/importar tudo num arquivo e pontos de partida sem caso especial

- Pacote = `.zip` com `manifest.json` (`formato: "pacote-projeto"`, `versao_formato` 1, `versao_programa`, `projeto`, `partes` com hash) e um arquivo por parte: `base.csv`, `vinculacao`, `dicionario_cabos`, `postes`, `exportacao`, `desenho`, `leitor_processamento`, `leitor_classificacao`, `regional` (`<parte>.json`). Parte ausente = não mexe. Formato mais novo que o do programa é recusado com mensagem; hash diferente também.
- Importar: `previa` mostra por parte novo/substitui/igual, contagens (hoje → pacote) e os avisos do validador (`validar_config`, `validar_base`, regras do leitor), sem gravar; `aplicar` valida tudo antes e só então grava parte por parte (admin, com confirmação). A base de um projeto de origem passa a ser do projeto de destino. Importar num projeto não altera outro.
- Pontos de partida: `data/pacotes/<id>/manifest.json` (`projeto_codigo_padrao`, `nome_projeto`, `origem`, `seeds_antigos`, `partes` apontando para os arquivos; os de Pernambuco continuam em `data/*_pernambuco_seed.json`). `projetos.pacote_inicial` (campo novo no cadastro; Supabase: coluna `pacote_inicial` opcional) vale mais que o manifesto. Saíram `PROJETO_PERNAMBUCO`, `_seed_pernambuco`, `_atualizar_seed_pernambuco` (agora `_atualizar_seed_do_pacote`) e `_TABELAS` (a tabela regional vem da parte `regional` importada ou do pacote de partida; tabela `projeto_partes`).
- Limites: contextos/ajustes (ADR-006) e modelos de obra ainda não viajam no pacote; os casos de referência entram com a TASK-178.

---

## 2026-11-16 — TASK-178: casos de referência por projeto: conferência automática da base

- Caso = pedido (o mesmo payload do cálculo: `cabos` + `outros`) + `obra_cabecalho` + `esperado` (por operação · tipo de mão de obra · código) + `tolerancia` + `origem` (programa | ICP | manual) + `nota` + `historico`. Guardado em `casos_referencia` (SQLite; espelho opcional no Supabase). Não muda o cálculo.
- `conferir`: roda os casos em lote com a base carregada UMA vez e devolve igual / mudou (código a código: antes → depois, `novo`/`sumiu`/`qtd`) / erro, mais o relatório em Markdown no estilo de `CON_EMB_mudancas.md`. `POST /api/casos/conferir` aceita `base_csv` (base candidata, sem gravar) e `ids` (lotes com progresso na tela). "Mudança esperada" (`/api/casos/aceitar`) exige nota e grava quem/quando no histórico do caso.
- A prévia de importação de pacote com base mostra o impacto nos casos antes de gravar; os casos viajam no pacote (parte `casos`).
- Desempenho medido: 312 pedidos contra a base da semente (2.326 linhas) em ~5 s (≈17 ms por pedido, base carregada uma vez); 300 casos numa base pequena em 0,1 s.
- Limites: o caso guarda o pedido de cálculo (não o `.obra` inteiro nem a Totalizadora); gerar o esperado pela ICP fica por conta da habilidade `orcamento-icp-neoenergia-pe` (basta gravar o caso com `origem: "ICP"`).

---

## 2026-11-18 — Ocultar parâmetros ao sair da linha (Outros) + rótulo "manual" no balão

- Nova opção **Configurações › Aparência › Avançado › "Ocultar parâmetros ao sair da linha"** (desligada por padrão; `localStorage 'ocultar_parametros_ao_sair'`, classe `body.ocultar-parametros`). Ligada, o ativo de Outros que tem parâmetros mostra só o nome quando a linha perde o foco (`RL, PR(D,X), ATRINT` → `RL, PR, ATRINT`); ao entrar na linha o texto inteiro volta; passar o mouse mostra o balão com os parâmetros.
- Só aparência: o `<input>` mantém o texto inteiro (cálculo, cópia, filtros e `.obra` inalterados); um rótulo (`.ativo-compacto`) cobre o campo via CSS. `static/ativo_compacto.js` (`AtivoCompacto.compactar`, testado em Node): tira `(...)` colado ao nome; parêntese sem fechar (digitando) não é mexido.
- Balão de parâmetros: o que o usuário preencheu à mão passa a aparecer como **manual** (antes "do pedido") — `ativo_status.js`.
- Avaliadas as TASK-181 e TASK-182 (seções "Avaliação para implementação" nos arquivos de `.ai/tasks/`).
- Testes: `tests/test_ativo_compacto.py`.

---

## 2026-11-18 — CORREÇÃO: aspas das fórmulas viravam `&quot;` na tela da Base (regressão de 17/11)

- A tela da Base escapava as aspas da fórmula duas vezes (`fatorTela` e `escAttr`, esta introduzida na correção das aspas em 17/11): o campo mostrava `&quot;` literal e "Salvar Alterações"/exportar CSV gravavam `=SE(tipo=&quot;PA&quot;;2;3)`, que o importador recusa ("fórmula inválida"). Corrigido: `fatorTela` não escapa mais (só `escAttr`); o servidor repara `&quot;` → `"` em `normalizar_fator` (salvar/importar) e em `_fator_do_uso` (cálculo), então bases já gravadas assim voltam a funcionar. Testes em `tests/test_fator_formula_109.py`.

---

## 2026-11-17 — Resumo: comentário em linha, operação antes do ativo e aviso "ATIVO NÃO EXISTE NA BASE"

- **Comentário** (menu de contexto da linha, Cabos e Outros): a linha vira uma anotação (💬, verde itálico) — é uma linha **congelada** com outra aparência (`comentario: true` + `congelada: true`; o `.obra` normaliza `comentario` → `congelada`). Continua editável; "Congelar" não mexe em comentário; "Tirar comentário" devolve a linha ao normal. Um passo de Ctrl+Z.
- **Operação antes do ativo:** nas linhas em branco de Cabos e Outros a coluna Operação agora é um seletor (já vem com a operação herdada). Escolhida antes de digitar o ativo, vale para a linha criada; sem escolher, continua herdando.
- **Balão ao passar o mouse** (com "Colorir ativos por parâmetros" ligado): ativo pedido que não existe na base (nem como ativo/apelido, componente ou código do projeto/origem) aparece como **"<ATIVO> · ATIVO NÃO EXISTE NA BASE"**, em Cabos e Outros, antes do status de parâmetros. Substitui o título nativo do campo vermelho.

---

## 2026-11-17 — Cadastro rápido: renomear o ativo ao editar

- Em **Editar ativo** o campo **Nome** agora pode ser trocado. Ao publicar, as linhas do nome antigo saem e as do nome novo entram num passo só (local + nuvem); o nome antigo vira **apelido** por padrão (caixa "Manter “X” como apelido (recomendado)") para chamadas `:=` e orçamentos antigos continuarem achando o ativo. Sem a caixa, a tela lista os ativos que chamam o nome antigo e a confirmação avisa.
- Recusas: nome novo que já é de outro ativo do projeto (erro na revisão e no servidor), ativo antigo que não existe mais, apelido de OUTRO ativo (o apelido do próprio ativo em edição deixou de dar falso erro; e a edição sem renomear deixou de mostrar o falso aviso "já existe").
- API: `renomear_de` em `/api/admin/ativo/previa`, `/api/admin/ativo/publicar` e `/api/orcamento/rascunho` (`rascunho_ativo.avaliar(..., tirar=[nome antigo])`). Prévia devolve `renomeacao {de, para, existe_destino, origem_existe}`; o diff compara com as linhas do nome antigo; a auditoria registra `renomeado_de`.
- **Desfazer/refazer** a renomeação: `ativo_versoes` guarda, além das linhas, a lista `limpar` (formato `{"linhas": [...], "limpar": [...]}`; publicação comum continua lista simples). Desfazer apaga o nome novo e restaura o antigo; refazer faz o inverso, sem acumular linhas.
- Testes: `tests/test_publicar_ativo_148.py` (renomear, destino existente, origem inexistente, prévia, desfazer/refazer, nuvem), `tests/test_base_cadastro_147.py`.

---

## 2026-11-17 — Cálculo: ativo com declaração sem COMPONENTE não fica mais vazio

- Ativo cadastrado como `N3_2, N32, N23` (declaração SEM componente, itens abaixo com `COMPONENTE=N3_2`) calculava **nada**, sem erro nem aviso: o cálculo pegava o componente só da PRIMEIRA linha do ativo (a declaração, vazia). Agora usa o primeiro componente preenchido entre as linhas do ativo (`services/orcamento_calc.py`). Teste: `tests/test_declaracao_sem_componente.py`.

- **Origem do problema:** o cadastro rápido (Novo ativo) gerava a declaração só com apelidos (`N3_2, N32, N23`) sem COMPONENTE (o servidor só preenche o componente de ativos com parâmetros). Agora `gerarLinhas` já grava o componente nessa declaração (`static/base_cadastro_dados.js`), como nas declarações importadas (ex. `9400D, DT9/400`).

---

## 2026-11-17 — Base de Orçamento mais rápida: ativo novo aparece na hora, pesquisa ágil e barra de status

- **Causa da lentidão:** cada abertura da base (`GET /api/orcamento/dados`) baixava a tabela mestre INTEIRA da nuvem (páginas de 1000) e a regravava no SQLite. Agora `sync_tabela_master_se_preciso()` só baixa de novo depois de `SYNC_MASTER_TTL` (30 s) — o que o próprio servidor grava (publicar, salvar, importar) já atualiza o local; `invalidar_sync_master()` força. Duas telas abrindo juntas baixam uma vez só (trava). Falha de sync não marca o horário.
- **Ativo novo nas sugestões:** `static/base_sinal.js` (`BaseSinal.avisar()/ouvir()`, BroadcastChannel + `storage`, sem duplicar o aviso) — salvar/publicar/sincronizar na Base avisa o Resumo/Orçamento, que recarregam a base na hora (e mostram "Base de orçamento atualizada"). Pedido de recarga durante um download em andamento baixa de novo ao terminar.
- **Pesquisa:** o modelo da Base só é refeito quando a base ou a tabela da tela mudam (antes, a cada foco na barra, lendo todos os inputs); base velha (> 5 min) é usada na hora e atualizada em segundo plano; índice de pesquisa em maiúsculas montado uma vez por modelo (`indiceDe`); digitação espera ~80 ms.
- **Barra de status** (`static/base_status.js`, `BaseStatus.iniciar/rodar`): Carregando/Montando linhas → Validando e salvando → Salvo/Erro; Sincronizando com a nuvem; no cadastro rápido: Validando → Conferindo prévia → Publicando (base local e nuvem) → Recarregando → Publicado/Desfeito. Botões marcados `data-status-bloqueia` travam durante a operação.
- **Carga da tabela em lotes** de 300 linhas (a tela responde, a barra mostra o andamento; troca de projeto cancela a carga antiga) e valores com aspas/&/< nos campos da tabela agora são escapados (antes quebravam a linha).
- Medido: validar+preencher componentes de 6000 linhas leva ~0,02 s no servidor; por isso **não** foi feito "salvar só as linhas alteradas" (mudaria a semântica do salvar por projeto sem ganho).

---

## 2026-11-17 — Cabos e Outros: ativo fora da base na linha, e linhas congeladas/bloqueadas

- **Ativo fora da base:** o interruptor "Colorir ativos por parâmetros" (Configurações › Aparência › Avançado) passa a pintar de **vermelho** o ativo que não existe na base, em **Cabos e Outros**, ao vivo (e ao carregar a base). Mesma cascata do cálculo para o nome (ativo/apelido → componente → código, projeto e origem) em `static/ativo_na_base.js`; confere o texto digitado (o Validar/C1-BASE confere depois das regras de conversão). Linha congelada não é pintada.
- **Congelar** (menu de contexto da linha, em Cabos e Outros; vale para as linhas selecionadas): a linha continua na tela e **editável** (serve de comentário), mas sai da Totalizadora, do orçamento, da validação (lista, IA, C1-BASE), dos ajustes, das cores, do resumo da rede, dos contadores de qtd de ativos e do autocomplete; estrutura congelada sai do vínculo. **Cabo congelado** ainda pode ser fonte de um vínculo **feito à mão** (a estrutura continua recebendo `cabo`/`tipocabo`); os algoritmos (por poste, mapa, script) não ligam cabo congelado sozinhos, e ele não gera os itens `VINCULO` das regras de conversão.
- **Bloquear:** a linha não pode ser editada (nem por colagem/limpeza em bloco) nem excluída; continua nos cálculos, a menos que também esteja congelada. Linha bloqueada só aceita desbloquear. Congelar e bloquear valem os dois ao mesmo tempo; Ctrl+Z desfaz.
- Contrato do `.obra`/dados: campos **opcionais** `congelada` e `bloqueada` (booleanos) em cada linha de Cabos/Outros; arquivo antigo abre sem nenhuma marca. Servidor: `montar_totalizadora`, `itens_vinculo_para_totalizadora` e `vinculacao_regras` (Python = JS) respeitam as marcas.

---

## 2026-11-17 — Login: "Lembrar de mim" passa a valer

- Antes a caixa era enviada ao servidor e ignorada (o token valia 24 h para todos). Agora: **marcada** → sessão de **30 dias** (`JWT_LEMBRAR_HORAS`) e o programa abre direto, sem a tela de login; **desmarcada** → token de 24 h e a tela de login volta ao reabrir o programa.
- Cliente: `login.js` grava `auth_lembrar` (localStorage) e `sessao_ativa` (sessionStorage); `auth_fetch.js` descarta o token numa abertura nova quando `auth_lembrar === 'false'` e não há `sessao_ativa`. Sessão criada antes desta versão (sem a marca) segue entrando direto. Logout limpa as duas marcas. Tokens: `create_jwt_token(..., horas=)`.

---

## 2026-11-17 — Novo projeto: a nota do SAP é o nome do projeto

- Em **Novo projeto** (e **Salvar como**), **Colar SAP** passa a preencher o **nome** (e o nome do arquivo `.obra`) e o **número da obra** com a **nota**; o **título** do SAP não vira mais o nome: fica guardado em `sap.titulo` do projeto e continua sendo o título sugerido nos relatórios PDF (sem ele, vale o nome do projeto).
- Em projeto já existente (Abrir / Propriedades) o Colar SAP segue como antes. O título digitado ao gerar relatório, se diferente do nome, também fica no projeto. `projeto_obra.sap_limpo` aceita `titulo`; `.obra` antigo abre igual.

---

## 2026-11-16 — TASK-179: guia e modelo para quem cria uma base nova

- `docs/base_guia.md`: ativo simples → parâmetro → grupos com condição → linha viva (LV/LM) → retirada → cabos e arredondamento → composto → `OBRA.*` → armadilhas → conferência; cada recurso com um exemplo curto e o resultado.
- `data/pacotes/exemplo/`: base de 20 linhas (poste, N1, N3 com parâmetro `cabo`, N4 com linha viva/morta, cabos, alça, conector e um composto), vinculação simples e 5 casos de referência. Importa e funciona sozinho (testado: prévia sem avisos, 5 casos iguais, pedidos do guia dão o resultado dito). É ponto de partida **sugerido** (não é pacote de projeto nenhum): `GET /api/pacote/embutidos` lista e `/api/pacote/embutidos/<id>.zip` entrega no formato do pacote; a aba Pacote tem "Começar de um pacote pronto".
- Ficha da base (`orcamento.html`): coluna "O que faz" com a frase em português de cada linha ("vale quando cabo contém A35P; soma 3 de 7204 na instalação e na retirada"; grupos, compostos, linha viva e arredondamento).
- Validador: `codigo_do_achado`/`com_codigo` dão a cada tipo de erro/aviso um código `B-0xx` + `docs/base_guia.md`; aparece no fim das mensagens de `resumo_erros` (o texto da mensagem não mudou). Aviso novo B-013: código de MATERIAL cuja descrição é de equipamento (trafo, chave, religador, regulador, capacitor) nas linhas gerais de um ativo com parâmetro do mesmo tipo.
- Critério da base `base_orcamento_20261115_CON_EMB.csv` / `…_TT.csv`: esses arquivos não estão no repositório; o caso foi reproduzido com uma base sintética de CE_2_TR (trafo de 45 kVA nas linhas gerais → aviso B-013; trafos só nos grupos do parâmetro → sem aviso). Conferir com os arquivos reais quando estiverem à mão.

---

## 2026-11-15 — TASK-180: Montar Orçamento sempre recalcula; resultado gravado só volta ao abrir o projeto

- Substitui a restauração da TASK-154 (payload igual → restaurava o resultado do `.obra`, mesmo com base nova). Agora `montarOrcamento` sempre chama `/api/orcamento/calcular`; `assinatura` segue gravada no `.obra` (formato igual) mas não decide mais nada.
- Abrir um projeto (modo Abrir; menu ou plugin) com a tela do Resultado aberta mostra o resultado gravado, com as edições, sem recalcular (`aoAbrirProjeto` → `mostrarResultadoDoProjeto`); `.obra` sem resultado esvazia a tabela com aviso e para de acompanhá-la. Mesclar/Sobrescrever não mexem na tabela. Recarregar a página recalcula.
- Edição manual detectada em memória (`resultadoEditado`, MutationObserver, ignora `tr.fantasma`); ao recalcular o log avisa que as edições foram substituídas. Limitação: a flag se perde ao reiniciar o programa.

---

## 2026-11-15 — TASK-172: validador da base com menos alarme falso de sobreposição

- Padrões `%X%` × `%Y%` só "podem valer juntos" se um texto contém o outro (`%30KVA%` × `%10KVA%` não; `%A35%` × `%A35P%` sim). Ativo com mais de 5 avisos de sobreposição vira UM aviso resumo na importação e no modal (lista completa recolhida); `validar_base` não muda. Tarefa: `.ai/tasks/TASK-172-15-11-2026.md`.

---

## 2026-11-15 — TASK-171: estrutura fictícia do cabo sugere os componentes avulsos (alças, conectores)

- Sob o cabo com estrutura fictícia (modal Vinculação) aparece `N4 (fictícia) → 3 × ALCA(cabo=…)  [+ Outros]`: os itens dos grupos da base que chamam ativos, só do lado do cabo novo; `[+ Outros]` cria as linhas em Outros (`origemAvulso` no `.obra`; a quantidade vai na notação `3-ALCA(…)`). Python + JS; nada entra no cálculo sem o clique. Tarefa: `.ai/tasks/TASK-171-15-11-2026.md`.

---

## 2026-11-15 — TASK-170: vinculação — nível do fictício, derivação, cabos por vão e parâmetros automáticos

- Fictício herda o nível da regra de cabo que aceita a estrutura; `lados: DERIVACAO` (derivação obrigatória + tronco de 2 iguais); `cabos_vaos` (1 cabo por vão); `define` (parâmetros automáticos pela linha do poste/1º cabo, ex.: `tipo=RL`, variantes do ATRINT_EQ). Python + JS, modal com as colunas novas, regras 081 atualizadas. Tarefa: `.ai/tasks/TASK-170-15-11-2026.md`.

---

## 2026-11-15 — TASK-169: equipamento herda os cabos do poste (RL, CFU, para-raio, chave-faca)

- Tabela de Estruturas ganha `ocupa` (false = equipamento: não consome ponta nem capacidade), `herda: POSTE`, `nivel`, `ordem` (MAIOR_PRIMEIRO), `distintos`, `se_faltar`. O equipamento recebe `cabo`/`cabo2` dos cabos que chegam no poste (desenho), dos ligados às estruturas do poste e dos fictícios, numa 2ª passada de `_analisar_base`; `distintos`/`ordem` valem também para estruturas que ocupam. `vincular_por_poste`, scripts e `achados` ignoram o equipamento. Python + JS; modal Vinculação com as colunas novas e o equipamento em cinza. Config sem os campos novos = como antes. Tarefa: `.ai/tasks/TASK-169-15-11-2026.md`.

---

## 2026-11-15 — TASK-168: o programa sempre abre no projeto default

- O projeto aberto da execução anterior é esquecido ao abrir o programa (id de sessão em cookie `gridr_sessao`, só no desktop; `GridRProjeto.reiniciarNoDefault`). Recarregar/trocar de tela na mesma execução mantém o projeto. Tarefa: `.ai/tasks/TASK-168-15-11-2026.md`.

---

## 2026-11-15 — TASK-167: balão e cores dos parâmetros preenchidos × que faltam

- Passar o mouse no ativo (Demais Ativos e Totalizadora) mostra os parâmetros preenchidos (origem: pedido/vínculo/obra/padrão) e os que faltam. Opção escondida em Configurações › Aparência › Avançado pinta o ativo de verde (tudo preenchido) ou laranja (falta algum) só em Demais Ativos; desligada por padrão. `static/ativo_status.js`. Tarefa: `.ai/tasks/TASK-167-15-11-2026.md`.

---

## 2026-11-14 — TASK-166: vinculação por poste (obra do plugin abre já vinculada)

- `vincular_por_poste` (`services/vinculacao_regras.py` + espelho JS): cada estrutura do poste liga nos cabos que CHEGAM nele (ponta do vão `_geo` × `_poste`), respeitando as Tabelas de Cabos/Estruturas; B1MT = 2 cabos iguais, cada cabo ganha a ponta do poste. Escolhe e avisa quando sobra; prioridade pela ordem da Tabela de Estruturas; manual mantido (`vinculoOrigem`), automático refeito (`"poste"`).
- Roda ao abrir `.obra` com `cad` (`projeto_obra.ler` → `cad`), no botão do modal Vinculação › Script e em `/api/plugin/conferir`. Plugin C# não alterado (o `.obra` já leva a geometria). Tarefa: `.ai/tasks/TASK-166-14-11-2026.md`.

---

## 2026-11-14 — TASK-165: parâmetros de obra (`OBRA.*`) e contexto no `.obra`

- Parâmetro de ativo `OBRA.MUNICIPIO|BAIRRO|UTD|UTEP|REGIAO|CONTEXTO` é preenchido pelo programa (cabeçalho da obra, contexto, tabela regional `data/regionais_pernambuco.json` para `OBRA.REGIAO`), nunca pelo pedido (erro) e pulado na ligação posicional. Ordem: pedido → obra → padrão da declaração (+ aviso `OBRA.X não informado na obra; usando …`). `services/obra_params.py`; `resolver_pedido(..., obra)` Python/JS.
- Payload `POST /api/orcamento/calcular` ganha `obra` (opcional); plugin (`/conferir`) e Totalizadora também. `.obra` guarda `projeto_trabalho.contexto` (salvar/abrir/recarregar; Mesclar mantém). Formato da base não muda. Tarefa: `.ai/tasks/TASK-165-14-11-2026.md`.

---

## 2026-11-14 — TASK-164: visão SAP estilo Excel, cópia idêntica à do Excel e cópia em blocos

- `static/resultado_orcamento.html` (SAP): célula só edita com 2 cliques/F2/Enter/digitar (sem `contenteditable` fixo); 1 clique seleciona, arrastar/Shift/setas ampliam, coluna `#` (nº da linha, não copiável) seleciona a linha; Esc cancela edição; Ctrl+A.
- Cópia agora é idêntica à do Excel: texto TSV com CRLF em TODA linha (inclusive a última) + tabela HTML (`ClipboardItem`; evento `copy` usa `setData`). Antes era LF sem quebra final, o que fazia a aplicação de destino perder a coluna CENTRO. Formato escolhível (Excel / Só texto / LF) e colunas da cópia (padrão TIPO MDO, CÓDIGO, QUANTIDADE, CENTRO), guardados em `localStorage`.
- Contador de selecionadas/faltam e cópia em blocos de N (padrão 10, editável) com progresso, ◀ bloco anterior e Recomeçar. Helpers puros em `static/orcamento_sap.js` (`textoCopia`, `htmlCopia`, `andamentoBlocos`, `blocoAnterior`, `colunasDaCopia`). Tarefa: `.ai/tasks/TASK-164-14-11-2026.md`.

---

## 2026-11-13 — TASK-163: plugin do nanoCAD integrado ao projeto de trabalho; sem gravação no banco

- O plugin grava `<desenho>.obra` (formato do 📁 Projeto, com a seção `cad` dele) na pasta do desenho. `services/projeto_obra.montar` PRESERVA a `cad` ao gravar.
- `POST /api/plugin/receber {chave, caminho}` (só 127.0.0.1, só no programa instalado) e `--abrir-projeto <caminho>` no `.exe`: a janela vai para `/resumo?plugin_pendente=<id>` e abre o projeto por `GridRProjeto.abrirDoPlugin` (mesmo caminho do menu Abrir: Abrir/Mesclar/Sobrescrever; se já é o projeto aberto, relê do arquivo). `GET /api/plugin/pendente/<id>` devolve `{caminho, projeto}`.
- **Removido** `POST /api/plugin/obra` (TASK-106: o plugin gravava a obra no banco). Plugin antigo que só manda a obra: as linhas aparecem na tela, sem gravar. Plugin 0.10.0 (TASK-020). Tarefa: `.ai/tasks/TASK-163-13-11-2026.md`.

## 2026-11-13 — TASK-162: barra de pesquisa fixa na Base de Orçamento

- O Ctrl+K da Base de Orçamento deixou de abrir pop-up: a pesquisa é uma barra `sticky` no topo da tela (Ctrl+K leva o cursor para ela; resultados num painel abaixo, ↑↓/Enter/Esc). A "Consulta Rápida" flutuante e `BaseFicha.pesquisaRapida` foram removidas.

---

## 2026-11-13 — TASK-161: relatório SAP no formato do ICP (PDF único)

- "Relatórios PDF" gera um arquivo só (ODI e depois ODD, cada um em página nova com cabeçalho); páginas de continuação sem cabeçalho nem linha verde; 48 linhas na 1ª página; duas casas nas quantidades. A leitura de PDF SAP passou a usar a posição do texto (o título que o ICP escreve depois da tabela não vira descrição).

---

## 2026-11-13 — TASK-160: etiqueta com o estado do automático, fim do alarme falso e PDF com ODI+ODD juntos

- Etiqueta do projeto sempre mostra "Automático ligado/desligado"; o aviso de arquivo alterado é uma etiqueta à parte. A assinatura do arquivo guardada na tela podia ficar velha se a janela fechasse no meio de uma gravação (alarme falso): agora o disco ao abrir é o ponto de partida e o servidor local tolera a assinatura da própria última gravação. O leitor de PDF SAP lê o ODI e o ODD de um mesmo PDF (relatório do ICP); antes jogava o ODD como ODI.

---

## 2026-11-13 — TASK-159: atualização automática com barra de progresso e botão Cancelar

- O download do `.exe` roda em segundo plano (`iniciar_atualizacao`); `GET /api/update/status` dá fase, MB e percentual e `POST /api/update/cancel` cancela (apaga o parcial, não toca no `.exe`). Para se ficar 20 s sem receber dados e tem tempo máximo de 30 min. `POST /api/update/apply` agora só inicia.

---

## 2026-11-13 — TASK-158: Salvar / Salvar como / salvamento automático e rascunho do default

- Botão dividido "💾 Salvar ▾" + chave "Salvamento automático" + etiqueta de estado no Resumo. Projeto novo: automático ligado; aberto: desligado (só Salvar grava, inclusive o Resultado); default: ligado e grava rascunho (`rascunhos/` junto do app) recuperável. Salvar como cria cópia. Pergunta ao sair com pendências. Corrigida a corrida que gerava o alarme falso "alterado por fora" (assinatura sempre gravada sobre o estado atual).

---

## 2026-11-13 — TASK-157: Comparar com .obra / PDFs PB / PDF SAP e Importar PDFs com SAP

- Comparar abre a escolha de formato (REC, .obra, PDFs PB, PDF SAP). Novo `services/importar_sap_pdf.py` + `POST /api/importar-sap-pdf` lêem o PDF SAP (gerado aqui ou feito por fora): ODI → I, ODD → R. Importar PDFs tem o seletor PB | SAP. `.obra` usa só `projeto_trabalho.resultado`.

---

## 2026-11-13 — TASK-156: nome OrçaCore (logo mantido)

- GridR → **OrçaCore** nas telas, janela, README e manuais; arquivos/pastas usam `OrcaCore` (sem acento): `OrcaCore.exe`, `%APPDATA%\OrcaCore\janela`. A Release publica `OrcaCore.exe` + cópias `GridR.exe`/`Leitor_Projetos.exe` (cada uma com `.sha256`) para as versões antigas se atualizarem; `auto_updater` aceita os três nomes. `"app": "GridReader"` do plugin do nanoCAD e o ícone não mudam; o código do plugin não foi tocado.

---

## 2026-11-13 — TASK-155: botão Colar SAP (JSON do Atlas)

- `GridRProjeto.interpretarSap`/`colarSap` (`static/projeto.js`): aceita só `atlas-projeto-rede` v1; `odiCodigo`/`oddCodigo` viram os IDs de investimento; `titulo`/`nota` viram nome e número do projeto. Botão em Novo/Propriedades do projeto, na caixinha do SAP e no relatório PDF (avisa se a nota difere do projeto).

---

## 2026-11-13 — TASK-154: projeto de trabalho (`.obra` numa pasta, gravação automática)

- Menu 📁 Projeto (só no aplicativo desktop): Novo/Abrir/Propriedades/Voltar ao default. O `.obra` é o envelope do `.obra.json` + `projeto_trabalho` (nome, número de obra, cabeçalho SAP, resultado do orçamento). Gravação parcial e atômica (`services/projeto_desktop.py`), conflito por SHA-1, fila única em `static/projeto.js`. Número de obra, título e SAP vêm do projeto; REC pede confirmação do número; relatórios abrem na pasta do projeto. Abrir com o Resumo preenchido: Abrir / Mesclar / Sobrescrever. Default = funcionamento de antes.

---

## 2026-11-12 — TASK-153: usabilidade e design (telas sem internet, hierarquia de botões, menus no Resumo)

- Nenhuma tela busca mais nada fora do programa: Tailwind do login/admin virou `static/tailwind_local.css` (regenerar com `scripts/gerar_tailwind_local.js`) e as fontes vêm de `static/fontes/`; `static/tema.css` (fontes + foco do teclado) entra em todas as telas.
- Base de Orçamento e Resultado: o `.btn-action` do `styles.css` vazava (degradê) e deixava todos os botões laranja; anulado nas duas páginas. Um principal por área, exclusão em vermelho, ações de administrador agrupadas.
- Resumo: menus "Mais ▾" e "Exibição ▾" (mesmos ids); "Limpar Tudo" longe de "Salvar no PDF". Tela inicial, Modo autônomo, manual (escuro), Painel Admin e rodapé (`data-rodape`) ajustados. Nome e logo não mudaram.

---

## 2026-11-12 — TASK-152: nome GridR e ícone

- Programa passa a se chamar GridR (telas, janela, README/manuais) e ganha o ícone (`static/gridr.svg/.ico`, favicon em todas as páginas, `--icon` no build). Release publica `GridR.exe` e a cópia `Leitor_Projetos.exe` (versões antigas continuam atualizando); atualizador aceita os dois nomes; perfil da janela migra para `%APPDATA%\GridR`. O identificador `"app": "GridReader"` da rota do plugin não muda.

---

## 2026-11-12 — TASK-151: executável salva/preenche senhas no login

- `ligar_senhas_salvas` esperava `CoreWebView2` já pronto e desistia em silêncio; agora aguarda a inicialização (thread) e registra o resultado no log. Não testado em WebView2 real.

---

## 2026-11-12 — TASK-149/150: teclado na dica de parâmetros (Outros) e botão Inverter no layout do Resumo

- Outros: ↓/↑ escolhem o valor, Enter/Tab inserem com ", " e vão ao próximo parâmetro (último fecha o ")"); máscaras simples viram valores. Cabos não mudou.
- Resumo: botão "⇄ Inverter" (Demais ativos à esquerda/em cima, Cabos à direita/embaixo), lembrado no navegador.

---

## 2026-11-12 — TASK-148: publicar só o ativo, desfazer e conflito (parte 3)

- `publicar_ativo.py` + rotas `/api/admin/ativo/*`: troca só as linhas do ativo (nuvem primeiro, depois local), prévia com diff, um passo anterior (`ativo_versoes`), aviso de edição simultânea (409). UI: Publicar no passo 3, Editar/Desfazer na ficha.

---

## 2026-11-12 — TASK-147: cadastro rápido de ativos (parte 2)

- Assistente "+ Novo ativo" na Base de Orçamento (grade estilo Excel, auto-preenchimento por código, colar do Excel/combinações, validação ao vivo e no servidor, simulador do rascunho); novo `rascunho_ativo.py` e `/api/orcamento/rascunho`.

---

## 2026-11-12 — TASK-146: pesquisa (Ctrl+K) e ficha do ativo na Base de Orçamento

- Novos `base_ficha_dados.js` (modelo/pesquisa/ficha, puro) e `base_ficha_ui.js` (pesquisa, ficha com abas, simulador, Ctrl+clique). Consulta Rápida usa a mesma pesquisa. Parte 1 do plano (cadastro rápido e publicação por ativo virão depois).

---

---

## 2026-11-12 — TASK-144: no plugin, editar o ativo com o autocompletar da aplicação, excluir qualquer linha e criar linhas digitadas

- `GET /api/plugin/catalogo?projeto=` (login): a Base de Orçamento do projeto para o autocompletar — nomes válidos no projeto (+ descrição, origem), apelidos e os parâmetros de cada ativo (`services/ativo_parametros`, a mesma lógica da dica da aplicação); a base é filtrada pelo NOME do projeto.
- Plugin 0.9.0 (TASK-017): ativo editável no pop-up, Excluir/Restaurar em qualquer linha, + Cabo / + Outros sem desenho. Tarefa: `.ai/tasks/TASK-144-12-11-2026.md`.

## 2026-11-12 — TASK-143: "Abrir na aplicação" do plugin abre a obra no GridReader instalado (não mais no navegador)

- `services/programa_local.py`: o .exe registra `%APPDATA%\OrcamentoNanoCAD\programa.json` (exe + porta) e aceita `--abrir-obra <arquivo>`. `GET /api/plugin/desktop` (pública) e `POST /api/plugin/receber` (pública, só 127.0.0.1 e só no programa instalado) guardam a obra em memória e trazem a janela para a frente (`janela_desktop.mostrar`) em `/resumo?plugin_pendente=<id>`; o Resumo busca (`GET /api/plugin/pendente/<id>`, com login), grava pelo `POST /api/plugin/obra` e abre. Plugin 0.8.0. Tarefa: `.ai/tasks/TASK-143-12-11-2026.md`.

## 2026-11-12 — TASK-142: arquivos no executável, barra de modais por projeto e base recarregada

- Ponte de arquivos da janela desktop (`ApiArquivos` + `desktop_arquivos.js`): baixar/importar pelas caixas nativas; `ALLOW_DOWNLOADS`.
- Trocar de projeto no Resumo recarrega a barra de modais; fechar a Tabela de orçamento recarrega a base; sincronizar/importar a master atualiza a cópia pessoal do admin.

---

## 2026-11-11 — TASK-141: ordem padrão da visão SAP

- Sem coluna escolhida: M antes de N e código em ordem numérica crescente (ODI e ODD); CÓDIGO passa a ordenar como número. Vale para tela, PDFs e Excel.

---

## 2026-11-11 — TASK-140: relatórios ODI/ODD em PDF (formato da nota do SAP)

- Novo serviço `exportar_sap_pdf.py` + rota `/api/orcamento/exportar-sap-pdf` + botão/diálogo "Relatórios PDF" na visão SAP do Resultado do Orçamento. Fontes Carlito em `data/fontes/` (fallback da Calibri).

---

## 2026-11-11 — TASK-139: ajuste somar_linhas (juntar o mesmo ativo de várias linhas)

- Nova ação de ajuste em Outros que soma, na 1ª linha, os ativos escolhidos repetidos em linhas diferentes (mesma operação e, por padrão, mesma entidade). Painel de ajustes ganhou a opção.

---

## 2026-11-11 — TASK-138: sugestão de ativos com nome iniciado por número

- `ativo_sugestao.js`: 3+ dígitos sugerem (postes `11200D`); só dígitos casam pelo começo do nome; 1-2 dígitos seguem como quantidade.

---

## 2026-11-11 — TASK-137: importar base de orçamento em CSV ANSI (Windows-1252)

- `/api/orcamento/upload` e `/api/admin/upload-master` decodificam UTF-8 e, se falhar, Windows-1252 (CSV do Excel). Antes o arquivo em ANSI era recusado com erro de codec.

---

## 2026-11-11 — TASK-136: sugestão de ativos do projeto e ativo de outro projeto não vale

- `ativo_sugestao.js` + botão 💡 Sugestões; `orcamento_calc.py` sem fallback para outro projeto. Tarefa: `.ai/tasks/TASK-136-11-11-2026.md`.

## 2026-11-11 — TASK-135: nome GridReader (era ReadGrid) e sem "Extraia textos…"

- Nome trocado em telas, janela, tela "Iniciando…", confirmação de fechar e plugin ("GridReader para nanoCAD", 0.7.1 — sobe para a atualização automática levar o nome). Cabeçalho "Grid**Reader**". Frase "Extraia textos, fontes, cores e mais com precisão." removida da tela inicial. Perfil da janela em `%APPDATA%\GridReader\janela`, levando junto o `ReadGrid\janela` (senhas salvas). Tarefa: `.ai/tasks/TASK-135-11-11-2026.md`.

## 2026-11-11 — TASK-134: composto com parâmetros e grupo que chama ativos

- Nome com parênteses antes do `:=` na Base de Orçamento; cálculo por fila; validação e JS em paridade. Tarefa: `.ai/tasks/TASK-134-11-11-2026.md`.

## 2026-11-11 — TASK-133: fechar a janela pede confirmação e encerra o servidor; login com senhas salvas

- `services/janela_desktop.py`: `confirm_close` com texto em português ("Fechar o ReadGrid? O servidor local também será encerrado."); fechada a janela, o uvicorn para e o processo sai (`os._exit`). Perfil persistente da janela (`%APPDATA%\ReadGrid\janela`) e, no WebView2, salvar senhas + preenchimento automático ligados.
- Login: campos com `name` + `autocomplete`, credencial guardada no gerenciador de senhas ao entrar e preenchida ao abrir (Credential Management API, onde houver). Tarefa: `.ai/tasks/TASK-133-11-11-2026.md`.

## 2026-11-11 — TASK-132: nome ReadGrid e "Desenvolvido por Gabriel Sales"

- `static/marca.js`: cabeçalho com o nome (`data-marca="cabecalho"`) no login e na tela inicial e rodapé "ReadGrid · Desenvolvido por Gabriel Sales" nas telas principais. Títulos das páginas, janela do programa, tela "Iniciando…" e plugin ("ReadGrid para nanoCAD") com o nome novo. Sem logo de empresa (decisão do usuário). Arquivos `Leitor_Projetos.exe` e `OrcamentoNanoCAD.dll` mantidos (atualização automática). Tarefa: `.ai/tasks/TASK-132-11-11-2026.md`.

## 2026-11-11 — TASK-131: o programa traz o plugin do nanoCAD; projetos e configuração de leitura do plugin na aplicação

- `services/plugin_instalador.py` + `build.bat` (DLLs do nanoCAD 26 e 5): o .exe instala/atualiza o plugin em `%APPDATA%\OrcamentoNanoCAD\plugin\<nanoCAD26|nanoCAD5>` ao abrir (sem voltar versão; DLL em uso é renomeada). Botão **Plugin nanoCAD** (`static/plugin_nanocad.js`): versão trazida × instalada, caminho para o APPLOAD, instalar agora, abrir pasta; no online, links de download. Release do GitHub leva as duas DLLs.
- `GET /api/plugin/projetos`, `GET|PUT /api/plugin/leitura` (tabela `plugin_leitura_config`, local + Supabase): configuração de leitura do desenho por projeto; vai em `/api/plugin/regras` e na versão. `GET|POST /api/plugin/instalacao`, `POST /api/plugin/instalacao/abrir-pasta`.
- Plugin 0.7.0 (nanoCAD 5 pronto, projetos da aplicação na tela inicial). Tarefa: `.ai/tasks/TASK-131-11-11-2026.md`.

## 2026-11-10 — TASK-130: atualização automática do plugin do nanoCAD

- `GET /api/plugin/atualizacao` e `/api/plugin/atualizacao/arquivo` (públicas): a aplicação publica a DLL do repositório (`plugin_nanocad/appload/nanoCAD26/` + `versao.json` com SHA-256, gravado pelo `compilar.sh`); manifesto que não confere com a DLL → 503. `build.bat` empacota a DLL no executável. O plugin (0.6.0) baixa, confere e troca a DLL; vale ao reabrir o nanoCAD. Tarefa: `.ai/tasks/TASK-130-10-11-2026.md`.

## 2026-11-10 — TASK-129: ")" automático e descrição/filtro da base no REC e na comparação

- `completarParenteses` no Enter/blur dos campos de Outros; `/api/orcamento/detalhes` devolve filtro (base do cálculo, preferindo o projeto); resultado_orcamento usa `buscarDetalhesCodigos`. Tarefa: `.ai/tasks/TASK-129-10-11-2026.md`.

## 2026-11-10 — TASK-128: desktop — abertura sem erro, campos de chave, telas dentro do programa, dica do ativo

- Tela "Iniciando…" até o servidor responder; `campos_segredo.js`; `abrir_na_app.js` (Montar Orçamento e outras telas na própria janela); `ativo_dica.js` sem falso "não encontrado" e base pré-carregada. Tarefa: `.ai/tasks/TASK-128-10-11-2026.md`.

## 2026-11-10 — TASK-127: executável com auto atualização pelo GitHub

- Correção: o .bat da troca iniciava com janela de console visível e o `timeout` não esperava; agora `CREATE_NO_WINDOW` + `ping`.

- Workflow `release-exe.yml`, `build.bat` genérico, `scripts/release.py`, `services/auto_updater.py` reescrito (Releases + SHA-256) e `/api/update/apply` sem URL vinda do navegador. Tarefa: `.ai/tasks/TASK-127-10-11-2026.md`.

## 2026-11-10 — TASK-126: modal de vinculação lista as estruturas mesmo sem cabos

- `abaVinculo` não sai mais com "Nenhum cabo na tabela". Tarefa: `.ai/tasks/TASK-126-10-11-2026.md`.

## 2026-11-10 — TASK-125: modal de vinculação lista as linhas de Outros sem classificação

- `ehLinhaEstrutura` (JS e Python) inclui entidade `0`/vazia com ativo. Tarefa: `.ai/tasks/TASK-125-10-11-2026.md`.

## 2026-11-10 — TASK-124: poste com nome compacto como ativo preenche o parâmetro `poste`

- `eh_poste`/`poste_para_compacto` (Python e JS) reconhecem `DT11/600`, `R12/1000` e `121000R`. Antes `121000R, N1` deixava o `N1` com `D`. Tarefa: `.ai/tasks/TASK-124-10-11-2026.md`.

## 2026-11-10 — TASK-123: Totalizadora avalia nome e parâmetros separados; validação usa a base do cálculo

**Tipo:** correção · `static/ativo_avaliar.js` (novo), `static/resumo.js`, `static/resumo.html`, `routers/validacao.py`, `services/validacao_planilhas.py`, `app.py`

- **Totalizadora:** a edição do ATIVO/ORIGEM comparava o texto inteiro com a base (`FLABIT(D,35)` ficava vermelho, sem descrição). Agora `avaliarTot` (edição) e a geração avaliam o nome (ativo/apelido → componente → código, projeto e origem) e, se existe, os parâmetros: vermelho = nome inexistente; **âmbar** = parâmetros com problema (fora do escopo, não declarados, a mais), mensagem na descrição.
- **Validação:** passa a usar a mesma base do cálculo (oficial), não a cópia pessoal de "Salvar Alterações" (que dava "ativo não encontrado" para o que a tela via); `C1-PARAM` também para pedidos com **apelido**.

---

## 2026-11-09 — TASK-122: a descrição não resolve nome; validação separa nome e parâmetros; dica de parâmetros em Outros

**Tipo:** mudança de regra (RN-06) + funcionalidade · `services/orcamento_calc.py`, `services/validacao_planilhas.py`, `static/ativo_dica.js` (novo), `static/resumo.js`, `static/resumo.html`

- **Cascata RN-06:** agora ativo (e apelidos) → componente → código. Removidos "DESC_ATIVO exato" e "parcial" (a descrição só informa). Nomes que só existem como descrição (ex.: `FLABIT`) deixam de ser encontrados, com ou sem parâmetros.
- **Validação:** `C1-BASE` avalia só o nome (antes do `(`) e respeita origem/projeto; novo `C1-PARAM` cobre os parâmetros de um ativo que existe (valor não cadastrado, desconhecido, a mais/menos, "ativo sem parâmetros declarados").
- **Tabela Outros:** ao digitar `(` mostra os parâmetros aceitos do ativo (padrão, valores, parâmetro da posição atual destacado, clique insere), "ativo sem parâmetros declarados" ou "ativo não encontrado na base".
- Testes atualizados: `test_orcamento_lv.py` (passos de descrição viraram "descrição não resolve"; cabo no formato do payload), `test_plugin_cad_api.py` (neutro cadastrado como `[CAZ7,94]`, literal com vírgula).

---

## 2026-11-09 — TASK-121: rede adaptativa (proxy/direto/offline) + cópia local do banco (só desktop)

**Tipo:** funcionalidade · `services/rede_adaptativa.py`, `services/copia_local.py`, `routers/rede.py` (novos); `services/supabase_client.py`, `services/connectivity_monitor.py`, `app.py`, `database.py`, `routers/auth.py`, `routers/admin.py`, `static/resumo.js`

O programa testa o Supabase por proxy e por direto ao iniciar e com ele aberto, escolhe sozinho e, sem nuvem, entra em modo offline (disjuntor: `get_supabase()` devolve None; sem esperar os ~21 s de timeout). A cada início (segundo plano) e ao voltar a nuvem, o banco local é posto igual ao da nuvem (inclui contas), preservando o que foi feito offline (mescla de 3 vias, conflitos em `conflitos_locais`). Login offline usa o mesmo `user_id` do online (`usuarios_locais.nuvem_id`). Versão online (`APP_MODE=server`) sem mudança de comportamento (a única alteração compartilhada é gravar/usar `nuvem_id`, que só torna o fallback de login coerente).

---

## 2026-11-06 — TASK-120: base de orçamento mantém a ordem da planilha ao sincronizar com o Supabase

**Tipo:** correção · `services/sync_service.py` (`sync_tabela_master`, `migrar_projetos_compartilhados`, `get_merged_orcamento`)

Depois de **Importar Tabela Master** a tela mostrava as linhas embaralhadas (e ao salvar a validação reclamava de "declaração fora da ordem"). A sincronização baixava a tabela do Supabase **sem ordenar** (ordem física do Postgres, embaralhada após apagar/reinserir; a paginação ainda podia repetir/pular linhas). Agora pede por `id` (ordem de inserção = ordem da planilha), e as leituras locais também têm `ORDER BY id` explícito. Para voltar ao normal, reimportar a planilha uma vez.

---

## 2026-11-06 — TASK-118: Resumo da rede (comprimento vindo de outro cabo) + Importar JSON com diálogo

**Tipo:** ajuste/funcionalidade · `static/resumo.js`, `static/resumo.html`, `static/resumo_layout.js`, `routers/obras.py`

- Resumo da rede: cabo sem comprimento usa também o comprimento de **outro cabo** na linha de baixo (`P 50` + `CAZ9,5 abc 50 m` → 50 m na Média), além da linha só com o número (TASK-116).
- **Importar JSON** abre direto um diálogo Carregar / Adicionar / Subtrair (sem salvar a obra no servidor). Rota nova `POST /api/obras/ler-arquivo` (valida como o /importar, não grava). Modelo continua importado como modelo salvo.

---

## 2026-11-06 — TASK-117: Limpar Tudo nos modais + configuração da barra de modais por projeto

**Tipo:** funcionalidade · `routers/barra_modais.py` (novo), `static/barra_modais.js` (novo), `static/resumo.js`, `static/resumo.html`, `static/resumo_layout.css`, `static/gerador_modal.js`, `static/modal_montavel_ui.js`, `app.py`

- **Limpar Tudo** esvazia também Ramais, Conexões e os geradores de Postes/Cabos; confirmação detalhada e Ctrl+Z/Refazer (o snapshot ganha o campo extra `modais`, fora da comparação de estados).
- **Modais do projeto** (painel Configurações): ocultar/mostrar, ordenar, excluir/editar/criar montáveis e copiar montáveis entre projetos. API `/api/barra-modais` (GET todos, PUT/copiar só admin), guardada em `atalhos_modais` (chave `barra_modais`), sem SQL novo. Cada botão da barra ganhou `data-modal-id`.

---

## 2026-11-06 — TASK-116: Resumo da rede (neutro fora da Média; comprimento na linha de baixo)

**Tipo:** ajuste · `static/resumo.js` (`calcularResumoRede`, `extrairPrefixoCaboSemComprimento`), `static/anotacao_cabos.js` (`comprimentoSozinho`)

O neutro do cabo de média deixa de somar na Média (a Baixa não mudou). Cabo reconhecido sem comprimento na linha (notação nova ou formato antigo) usa o comprimento da linha seguinte quando ela traz só o número (`45`, `45 m`).

---

## 2026-11-05 — TASK-115: parâmetro usado em fórmula aceita qualquer valor (mesmo com grupos parciais)

**Tipo:** correção · `services/ativo_parametros.py` ⇄ `static/ativo_parametros.js` (`assinaturas_das_linhas`/`assinaturasDasLinhas`)

Base real `ATERINT` (`tipo` com grupos só para `MRT%` e `MRT_RDC`, e valores `PR/RDC/RPN` só nas fórmulas): `ATERINT(PR)` dava `TIPO=PR não cadastrado em ATERINT (existem: MRT_RDC, MRT%)`. A regra da TASK-114 (abrir só quando não há grupos) foi generalizada: parâmetro usado em fórmula de FATOR do ativo deixa de ser fechado, e o valor sem grupo recebe só as linhas comuns + fórmula. Parâmetro fora de fórmulas continua fechado.

---

## 2026-11-05 — TASK-114: parâmetro com padrão usado só em fórmula não exige grupo e aceita qualquer valor

**Tipo:** correção · `services/ativo_parametros.py` ⇄ `static/ativo_parametros.js` (`assinaturas_das_linhas`/`assinaturasDasLinhas`, `_formula_por_ativo`, `validar_base`)

Base real (`CERCA=CERCA(descr=5F,tipo=PA)` com `tipo` só nas fórmulas): a importação recusava (`padrão TIPO=PA … não tem nenhum grupo cadastrado`) e, mesmo importada, o pedido com `tipo=T`/`P` seria rejeitado (`não cadastrado (existem: PA)`). Agora um parâmetro declarado com padrão, usado em fórmula de FATOR do ativo e sem nenhum grupo/máscara que o restrinja, não exige grupo e fica **aberto** (o padrão é só o padrão). Parâmetros com grupos continuam fechados.

---

## 2026-11-05 — TASK-113: parâmetro livre (só na fórmula do fator) aceito no pedido

**Tipo:** correção/feature · `services/ativo_parametros.py` ⇄ `static/ativo_parametros.js` (`assinaturas_das_linhas`/`assinaturasDasLinhas`, `resolver_pedido`/`resolverPedido`), `services/orcamento_calc.py`

O cálculo recusava `3-CERCA(fios=8F)` quando `fios` só aparecia na fórmula do FATOR (`CERCA não tem variantes cadastradas` / `parâmetro FIOS desconhecido`). Agora os nomes usados nas fórmulas e não declarados viram **parâmetros livres** da assinatura do ativo (`sig["livres"]`): o pedido os informa **por nome** e eles chegam à fórmula (sem lista de valores; por posição não vale). Vale no cálculo, na Totalizadora (Python e JS em paridade) e para ativos sem nenhuma declaração.

---

## 2026-11-05 — TASK-112: parâmetro só na fórmula do fator vira aviso

**Tipo:** ajuste · `services/ativo_parametros.py` (`_validar_formulas`)

Fórmula do FATOR (`=NUM(fios)…`) que usa parâmetro que o ativo não declara deixou de recusar a importação/salvamento da Base: agora é **aviso**. Erro de sintaxe continua sendo erro. Os parâmetros declarados passam a contar os do projeto da linha e os das linhas genéricas.

---

## 2026-11-05 — TASK-111: novo layout da tela Resumo do Orçamento

**Tipo:** UI · `static/resumo.html`, `static/resumo_layout.css`/`.js` (novos), `static/resumo.js`, `static/grade_tabela.js`, `routers/obras.py`, `app.py`

Barra de título (nº da obra, Ler projeto, Montar orçamento, carregar/salvar/modelo, importar/exportar JSON), barra de modais ocultável, controles (visão, visibilidade e layout lado a lado/empilhadas das tabelas, resumo compacto) e dois painéis laterais ocultos por padrão e fixáveis (Configurações à esquerda, Assistente IA à direita). As tabelas ocupam a altura da tela e mostram **linhas vazias estilo Excel** (digitar ou colar cria as linhas). Regra de operação das linhas novas/coladas: `I` se a tabela está vazia, senão a da linha preenchida acima. Nova rota `POST /api/obras/exportar-atual` (JSON da obra da tela). Sem mudança de banco; os `id` do HTML foram mantidos.

---

## 2026-11-05 — TASK-110: ativo composto na Base (`N1C:=N1(C) + 2*N2(C)`)

**Tipo:** feature · `services/ativo_parametros.py` ⇄ `static/ativo_parametros.js` (`analisar_composto`, `_achar_composto`, `compostos_do_projeto`, `_validar_compostos`), `services/orcamento_calc.py` (`_expandir_compostos`)

Uma linha da Base com ATIVO `NOME:=item + qtd*item …` (sem CÓDIGO/COMPONENTE) define um ativo que agrupa outros: itens podem ser ativos simples, com parâmetros (`N1(C)`), `[nomes literais]` ou outros compostos, com quantidade por item (`2*N2(C)`, ponto decimal) e apelidos na mesma célula (`KIT:=…, K`). Pedir `3-N1C` no cálculo vale como `3-N1(C)` e `6-N2(C)` (a Totalizadora mostra `N1C`; o servidor expande). Composto por projeto (o do projeto vence o genérico); ciclo no cálculo avisa. A importação recusa composto com CÓDIGO/COMPONENTE, nome repetido ou igual a ativo existente, item que não existe na base ou com parâmetro inválido, quantidade ≤ 0 e ciclos.

---

## 2026-11-04 — TASK-109: FATOR da Base em função de parâmetro (fórmula)

**Tipo:** feature + migração · `services/fator_formula.py` (novo), `services/orcamento_calc.py`, `services/ativo_parametros.py` (`validar_base`), `database.py`, `routers/orcamento.py`, `routers/admin.py`, `services/sync_service.py`, `services/realtime_sync.py`, `static/orcamento.html`, `scripts/schema_supabase.sql`

FATOR I / FATOR R aceitam um número (como sempre) ou uma FÓRMULA começando com `=` que usa os parâmetros do pedido: `=NUM(fios)/8`, `=SE(tipo="A"; 2; 1)`. Avaliador próprio (sem `eval`): `+ - * /`, comparações, `NUM`, `ABS`, `MAX`, `MIN`, `ARRED`, `CIMA`, `BAIXO`, `SE`, `E`, `OU`, `NAO`; argumentos separados por `;`. A fórmula fica nas colunas novas `formula_i` / `formula_r` (o fator numérico vira 0); linhas sem fórmula não mudam. No cálculo cada uso da linha avalia a fórmula com os parâmetros do pedido (antes do ARRED da TASK-107); erro de fórmula avisa e a linha não soma. A importação recusa fórmula inválida ou com parâmetro que o ativo não declara. O administrador precisa rodar o `schema_supabase.sql` no Supabase.

---

## 2026-11-04 — TASK-108: nome literal entre colchetes ([S336,4]) na Base e nos pedidos

**Tipo:** feature · `services/ativo_parametros.py` ⇄ `static/ativo_parametros.js` (`literal_de`/`sem_colchetes`, `separar_apelidos`, `nome_base`, `analisar_ativo`, `dividir_tokens`), `services/anotacao_outros.py` ⇄ `static/anotacao_outros.js`, `services/orcamento_calc.py`, `services/autonomo/totalizadora.py`, `static/resumo.js`

`[nome]` é UM ativo simples cujo nome pode ter vírgula, espaço, hífen ou qualquer caractere (menos `(`, `)` e colchetes) — sem parâmetros. Na célula ATIVO da base (`[S336,4], 11336D` = nome + apelido) e nos pedidos de Outros (`2-[CAZ 7,94]`, `[S336,A], 3 N1`). O nome oficial é o conteúdo sem colchetes. Colchete mal formado/vazio/com parênteses = erro na validação da base. Sem colchetes, nada muda. O pedido mantém os colchetes até a busca (`sem_colchetes`), a Totalizadora casa `CAZ7,94` (neutro do cabo) com `[CAZ7,94]` e as regras aceitam `[ ]` em ATIVO_DE/ATIVO_PARA.

---

## 2026-11-04 — TASK-107: coluna ARRED (arredondamento por linha) na Base de Orçamento

**Tipo:** feature + migração · `services/orcamento_arred.py`, `services/orcamento_calc.py`, `database.py`, `routers/orcamento.py`, `routers/admin.py`, `services/sync_service.py`, `services/realtime_sync.py`, `static/orcamento.html`, `scripts/schema_supabase.sql`

Nova coluna `arred` (NORMAL / INTEIRO / CIMA / BAIXO; padrão NORMAL) em `tabela_orcamento` e `tabela_orcamento_master`: arredonda a quantidade que a linha gera A CADA USO (`qtd × fator`, antes da soma por código). NORMAL não arredonda cada uso (o total continua com 2 casas — resultado idêntico ao anterior). Dados existentes (SQLite e Supabase) viram NORMAL; a importação (local, master e tela) aceita a coluna `ARRED` (opcional; valor desconhecido recusa o arquivo sem alterar nada; `PARA CIMA`/`PARA BAIXO` também valem). O administrador precisa rodar o `schema_supabase.sql` no Supabase. Código direto (`9001`) segue com fatores e arredondamento neutros.

---

## 2026-11-02 — TASK-106: "Abrir na aplicação" (plugin nanoCAD)

**Tipo:** feature · `routers/plugin_cad.py`, `static/resumo.js`, `static/login.js`, `static/auth_fetch.js` · plugin: TASK-004 a 010 (`plugin_nanocad/.ai/`)

`POST /api/plugin/obra` grava a obra do desenho (particular, sem duplicar) e devolve o link do Resumo; o Resumo abre `?obra=` já carregado e o login volta
para o link. No plugin (0.3.0): MLEADER/guias, configuração de leitura, Ignorados, diagnóstico, seleção manual, navegação e visual da aplicação.

---

## 2026-11-02 — TASK-105: cálculo mostra o que o formato antigo de Outros descarta

**Tipo:** correção · `services/orcamento_calc.py` (`interpretar_outros`, `_descartes_formato_antigo`)

`1-U3-TM` virava os tokens `1 U3 TM` e o `-TM` era descartado em silêncio no cálculo (a camada 1 já acusava `C1-OUT-ORFAO`, mas **Montar Orçamento** seguia). Agora o cálculo coloca em `nao_encontrados` o token sem par, a quantidade não numérica e o `-` dentro do ativo (com a dica do compacto `U3-TM, 2 N1`); o que se entende (`U3`) continua calculado. Erros de formato de Outros (ex.: vírgula sem espaço) também passam a aparecer. Leitura e validação inalteradas. `test_orcamento_lv`: nomes com espaço passaram para a notação compacta (no formato antigo o 2º token era descartado).

---

## 2026-10-31 — TASK-104: API do plugin nanoCAD (regras e base ao vivo)

**Tipo:** feature · `routers/plugin_cad.py`, `app.py`, `tests/test_plugin_cad_api.py` · plugin: `plugin_nanocad/` (TASK-003 do plugin)

`/api/plugin/versao`, `/api/plugin/regras` e `/api/plugin/conferir`: o plugin usa as regras gravadas na aplicação, percebe mudanças em poucos segundos e
recebe, por linha da obra, os ativos que não estão na Base de Orçamento (mesmo caminho do modo autônomo: Totalizadora → cálculo).

---

## 2026-10-31 — TASK-103: subproduto plugin nanoCAD e vínculo preservado no `.obra.json`

**Tipo:** novo subproduto + contrato · `plugin_nanocad/` (com `.ai` próprio), `services/obras_arquivo.py`, `tests/test_plugin_nanocad_contrato.py`

Plugin C# para o nanoCAD (26 pronto; 5 compilado no Windows do usuário) que lê o desenho com as regras do projeto e grava
`<desenho>.obra.json`. Importar/exportar obra deixa de descartar `vinculoEstruturas`, `textoOriginal`, `_x/_y/_geo/_geoTipo/_poste`
e fictícios (validados campo a campo). Detalhes: `plugin_nanocad/.ai/`.

---

## 2026-10-28 — TASK-102: declaração com condições, apelido na declaração e cabo2

**Tipo:** regra de negócio · `static/ativo_parametros.js` ⇄ `services/ativo_parametros.py`, `services/orcamento_calc.py`, `static/vinculacao_regras.js` ⇄ `services/vinculacao_regras.py`, `data/vinculacao_pernambuco_seed.json`

A linha `N1=N1(poste=%D,tipocabo=A)` da Base de Orçamento passa a valer como declaração e 1º grupo; apelidos (`B1-MT`) podem acompanhar declaração/grupo; `cabo2` = descrição do 2º cabo.

---

## 2026-10-27 — TASK-101: `cabo` × `tipocabo` nos parâmetros da estrutura

**Tipo:** regra de negócio · `static/vinculacao_regras.js` ⇄ `services/vinculacao_regras.py`, `static/ativo_parametros.js` ⇄ `services/ativo_parametros.py`, `static/vinculacao_ui.js`, `data/vinculacao_pernambuco_seed.json`

`cabo` = descrição completa do cabo (sem o comprimento); `tipocabo` = VALOR DE CABO da Tabela de Cabos. Valores de parâmetro aceitam `-`. Causa do "padrão A": o código S04 casava a regra de alumínio.

---

## 2026-10-27 — TASK-100: cores por TIPO MDO no Resultado

**Tipo:** UI · `static/resultado_orcamento.html`

Mão de obra (N) em cinza escuro mais escuro e material (M) em cinza escuro mais claro, no REC e na SAP.

---

## 2026-10-27 — TASK-099: vínculo por coordenadas com reserva por sorteio

**Tipo:** feature · `static/vinculacao_regras.js` ⇄ `services/vinculacao_regras.py`, `static/vinculacao_ui.js`, `data/vinculacao_script_coordenadas_com_sorteio.json`

Passo `aleatorio` (sorteio reproduzível por semente, respeitando as tabelas) e opção `completar` nos scripts de vínculo.

---

## 2026-10-27 — TASK-098: tensão com ponto decimal no cabo (13.8KV)

**Tipo:** correção · `data/regras_leitor_processamento_pernambuco_seed.json`, `routers/regras_leitor.py`, `static/anotacao_cabos.js` ⇄ `services/anotacao_cabos.py`

`3XS1/0-13.8KV 35 M` agora é CABO no Leitor de Pernambuco e na notação do cabo (vírgula ou ponto na tensão). Seed antigo intocado é atualizado sozinho.

---

## 2026-10-27 — TASK-097: Outros com ativo de nome composto (CE3 PR)

**Tipo:** regra de negócio · `static/anotacao_outros.js` ⇄ `services/anotacao_outros.py`

Na notação compacta o único delimitador é ", "; `CE3 PR` é um item. Linha sozinha sem `qtd-ativo` é compacta (um item só).

---

## 2026-10-27 — TASK-096: parâmetros repetidos na estrutura (N4 com dois cabos)

**Tipo:** regra de negócio · `static/ativo_parametros.js` ⇄ `services/ativo_parametros.py`, `static/vinculacao_regras.js` ⇄ `services/vinculacao_regras.py`, `static/anotacao_outros.js` ⇄ `services/anotacao_outros.py`, `static/vinculacao_ui.js`, `static/resumo.js`, `services/autonomo/totalizadora.py`, `data/vinculacao_pernambuco_seed.json`

O mesmo nome de parâmetro pode se repetir (N4(cabo, cabo)); a estrutura recebe um valor por cabo vinculado (ativo do cabo ou tipo), por estrutura da linha. Valores aceitam `/`.

---

## 2026-10-27 — TASK-095: vínculo só com I/M e estruturas fictícias no cabo

**Tipo:** regra de negócio · `static/vinculacao_regras.js` ⇄ `services/vinculacao_regras.py`, `static/vinculacao_ui.js`, `static/resumo.js`, `services/autonomo/vinculacao.py`

Cabos e estruturas com operação R/*R saem do vínculo (análise, modal, script, Totalizadora). O cabo ganha `estruturasFicticias` (ocupam pontas, validadas pelas regras, viram `CABO_ESTRUTURA` na Totalizadora).

---

## 2026-10-27 — TASK-094: várias estruturas por linha de Outros, em blocos por poste

**Tipo:** feature · `static/vinculacao_regras.js` ⇄ `services/vinculacao_regras.py`, `static/vinculacao_ui.js`, `static/resumo.js`, `static/anotacao_outros.js`, `services/anotacao_outros.py`, `services/autonomo/*`

`9400D, N1, N4` agora é um poste e DUAS estruturas vinculáveis. Id da unidade = linha (1ª) ou linha + k/100; vínculos antigos continuam válidos. O modal mostra uma linha por poste com as estruturas lado a lado.

---

## 2026-10-27 — TASK-092/093: texto original do cabo e estrutura com fictício opcional

**Tipo:** feature · `static/vinculacao_regras.js` ⇄ `services/vinculacao_regras.py`, `static/vinculacao_ui.js`, `static/resumo.js`/`.html`, `services/autonomo/montagem.py`, `data/vinculacao_pernambuco_seed.json`

- 092: o cabo guarda `textoOriginal`; colunas Fases/Tensão/Nível (derivadas, editáveis, ocultáveis) e filtros `fases`/`nivel`/`padrao_texto` nas regras da Tabela de Cabos.
- 093: regra de estrutura com `qtd_min`/`qtd_max`/`max_reais` (SI3 = 1 real + 1 fictício opcional).

---

## 2026-10-27 — TASK-091: mapa do vínculo com rede MT/BT e postes (DXF)

**Tipo:** feature · `services/dxf_geometria.py` (novo), `services/dxf_service.py`, `routers/upload.py`, `static/vinculacao_ui.js`, `static/script.js`, `static/ler_projeto.js`, `static/resumo.js`

Leitura implícita (não vira linha na tabela) de LINE/POLYLINE em REDE_MT (vermelha)/REDE_BT (verde) — cor decide, layer confirma —, blocos com "poste" no nome e guias. Cada texto ganha `_geo/_geoTipo/_poste` por proximidade (4× a altura, seguindo a guia). O mapa mostra cabo = linha da rede, estrutura = poste, rede ao fundo opcional. Exige reextrair o DXF.

---

## 2026-10-27 — TASK-085…090: seis ajustes de usabilidade

**Tipo:** ajustes de UI · `static/resultado_orcamento.html`, `static/resumo.js`, `static/index.html`, `static/script.js`, `static/orcamento.html`, `static/base_grade.js` (novo, rota sem cache em `app.py`)

- 085: visão SAP com minimizar/trocar posição ODI×ODD e ajuste de coluna ao conteúdo.
- 086: linhas manuais em branco (sem `NOVO_COD` no REC).
- 087: Delete na tabela Outros apaga a linha certa (stopPropagation + limparGrade).
- 088: Entidade oculta na tabela Outros (menu Exibir); chip `cabo=C` virou tooltip no Ativo.
- 089: coluna Layer visível no Leitor.
- 090: Base de Orçamento estilo Excel, inserção acima/abaixo da selecionada, duplicar/excluir várias.

---

## 2026-10-08 — TASK-055: suporte real a OpenAI e Claude como provedor de IA, Haiku como padrão

**Tipo:** feature (integração de IA) · `services/validacao_ia.py`, `routers/validacao.py`,
`routers/ai_chat.py`, `models.py`, `static/resumo.html`, `static/resumo.js`, `requirements.txt`,
`.env.example` · `.ai/tasks/TASK-055-08-10-2026.md`

Depois de uma investigação de custo/arquitetura, usuário pediu suporte real a dois provedores de IA
novos — OpenAI e Claude (Anthropic) — nos dois lugares que até então só falavam com o Gemini: a
Camada 3 (validação/correção/ajustes de planilha por IA) e o chat geral. Pedido explícito: Claude
usando o modelo Haiku (`claude-haiku-5-5`) como padrão, por ser o mais rápido e barato da linha, com
Sonnet (`claude-sonnet-5-5`) como reserva de qualidade.

Investigação prévia (nesta mesma conversa) já tinha mapeado o ponto de injeção certo pra Camada 3:
`routers/validacao.py:_preparar_ia`, que hoje só sabia montar `chamar_gemini` e injetar numa função
`chamar(texto) -> str` que nem `executar_em_lotes` nem os validadores/corretores precisam saber de
qual provedor vem. `_preparar_ia` ficou provider-aware: lê `ai_provider` salvo, resolve a chave do
provedor certo (headers novos `X-OpenAI-Key`/`X-Anthropic-Key`, mesma precedência usuário > salva >
padrão do ambiente já usada pro Gemini) e escolhe entre `chamar_gemini`/`chamar_openai`/
`chamar_claude` — as duas novas em `services/validacao_ia.py`, mesmo contrato de `chamar_gemini`
(`chamar(api_key, modelos, texto, temperatura) -> str`, tentando os modelos da lista em ordem só
pra erro de cota/indisponibilidade). Decisão técnica: a função de chamada é resolvida pelo NOME via
`globals()` dentro de `_preparar_ia`, não por referência direta — necessário pra preservar a
testabilidade existente (os testes substituem `chamar_gemini` via `monkeypatch.setattr(rota,
"chamar_gemini", fake)`, que só funciona se a busca for dinâmica). O campo `modelo` do cabeçalho dos
prompts salvos (ex. `modelo: gemini-3.1-flash-lite`) só é usado como primeira tentativa quando o
provedor é Gemini — é sempre um ID Gemini, sem equivalente para OpenAI/Claude nos prompts salvos.

No chat geral, `routers/ai_chat.py:gemini_chat` ganhou branches reais de streaming pra `openai` e
pra `claude`/`anthropic`, com histórico e `system_instruction` aplicados do mesmo jeito que no
Gemini (antes, o branch de provedor diferente de Gemini simplesmente devolvia "Provider não
suportado"). `resolver_credencial`/`_chave_padrao` ganharam um parâmetro `provider` (compatível com
as chamadas existentes, que continuam usando o padrão "gemini") pra resolver a variável de ambiente
certa (`OPENAI_API_KEY`/`ANTHROPIC_API_KEY`). `GET /api/gemini/models` ganhou uma listagem fixa pra
Claude (Anthropic não tem um endpoint de listagem tão aberto quanto Gemini/OpenAI) e passou a usar o
header próprio `X-OpenAI-Key` pra OpenAI, em vez de reaproveitar a chave do Gemini como fazia antes
(esse reaproveitamento não tinha uso real, já que a UI só tinha a seção do Gemini).

UI: o modal único de configuração de IA (`static/resumo.html`/`resumo.js`) ganhou um seletor de
provedor com 3 botões e duas seções novas (OpenAI, Claude), reaproveitando o modal existente em vez
de criar um painel novo — ao trocar de provedor, aparece a seção certa (campo de chave + lista de
modelos), e o Claude já vem com o Haiku pré-selecionado. De passagem, corrigido um bug já
identificado: o botão Salvar gravava `ai_provider: 'gemini'` fixo, sempre, independente do que o
usuário tinha escolhido no seletor — agora grava o provedor realmente ativo.

`anthropic` adicionado ao `requirements.txt` (`openai` já estava); `.env.example` passou a
documentar `OPENAI_API_KEY`/`ANTHROPIC_API_KEY`.

Testes: 14 novos em `tests/test_validacao_ia.py` (`chamar_openai`/`chamar_claude` com fake de SDK —
sucesso, fallback só em erro de cota/indisponibilidade, erro sem gatilho não tenta outro modelo —
e dispatch de provedor em `_preparar_ia`/rota `/api/validacao/ia`: provedor salvo escolhe a função
certa, chave por header/salva/padrão do ambiente, `"anthropic"` tratado como `"claude"`, provedor
desconhecido cai pro Gemini) e parte dos 23 novos em `tests/test_ai_chat.py` (branches de chat
OpenAI/Claude em streaming, fallback, `/models` da Claude, UI estática do seletor e do bug
corrigido). Nenhum teste faz chamada de rede real — SDKs sempre mockados. Suíte completa (843
testes) sem regressão. **Sem chave de teste real de OpenAI/Anthropic disponível neste ambiente de
implementação** — a chamada real ponta-a-ponta ainda precisa ser confirmada pelo usuário com a
própria chave; validado com fakes de SDK e com um smoke test real-server + Playwright (servidor
uvicorn real, SDKs mockados em processo, banco SQLite temporário) confirmando a troca de provedor no
modal, o modelo padrão do Claude pré-selecionado e a gravação correta de `ai_provider` pelo
navegador de verdade.

---

## 2026-10-08 — TASK-054: chat sem o erro "Multiturn chat is not enabled", limpeza de modelos mortos

**Tipo:** fix (integração com o Gemini) · `routers/ai_chat.py`, `services/validacao_ia.py` ·
`.ai/tasks/TASK-054-08-10-2026.md`

Usuário mandou `"substitua o texto S3 por S3T na tabela outros"` no chat e recebeu `[ERRO] 400
INVALID_ARGUMENT. {'error': {'code': 400, 'message': 'Multiturn chat is not enabled for this
model', 'status': 'INVALID_ARGUMENT'}}`. Investigação: o chat (`routers/ai_chat.py`) usava a API de
**sessão com histórico** do Gemini (`client.aio.chats.create(..., history=...)` +
`chat.send_message_stream(...)`) — API que alguns modelos (sobretudo variantes "lite") recusam. O
fallback do código só reconhecia erro de cota/indisponibilidade (`429/quota/exhausted/not found/
404/unavailable`) pra tentar o próximo modelo da lista; esse erro não batia com nenhuma dessas
palavras, então o código desistia em vez de tentar o próximo modelo da lista.

Correção robusta (em vez de só ampliar palavras-chave de fallback): trocada a chamada de sessão por
uma chamada única de `generate_content_stream` (streaming, SEM sessão) com o histórico montado
manualmente dentro de `contents` — exatamente a mesma API de chamada única (`generate_content`,
sem streaming) que `chamar_gemini` (Camada 3) já usava e que nunca teve esse bug. Confirmado por
introspecção do SDK instalado (`pip show google-genai` → 2.28.0;
`inspect.signature(AsyncModels.generate_content_stream)`) que o método e a assinatura do plano
original estavam certos: uma coroutine que devolve um `AsyncIterator`, usada como `response = await
client.aio.models.generate_content_stream(...)` + `async for chunk in response:` — mesmo padrão já
usado com `send_message_stream`. De defesa em profundidade, os gatilhos de fallback (tanto no chat
quanto em `services/validacao_ia.py:_ERROS_DE_FALLBACK`) ganharam `invalid_argument`/`multiturn`,
pro caso de um modelo futuro ter outra restrição parecida.

No mesmo pedido, usuário apontou (e uma pesquisa nesta conversa confirmou) modelos mortos nas listas
de fallback: `gemini-1.5-flash`/`gemini-1.5-pro` desligados desde 2025 (404/model-not-found);
`gemini-2.5-flash` com desligamento anunciado pela Google pra 16/10/2026 — a poucos dias desta
implementação. As duas listas (`routers/ai_chat.py` preferências do chat e
`services/validacao_ia.py:MODELOS_RESERVA`, que já eram quase idênticas) ficaram iguais e limpas:
`["gemini-3.1-flash-lite", "gemini-3.6-flash", "gemini-3.5-flash"]`, mantendo `gemini-3.1-flash-lite`
como padrão (já era antes desta tarefa). **Sem chave de API Gemini real disponível neste ambiente de
implementação**, não foi possível confirmar esses IDs contra `GET /api/gemini/models` como pedido —
usuário precisa confirmar com a própria chave antes de considerar a lista definitiva.

Testes: 23 novos em `tests/test_ai_chat.py` (primeiro teste dedicado à rota `/api/gemini/chat` —
sucesso com `generate_content_stream` sem sessão, histórico enviado em `contents` na ordem certa,
fallback acionado pelo erro exato "Multiturn chat is not enabled", erro sem gatilho não tenta outro
modelo, lista de preferências sem as entradas mortas). Nenhum teste faz chamada de rede real — SDK
do Gemini sempre mockado. Suíte completa (843 testes) sem regressão. Smoke test real-server +
Playwright (servidor uvicorn real, SDK mockado em processo, banco SQLite temporário): a mensagem
exata do bug reportado, com histórico prévio de uma troca de mensagens, respondeu sem o erro de
multiturn.

---

## 2026-10-08 — TASK-053: motor de ajustes ganha a ação `criar_linha_derivada`

**Tipo:** feature (motor de ajustes) · `services/ajustes_planilhas.py`, `static/painel_ajustes.js`,
`data/manual_regras_e_ajustes.md` · `.ai/tasks/TASK-053-08-10-2026.md`

Usuário pediu pra ajustar a regra `"AJUSTE DO (2)"` (Cabos: `"(2)"` depois da bitola = "tem neutro,
bitola 2") pra que, quando a bitola da linha for igual à do neutro, junte o N na fase na mesma
linha; quando for diferente, limpe o `(2)` e crie uma linha nova com o neutro, usando o MESMO
comprimento capturado na linha original.

Investigação em `services/ajustes_planilhas.py` (leitura completa): juntar N na fase e limpar o
`(2)` sem criar linha já eram possíveis com `substituir` + regex/backreferences existentes — sem
precisar de código novo. Criar a linha nova com conteúdo DINÂMICO (dependente do que foi capturado
em cada linha) não era possível: a única ação que insere linha (`adicionar_linha`) usa texto
literal fixo, sem regex, insere uma vez só por execução do ajuste — não existia primitivo pra "pra
cada linha que bate um regex, derive uma linha-filha a partir dos grupos capturados dessa mesma
linha".

Nova ação `criar_linha_derivada` (família `EXECUTORES`, mesmo padrão de `adicionar_linha` —
`_novo_erro_c1` pro guard de Camada 1, `_alvos`/`_quando`/`_Estado.marcar` reaproveitados): pra cada
linha que casa com `de` (regex com grupos), insere uma linha nova logo depois (ou antes) dela, com
`ativo_novo` resolvido via `match.expand()` — backreferences do match DAQUELA linha, nunca de
outra. Itera um snapshot das linhas que já existiam antes da ação começar, pra uma linha recém-
criada nunca ser reavaliada contra `de` na mesma execução (sem risco de crescimento em cadeia).
Campos `operacao_nova`/`entidade_nova` reaproveitam nomes/validação já existentes no motor.
`static/painel_ajustes.js` ganhou a seção de campos do editor; `data/manual_regras_e_ajustes.md`
documenta a ação com a própria regra do pedido como exemplo. Deliberadamente não adicionada ao
prompt de correção por IA (TASK-025) — fora do escopo daquele prompt, pensado pra correções
simples, não pra gerar regex/linhas novas sem supervisão.

18 testes novos (`tests/test_ajustes_planilhas.py`), incluindo a regra completa dos 4 passos sobre
as três linhas de exemplo do pedido. Suíte completa (806 testes) sem regressão. Smoke test
real-server + Playwright: editor renderizado sem erro com os campos da ação nova, pré-visualização
confirmando os três resultados exatos do pedido (`"CA 2(2) ABC 35m"` → `"CA 2 ABCN 35 m"`;
`"CA 1/0(2) ABC 35M"` → `"CA 1/0 ABC 35 m"` + nova linha `"CA 2 N 35 m"`; `"CA 2(2) AN 35M"` →
`"CA 2 AN 35 m"`).

---

## 2026-10-08 — TASK-052: base de orçamento totalmente separada por projeto

**Tipo:** feature (dados + admin) · `services/sync_service.py`, `routers/orcamento.py`,
`routers/admin.py`, `models.py`, `static/orcamento.html`, `scripts/schema_supabase.sql` ·
`.ai/tasks/TASK-052-08-10-2026.md`

Usuário relatou: "da forma como está estruturada eu posso colocar uma regra de um projeto estando
dentro de outro" e pediu uma estratégia pra editar a base técnica (`tabela_orcamento_master`)
escopada por projeto. Pedido de análise antes de implementar: investigação do CSV real de produção
(2.646 linhas, exportado pelo usuário do Supabase — o sandbox não tinha acesso de rede pra ler
direto) revelou que `projeto` já suporta mais de um valor por linha, separado por `/` (ex.
`"PARAIBA/PARAIBANOVO"`, 108 linhas reais), lido por `processar_calculo`. Pedido final do usuário,
depois desse achado: "projetos totalmente manipuláveis de forma separada... editando inclusive as
linhas compartilhadas".

Resolvido com uma migração única (dado, não schema): toda linha com `/` em `projeto` é dividida em
uma cópia exclusiva por projeto da lista, valores idênticos no momento da migração — local e
Supabase, cada um migrado a partir da própria consulta (os `id`s de um não correspondem aos do
outro). Depois da migração não existe mais "linha compartilhada": `GET /api/orcamento/dados` ganhou
os parâmetros opcionais `projeto` (igualdade exata), `genericas` e `nao_reconhecido`, usados pela
tela `/orcamento` (seletor de projeto + abas, reaproveitando o padrão de `#select-projeto` de
`resumo.html`); `POST /api/admin/upload-master` e `/sync-master-all` ganharam um `projeto` opcional
que escopa o `DELETE` (antes sempre apagava a tabela inteira, todos os projetos, local e Supabase)
e valida que toda linha do CSV/payload pertence ao projeto selecionado; `POST /api/orcamento/salvar`
(cópia pessoal do usuário) recebeu o mesmo tratamento — achado durante a implementação que esse
endpoint tinha o mesmo risco (DELETE sem filtro de projeto), fora do previsto no desenho original.
Campo Projeto da tela fica travado (`disabled`) no projeto da aba atual, exceto na aba "Projeto não
reconhecido" (pensada pra corrigir o valor). Modo "todos os projetos" continua existindo, ação
separada, com aviso reforçado. Normalização de projeto extraída uma única vez
(`services/sync_service.py`), reaproveitada em vez de duplicada pela terceira vez.

18 testes novos (`tests/test_orcamento_projeto.py`), incluindo Supabase simulado para a migração e
para o DELETE escopado. Suíte completa sem regressão. Smoke test real-server + Playwright: populou
as 3 categorias (exclusiva, genérica, não reconhecida) e uma linha compartilhada, confirmou a visão
por projeto, o campo travado, a migração ao vivo (a linha compartilhada virou duas cópias, cada
uma só na visão do seu projeto) e que importar um CSV escopado em PARAIBA não afeta RONDONIA nem a
cópia de PARAIBANOVO.

---

## 2026-10-08 — TASK-051: checagem obrigatória (contrato + ativo não encontrado) ao Ajustar/Montar Orçamento

**Tipo:** feature (validação) · `static/resumo.js` · `.ai/tasks/TASK-051-07-10-2026.md`

Usuário pediu que erros na tabela Outros fossem sinalizados pra evitar erros em grandes obras.
Depois de descartar a opção de unificar os 4 parsers de Outros (risco alto, muitos pontos de
regressão), desenhamos destaque ao vivo na tabela — mas o usuário mudou o modelo de disparo: "quero
que esse tipo de sinalização seja disparada quando clicar em ajustar, validar ou montar orçamento.
Exiba uma mensagem e não prossiga devido às inconsistências." Confirmado: bloqueia mas permite
"continuar mesmo assim" (mesmo padrão que Montar Orçamento já tinha, opcionalmente); roda sempre,
automaticamente, sem precisar ligar nenhuma preferência.

Investigação mostrou que **Validar** já cobria os dois sinais pedidos por padrão (Camada 1 de
contrato + `validar_base`/ativo não encontrado, via `MODOS_PADRAO.det = true`) — zero mudança
necessária ali. Faltava: (1) Ajustar não validava nada antes de abrir a gaveta; (2) Montar Orçamento
só validava se a preferência "Validar ao montar orçamento" (desligada por padrão) estivesse ligada.

Implementado com reaproveitamento total da lógica existente, sem duplicar nada: `executarValidacao`
ganhou um parâmetro opcional `somenteContrato` que, quando `true`, ignora a preferência salva do
usuário e força `{det:true, dominio:false, ia:false, pularIa:true}` — mesma montagem de payload,
mesma chamada a `/api/validacao/planilhas`, mesmo formato de retorno de sempre, só sem regras de
domínio nem IA. Nova função `checarInconsistenciasBasicas()` chama `executarValidacao(true)` e, se
houver achado de erro/aviso, usa o mesmo `cicloValidacao(res, true)` que o Montar Orçamento opcional
já usava pra mostrar o painel e perguntar "continuar mesmo assim" — nenhuma UI nova. `btnAjustar`
agora chama essa checagem antes de abrir a gaveta; `btnMontarOrcamento` passou a chamá-la sempre,
antes da checagem opcional mais completa (que continua intacta, pra quem já liga "Validar ao montar
orçamento" e quer domínio/IA também). `btnValidar` não mudou.

Suíte completa (770 testes) sem regressão — nenhuma mudança de backend. Smoke test real-server +
Playwright: linha de Outros com quantidade não numérica bloqueou Ajustar e Montar Orçamento (mesmo
com a preferência desligada) com a mensagem certa no painel; "continuar mesmo assim" seguiu
normalmente; linha com ativo não encontrado na base técnica também bloqueou; linha sem nenhuma
inconsistência abriu a gaveta/montou o orçamento direto, sem painel.

---

## 2026-10-07 — TASK-050: quantidade decimal com vírgula em Outros não aparecia na Totalizadora

**Tipo:** correção de bug · `static/resumo.js`, `services/autonomo/totalizadora.py` ·
`.ai/tasks/TASK-050-07-10-2026.md`

Usuário relatou: "na tabela outros não é possível utilizar uma quantidade quebrada como, por
exemplo 0,7-M335". Investigado: o projeto tem 4 implementações independentes do parser de Outros
(`<qtd>-<ativo>`). Três já aceitavam vírgula decimal (`orcamento_calc.py`,
`resumo.js::extrairParesQtdAtivoOutros`, o motor de Ajustes). A quarta — `syncTotalizadora`
(`static/resumo.js`, monta a Totalizadora/payload de cálculo) e seu porte fiel documentado no modo
autônomo (`services/autonomo/totalizadora.py`) — não: o regex só reconhecia `.` como separador
decimal, e mesmo reconhecendo, `parseFloat`/`js_parse_float` também não leem vírgula sem
`.replace(',', '.')` antes. Efeito: `"0,7-M335"` não casava o regex, caía no fallback que trata a
linha inteira como nome do ativo com quantidade 1 — sem erro visível, silenciosamente errado.

Corrigido nos dois lugares em paralelo (JS e seu porte Python, que precisam mudar juntos pra manter
a paridade testada em `tests/test_autonomo_totalizadora.py` contra o JS real via Node/QuickJS):
regex passou a aceitar `[.,]`, e o valor captado troca vírgula por ponto antes de virar número.
Mudança 100% aditiva — nenhum valor que já funcionava (inteiro, com ponto) muda de comportamento.

2 testes novos (`tests/test_autonomo_totalizadora.py`), mais um caso decimal adicionado ao conjunto
de dados do teste de fuzz existente (roda contra o JS real). Suíte completa (770 testes) sem
regressão. Smoke test real-server + Playwright na tela de verdade: digitou `"0,7-M335 *0,5-TR3
2-U4"` na célula Ativo da tabela Outros, confirmou via DOM que a Totalizadora mostra `M335` qtd
`0.7`, `TR3` qtd `-0.5`, `U4` qtd `2`.

---

## 2026-10-06 — TASK-049: `adicionar_ativo` — item da lista com fator próprio sobre a qtd dinâmica

**Tipo:** feature (motor de Ajustes) · `services/ajustes_planilhas.py`, `static/painel_ajustes.js` ·
`.ai/tasks/TASK-049-06-10-2026.md`

Usuário tentou combinar a sintaxe `código:qtd` (TASK-048) com "mesma quantidade de outro ativo"
(U4, fator 2), esperando "duas vezes a quantidade de U4, um código negativo e outro positivo" — não
funcionou, porque `qtd` própria SUBSTITUI a compartilhada por um valor literal fixo (comportamento
correto da TASK-048, mas não servia aqui). Perguntado "não podemos adicionar a lógica do fator?" —
confirmado.

Implementado: um item da lista agora também pode ser `{"ativo": código, "fator"?: número}` —
MULTIPLICA a mesma base da `qtd` compartilhada (a soma, se `qtd` for dinâmico; o próprio número, se
for literal) em vez de substituí-la por um valor fixo. `"qtd"` e `"fator"` no mesmo item são
mutuamente exclusivos. Nova `_qtd_base()` extrai a magnitude não escalada (sem fator), reaproveitada
tanto por `_qtd_dinamica()` (TASK-045, refatorada para usá-la, comportamento idêntico) quanto pelo
novo caso de item com fator próprio. UI: sintaxe `código:xN` no mesmo campo "códigos".

Exemplo exato do pedido: `{"acao": "adicionar_ativo", "tabela": "outros", "ativo": [{"ativo":
"90277", "fator": -2}, {"ativo": "90279", "fator": 2}], "qtd": {"soma": "U4"}, "quando": {"tem":
"U4"}}` — com `2-U4` na linha, gera `*4-90277` e `4-90279`.

6 testes novos (`tests/test_ajustes_planilhas.py`). Suíte completa sem regressão. Smoke test
real-server + Playwright confirma o fluxo completo pela UI (sintaxe `código:xN`, modelo em memória
e frase ao vivo corretos).

---

## 2026-10-06 — TASK-048: `adicionar_ativo` — item da lista com quantidade própria

**Tipo:** feature (motor de Ajustes) · `services/ajustes_planilhas.py`, `static/painel_ajustes.js` ·
`.ai/tasks/TASK-048-06-10-2026.md`

Logo após a TASK-047, usuário perguntou se a lista aceitava mais de um fator — respondido que, como
implementada, a quantidade era compartilhada por toda a lista. Usuário confirmou o pedido real:
"quantidades diferentes por código pra diminuir o número de regras".

Implementado: um item da lista de `ativo` agora pode ser `{"ativo": código, "qtd"?: número}` em vez
de uma string simples — tem a PRÓPRIA quantidade, substituindo a `qtd` compartilhada da ação só para
aquele código. Itens string continuam usando a `qtd` compartilhada (comportamento da TASK-047,
inalterado) — os dois formatos podem ser MISTURADOS na mesma lista. UI: o campo "códigos" ganhou a
sintaxe `código:qtd` (ex.: `90525:-1, 90542:-2, 92540`), sem precisar de um seletor de modo novo.

Exemplo exato do pedido: `{"acao": "adicionar_ativo", "tabela": "outros", "ativo": [{"ativo":
"90525", "qtd": -1}, {"ativo": "90542", "qtd": -2}]}`.

7 testes novos (`tests/test_ajustes_planilhas.py`). Suíte completa sem regressão. Smoke test
real-server + Playwright confirma o fluxo completo pela UI (sintaxe `código:qtd`, modelo em memória
e frase ao vivo corretos).

---

## 2026-10-06 — TASK-047: `adicionar_ativo` com lista de ativos fixos (mesma quantidade pra todos)

**Tipo:** feature (motor de Ajustes) · `services/ajustes_planilhas.py`, `static/painel_ajustes.js` ·
`.ai/tasks/TASK-047-06-10-2026.md`

Usuário perguntou se dava pra passar uma lista (ex.: `90525, 90542, 92540`) numa única regra. Já era
possível adicionar vários ativos no mesmo ajuste usando várias ações `adicionar_ativo` (o array
`acoes` já suportava isso desde a TASK-023), mas o usuário queria algo mais compacto. Confirmado:
"quantidade igual pra todos" antes de implementar.

Implementado: `ativo` passa a aceitar também uma LISTA de códigos fixos — um token novo por código,
todos com a MESMA `qtd` (fixa ou dinâmica via `{"soma": ...}` da TASK-045, resolvida uma única vez
antes de qualquer token ser adicionado). Extraída `_validar_qtd_adicionar()` para eliminar a
duplicação de validação entre o caso de ativo único e o de lista. Diferente do ativo dinâmico da
TASK-046 (que não combina com `qtd` dinâmico), lista SIM combina — os dois recursos são
independentes.

Exemplo exato do pedido: `{"acao": "adicionar_ativo", "tabela": "outros", "ativo": ["90525",
"90542", "92540"], "qtd": -1}`.

7 testes novos (`tests/test_ajustes_planilhas.py`). Suíte completa (755 testes) sem regressão.
Smoke test real-server + Playwright confirma o fluxo completo pela UI (alternar "ativo" para "lista
de códigos", preencher códigos separados por vírgula, modelo em memória e frase ao vivo corretos).

---

## 2026-10-06 — TASK-046: `adicionar_ativo` com ativo dinâmico (nome = código encontrado na própria linha)

**Tipo:** feature (motor de Ajustes) · `services/ajustes_planilhas.py`, `static/painel_ajustes.js` ·
`.ai/tasks/TASK-046-06-10-2026.md`

Usuário pediu, logo após a TASK-045: "caso tenha um trafo mono TR1* adicione com a mesma quantidade
negativa o texto TR1* acrescido com VP no fim desse texto" — o nome do ativo adicionado precisa ser
o CÓDIGO especificamente encontrado na linha (ex.: `TR110`, `TR115`...), não um literal fixo
`"TR1*VP"` (curinga de busca não é um código válido).

Implementado: `ativo` passa a aceitar também `{"igual_a": SELETOR, "prefixo"?: texto, "sufixo"?:
texto}` — para CADA item da própria linha que casa com `SELETOR`, gera um token novo
`"<prefixo><código do item><sufixo>"`, com quantidade = quantidade DAQUELE item × `qtd` (aqui `qtd`
é o FATOR por item, não soma — não combina com `{"soma": ...}` da TASK-045). Reaproveitados os
mesmos `_contexto()`/`_casa()` da TASK-045, nenhuma mudança em `regras_dominio.py`.
`_aplicar_um_ativo()` extraída de `_t_adicionar_ativo` (merge/append de um token) para ser chamada
em loop, uma vez por item casado. **Guard de idempotência:** como `igual_a` tipicamente é um
curinga amplo (`"TR1*"`) que também casaria com o próprio token recém-criado (`"TR110VP"` começa
com `"TR1"`), um item cujo nome já carrega o `prefixo`/`sufixo` configurado é excluído dos
candidatos — sem isso, reaplicar a ação geraria `"TR110VPVPVP..."` indefinidamente.

Exemplo exato do pedido: `{"acao": "adicionar_ativo", "tabela": "outros", "ativo": {"igual_a":
"TR1*", "sufixo": "VP"}, "qtd": -1, "quando": {"tem": "TR1*"}}`.

8 testes novos (`tests/test_ajustes_planilhas.py`). Suíte completa (748 testes) sem regressão.
Smoke test real-server (uvicorn em processo, `database.DB_PATH` apontado para um SQLite temporário
antes do `init_db()`, nunca `banco_resumo.db`) + Playwright confirma o fluxo completo pela UI
(alternar "ativo" para dinâmico, preencher seletor/prefixo/sufixo/fator, modelo em memória e frase
ao vivo corretos). Mesmo achado colateral da TASK-045 reencontrado (guard `C1-OUT-NUM` com código
hifenado + token negativo) — teste ajustado para evitar, fora do escopo.

---

## 2026-10-06 — TASK-045: `adicionar_ativo` com quantidade dinâmica (mesma quantidade de outro ativo)

**Tipo:** feature (motor de Ajustes) · `services/ajustes_planilhas.py`, `static/painel_ajustes.js` ·
`.ai/tasks/TASK-045-06-10-2026.md`

Usuário, cadastrando uma regra dentro de um contexto (TASK-044), pediu: "se tem U3 adicione em mesma
quantidade 90277 negativo" — a quantidade de "90277" precisa igualar a quantidade de U3 encontrada
NAQUELA linha, que varia (`1-U3`, `2-U3`, `3-U3`...), não um número fixo. Confirmado com o usuário
que U3 realmente varia (não é sempre 1) antes de implementar, já que `adicionar_ativo.qtd` só
aceitava um número literal.

Implementado: `qtd` passa a aceitar também `{"soma": SELETOR, "fator"?: número}` — soma das
quantidades dos itens da PRÓPRIA linha que casam com `SELETOR` (mesma linguagem de seletor de
sempre: código, curinga, `@GRUPO`, lista, regex), multiplicada por `fator` (padrão 1; `-1` gera o
prefixo `"*"` automaticamente, mesma regra de sinal da TASK-038). Reaproveitado:
`services/ajustes_planilhas.py::_contexto()` (já construía um `_Linha` a partir do texto de
qualquer linha, usado por `_quando()`) e o mesmo `_casa()` que `_cond()`'s `soma` já usa para
validação — nenhuma mudança em `regras_dominio.py`. Soma zero (seletor não casa na linha) é no-op,
não erro. `descrever_acao()` e a UI da gaveta (`painel_ajustes.js`) ganharam suporte ao novo modo
("quantidade": fixa/dinâmica, com campos "soma de"/"fator").

Exemplo exato do pedido: `{"acao": "adicionar_ativo", "ativo": "90277", "qtd": {"soma": "U3",
"fator": -1}, "quando": {"tem": "U3"}}`.

8 testes novos (`tests/test_ajustes_planilhas.py`). Suíte completa (740 testes) sem regressão.
Smoke test real-server + Playwright confirma o fluxo completo pela UI (trocar para quantidade
dinâmica, preencher seletor+fator, modelo em memória correto). Achado colateral, não corrigido
(fora do escopo): um guard de contrato pré-existente (`C1-OUT-NUM`) trata mal códigos hifenados
(ex.: `SUP-L`) combinados a um token negativo `"*"` — reproduz igual com `qtd` fixo, não é algo
introduzido por esta tarefa.

---

## 2026-10-06 — TASK-044: contextos dentro de um projeto (terceira camada de Ajustes)

**Tipo:** feature (arquitetura de dados + UI) · `database.py`, `scripts/schema_supabase.sql`,
`services/contextos.py`, `services/repo_json.py`, `routers/validacao_ajustes.py`,
`static/resumo.html`, `static/resumo.js`, `static/painel_ajustes.js` ·
`.ai/tasks/TASK-044-06-10-2026.md`

Usuário pediu: dentro de um mesmo projeto, cadastrar "contextos" nomeados (ex.: "obras de 34,5kV"),
cada um com seu próprio conjunto de ajustes ligados/desligados/sobrescritos, selecionáveis por um
seletor na linha de totais (postes/cabos/equipamentos).

Desenho (decidido com o usuário após 3 perguntas de estratégia): **terceira camada de overlay**
sobre o que já existia — `padrão (DEFAULT) → overlay do projeto → overlay do contexto`. A peça-chave
é que `services/ajustes_camadas.py::efetivo(padrao, overlay)` não sabe que só existe uma camada —
ela só recebe uma lista e um overlay, então pode ser **chamada em cadeia** sem alterar uma linha
dela: `efetivo(efetivo(padrao, overlay_projeto)[0], overlay_contexto)`. Zero mudança no motor de
camadas existente.

Implementado:
- **Tabelas novas**: `contextos` (metadados de quais contextos existem por projeto — pensada para
  um dia ser compartilhada por Regras de Domínio também) e `ajustes_contextos`/
  `ajustes_contextos_historico` (overlay de Ajustes por contexto, mesmo shape do overlay de
  projeto). Nenhuma tabela existente foi alterada.
- **`services/repo_json.py`**: nova classe `RepoJsonContexto`, chaveada por
  `(projeto_codigo, contexto)` — irmã de `RepoJson` (que continua intacta, chaveada só por
  `projeto_codigo`).
- **`routers/validacao_ajustes.py`**: `resolver(projeto_codigo, contexto=None)` resolve a cadeia;
  toda rota (`GET/POST /`, `/preview`, `/preview-lote`) ganhou o campo opcional `contexto` —
  **sem ele, comportamento idêntico a antes** (garantido por teste de regressão). Novas rotas
  `GET/POST /contextos` e `POST /contextos/excluir`.
- **Frontend**: novo seletor "Contexto" na linha de totais do Resumo (`#resumo-rede-painel`),
  mesmo padrão visual do seletor de projeto, com "+ Novo"/"✕ Excluir" admin-only, persistido por
  projeto em `localStorage`. O botão "Ajustar" e o editor de Ajustes da gaveta (`painel_ajustes.js`)
  passam a considerar o contexto selecionado — a gaveta mostra "Editando o contexto..." e desabilita
  Histórico/Restaurar nesse modo (decisão de escopo: histórico por contexto fica para outra tarefa).
- Trocar de contexto **não recalcula nada na tela na hora** — só define o que roda na próxima vez
  que o usuário clicar em "Ajustar" (decisão do usuário).

Escopo desta rodada: só **Ajustes**. Modo autônomo (pasta monitorada) e Regras de Domínio não usam
contexto ainda — ficam para tarefas futuras, reaproveitando o mesmo desenho de dados.

11 testes novos (`tests/test_ajustes_contextos.py`): CRUD, resolução em cadeia, independência entre
contextos, exclusão isolada, regressão byte a byte sem `contexto`, `preview-lote` respeitando o
contexto. Suíte completa (732 testes) sem regressão. Smoke test real-server + Playwright confirma o
fluxo completo pela UI: criar contexto, editar dentro dele na gaveta, salvar, e confirmar que o
projeto (sem contexto) permanece intacto.

---

## 2026-10-06 — TASK-043: suporte a .dwg via conversão online (CloudConvert)

**Tipo:** feature · `services/cloudconvert_service.py`, `routers/upload.py`, `routers/health.py`,
`services/autonomo/pipeline.py`, `static/index.html`, `static/script.js`, `requirements.txt` ·
`.ai/tasks/TASK-043-06-10-2026.md`

Usuário perguntou se o programa conseguia ler `.dwg`. Resposta: não diretamente — é formato binário
proprietário da Autodesk, e `ezdxf` (usado em `services/dxf_service.py`) só lê `.dxf`. Usuário pediu
para o próprio programa converter internamente usando uma ferramenta online. Duas decisões
confirmadas antes de implementar: (1) serviço CloudConvert (API v2, chave simples, free tier) em
vez de Autodesk Platform Services (OAuth2 mais complexo); (2) credencial no mesmo padrão já usado
para a chave do Gemini (variável de ambiente como padrão + o usuário pode digitar a própria, salva
no servidor só depois de confirmada por um uso com sucesso).

Implementado:
- **`services/cloudconvert_service.py`** (novo): `converter_dwg_para_dxf()` — cria job no
  CloudConvert (import/upload → convert → export/url), envia o arquivo, espera (polling), baixa o
  `.dxf` resultante. `ErroConversaoDwg` com mensagem clara em qualquer falha.
- **`routers/upload.py`**: resolve a credencial (header `X-CloudConvert-Key` > `configuracoes` >
  `CLOUDCONVERT_API_KEY` do ambiente) e converte `.dwg` → `.dxf` antes de extrair, em `POST /upload`
  e `GET /extract-local`. `.dxf`/`.pdf` continuam sem tocar no CloudConvert.
- **`services/autonomo/pipeline.py`**: `.dwg` aceito na pasta monitorada (`EXTENSOES`), convertido
  antes de ler (só credencial salva/padrão — sem header possível em segundo plano).
- **Frontend** (`static/index.html`/`script.js`): `.dwg` aceito no seletor de arquivo da tela
  Leitor; novo botão "Chave DWG" + modal para configurar a própria chave do CloudConvert
  (localStorage, enviada como header).
- **`routers/health.py`**: novo campo `cloudconvert_key_source`, mesmo padrão de `ai_key_source`.

Pela primeira vez no projeto, uma funcionalidade de **leitura de arquivo** passa a depender de
internet (só para `.dwg`) — aceito como exceção pontual, mesmo padrão já aplicado à IA e à
sincronização com Supabase.

Testado com `httpx` sempre mockado (nenhuma chamada de rede real em teste): 9 testes novos de
`cloudconvert_service`, 10 de dispatch em `upload.py`, 3 do pipeline autônomo. `tests/conftest.py`
passou a zerar `CLOUDCONVERT_API_KEY` do ambiente também. Suíte completa (721 testes) sem
regressão. Smoke test real-server + Playwright confirma a UI nova sem erros de JS.

---

## 2026-10-06 — TASK-042: editor de Regras de Vinculação e Regras de Conectores dentro do modal

**Tipo:** feature (UI) · `static/resumo.html`, `static/resumo.js` · `.ai/tasks/TASK-042-06-10-2026.md`

Usuário pediu para tornar as `regras_vinculacao` (quantidade de cabos e compatibilidade por tipo
de estrutura) editáveis numa UI dentro do modal de vinculação, e para poder acessar o ativo
composto (cabo+estrutura vinculados, ex. `CAA2_U4`) a partir de uma regra com formato amigável,
também dentro do modal.

Investigação confirmou dois mecanismos já existentes, mas sem UI própria no lugar certo:

- `regras_vinculacao` já tinha endpoint (`GET/POST /api/regras-vinculacao`), mas nenhuma tela
  editava — só o algoritmo de vínculo automático lia.
- O ativo composto só existe em memória durante o cálculo da Totalizadora
  (`itens_vinculo_para_totalizadora`) e já podia ser convertido em algo real (ex. um conector) via
  **Regra de Conversão com `origem: VINCULO`** (TASK-011/032/034) — só que misturado com as regras
  de CABOS/OUTROS na view Totalizadora, exigindo digitar `ativo_de` cru (`"CAA2_U4"`).

**Decisão: não estender o motor de Ajustes** (`services/ajustes_planilhas.py`) para ler ativos
compostos — ele roda antes do vínculo/Totalizadora existir, e replicar a lógica de vínculo ali
duplicaria o que a Regra de Conversão com `origem: VINCULO` já resolve. Em vez disso, duas seções
accordion novas foram adicionadas dentro de `#modal-vinculacao`:

1. **Regras de Vinculação**: tabela editável (tipo/qtd/compatibilidade/descrição) sobre o endpoint
   já existente `/api/regras-vinculacao`.
2. **Regras de Conectores**: formulário amigável (tipo de cabo + tipo de estrutura + ativo
   resultante + quantidade) sobre a MESMA tabela de Regras de Conversão da Totalizadora
   (`tableStates.regras.data`, `/api/regras/conversao`), filtrada para `origem === 'VINCULO'` e
   traduzindo `ativo_de ↔ "<CABO>_<ESTRUTURA>"` automaticamente. Editar aqui ou na Totalizadora é a
   mesma regra.

Nenhum endpoint ou tabela nova. Verificado via smoke test real-server + Playwright (DB temporário,
nunca o `banco_resumo.db`): adicionar/editar/salvar nas duas seções persiste corretamente
(confirmado por `GET` de volta), e uma linha inválida é rejeitada pelo backend com erro visível.

---

## 2026-10-05 — TASK-039/040/041: disponibilidade de dados e reconexão com o Supabase

**Tipo:** correção de robustez + feature · `services/supabase_client.py`,
`services/connectivity_monitor.py`, `app.py`, `routers/obras.py`, `recs.py`, `projetos.py`,
`admin.py`, `static/resumo.html`, `static/resumo.js`, `tests/test_supabase_client.py` ·
`.ai/tasks/TASK-039-05-10-2026.md`, `TASK-040-05-10-2026.md`, `TASK-041-05-10-2026.md`

Usuário relatou "a conexão com o Supabase está se perdendo após um tempo" e pediu garantia de que
os dados fiquem sempre disponíveis, mantendo a última versão carregada em caso de queda. Investigação
confirmou a causa raiz:

- `services/supabase_client.py` criava o cliente Supabase como singleton **uma única vez por
  processo** — existia `reset_supabase_client()`, mas nada no projeto a chamava.
- Toda falha de chamada à nuvem era engolida (`except Exception: pass`, sem log nenhum) em
  `routers/obras.py`, `recs.py`, `projetos.py`, `admin.py` (padrão já catalogado em `.ai/STATE.md`
  item 21).
- Resultado: quando o Supabase fecha uma conexão ociosa (comportamento normal do
  PostgREST/pgbouncer), a primeira falha era silenciosa e **nunca mais havia tentativa de
  reconexão** até o processo reiniciar.
- `services/connectivity_monitor.py` já existia pronto (ping periódico + evento WebSocket), mas
  nunca era chamado, e era restrito a `APP_MODE != "server"`.
- A gravação local (SQLite) **já era garantida** independente da nuvem (`save_obra` sempre grava
  local) — o problema real na escrita era a falta de aviso, não perda de dado.

Implementadas as três tarefas registradas:

- **TASK-039** (reconexão reativa): `get_supabase()` ganhou um TTL proativo de 180s (recria o
  cliente sozinho mesmo sem falha observada) combinado com reset reativo via nova função
  `registrar_falha(contexto, exc)`, que reseta o singleton só quando a exceção é classificada como
  erro de rede (`httpx.TransportError`/`httpx.TimeoutException`/`ConnectionError`/`OSError`/
  `TimeoutError`), nunca em erro de autenticação/dados.
- **TASK-040** (monitor de conectividade): `services/connectivity_monitor.py` corrigido (loop
  `asyncio` capturado uma vez no startup, broadcast via `run_coroutine_threadsafe` em vez de
  criar/fechar loops a cada ciclo) e ligado em `app.py` no startup, **sem** a restrição de
  `APP_MODE` — roda em desktop e servidor. A transição offline→online agora força
  `reset_supabase_client()`. Novo indicador visual no frontend (`static/resumo.html`/`resumo.js`,
  `#indicador-conectividade`) escuta o evento WebSocket `{"type": "connectivity", ...}` e exibe um
  selo "Sem conexão com a nuvem — trabalhando localmente" quando offline.
- **TASK-041** (visibilidade de falha): os 16 pontos de `except Exception: pass` silenciosos em
  `obras.py`/`recs.py`/`projetos.py`/`admin.py` agora chamam `registrar_falha(contexto, e)` —
  sempre loga em `warning`, nunca mais 100% silencioso. Decisão tomada: manter o comportamento
  atual de escrita (sempre grava local, sempre responde sucesso), só com log — mudança de maior
  impacto (erro visível ao usuário) descartada por não ter sido pedida explicitamente e por exigir
  tratamento em todas as telas de salvamento. Confirmado por leitura de código que nenhuma leitura
  troca dado local bom por resultado de nuvem vazio por falha disfarçada (o `except` sempre
  intercepta antes do `if res.data is not None`).

Novo `tests/test_supabase_client.py` (14 testes) cobre TTL, reset reativo por tipo de exceção e log
sempre emitido. Suíte completa (699 testes) sem regressões.

---

## 2026-10-05 — TASK-038: `adicionar_ativo` aceita quantidade negativa (`*`, linha viva)

**Tipo:** melhoria (backend) · `.ai/tasks/TASK-038-05-10-2026.md`

Pedido do usuário: um ajuste que, ao achar TR110 (trafo mono) na linha, adiciona um conjunto de
ativos — quatro positivos e um negativo (`*1-PR`). A ação `adicionar_ativo` (TASK-022) exigia
`qtd > 0` e, mesmo que aceitasse, formatava o token com `-` literal, que o tokenizador de Outros
não lê como sinal (só `*` é lido como negativo, por `services/orcamento_calc.py:tokenizar_outros`).

- `services/ajustes_planilhas.py`: `validar_acoes` passa a aceitar `qtd` negativo (só `0` é
  rejeitado); `_num` agora reconhece um token existente com `*` e devolve negativo; nova
  `_fmt_qtd()` (inverso de `_num`) gera o prefixo `*` corretamente; usada nas 3 montagens de token
  de `_t_adicionar_ativo` (novo, somar, substituir) e na frase legível do ajuste.
- `data/manual_regras_e_ajustes.md`: documentado o novo comportamento com exemplo.
- Confirmado isolado: `_num`/`_juntar`/`_RE_NUM`/`_fmt_qtd` são privados deste módulo (não
  importados de `regras_dominio.py`), sem efeito em validação de domínio, Totalizadora ou modo
  autônomo.

Testado: `pytest tests/test_ajustes_planilhas.py` (76 testes, 5 novos) + `pytest tests/` completo
(684 testes) + verificação manual do exemplo completo do usuário (5 ações, mesma condição `{"tem":
"TR110"}`) produzindo exatamente `1-TR110 1-PR15 2-P50 1-PR127 *1-PR 3-EST35`.

---

## 2026-10-03 — TASK-037: coordenada do Leitor chega ao Resumo (fecha o 4º bug do vínculo)

**Tipo:** correção de bug · `.ai/tasks/TASK-037-03-10-2026.md`

Último elo que faltava para o vínculo automático por coordenada (TASK-032) funcionar de ponta a
ponta na **tela manual**. Os três bugs da TASK-036 (dígito do tipo, ativo composto, descarte de
`_x`/`_y` no DXF) já tinham corrigido o lado do dado; faltava a ponte entre as duas telas: o botão
"Processar" da tela Leitor (`static/script.js`) montava o objeto exportado para
`localStorage['processar_dados']` só com `{pagina, texto, cor, layer, entidade, operacao, ativo}` —
sem `_x`/`_y`, mesmo quando `extractedDataCache[idx]` já os tinha.

- `static/script.js`: inclui `_x`/`_y` condicionalmente no objeto exportado, mesmo padrão já usado
  em `services/autonomo/montagem.py:_item_exportado`.

**Testado end-to-end com um DXF real, sem nenhum atalho sintético** (primeira vez que isso acontece
nesta funcionalidade): servidor real + Playwright — upload de um DXF com 1 cabo e 2 estruturas →
extração real → "Processar" → aba do Resumo já recebe `_x`/`_y` → vínculo automático cadastrado via
API → `tentarVincularAutomaticamente()` vincula corretamente, respeitando a ordem de proximidade.
`pytest tests/` completo continua em 679/679 (mudança restrita a `static/script.js`, sem teste
Python tocando esse arquivo).

Com isso, a cadeia completa TASK-032/033/034/036/037 está funcional nos dois modos (manual e
autônomo). Única pendência remanescente conhecida: nenhuma, para esta funcionalidade.

---

## 2026-10-03 — TASK-036: vínculo cabo↔estrutura no modo autônomo (+ 3 bugs achados e corrigidos)

**Tipo:** nova funcionalidade (backend) + correção de bugs pré-existentes · `.ai/tasks/TASK-036-03-10-2026.md`

Implementa o que a TASK-036 tinha registrado: o pipeline do modo autônomo (TASK-031) agora gera
vínculo cabo↔estrutura/poste por coordenada, roda a validação de vínculo (TASK-033) e aplica a
origem `VINCULO` das Regras de Conversão (TASK-034) ao montar o orçamento — as três coisas que,
até aqui, só existiam dentro de `static/resumo.js`.

- `services/autonomo/montagem.py`: `_x`/`_y` passam a chegar até as linhas de Cabos/Outros; nova
  `prefixo_cabo()`.
- Novo módulo `services/autonomo/vinculacao.py`: porte fiel de `tentarVincularAutomaticamente`,
  `avaliarVinculacaoLocal` e do bloco de ativo composto de `syncTotalizadora`.
- `services/autonomo/pipeline.py`: nova etapa `vincular`; `Contexto.regras_vinculacao`.
- `services/autonomo/totalizadora.py`: `montar_totalizadora()` ganhou `itens_vinculo`; composto
  sem regra correspondente nunca é mantido (decisão já tomada na TASK-034, replicada aqui).
- **Decisão do usuário:** auto-link sem combinação válida nunca bloqueia nem vira pendência de
  confirmação — só um achado de aviso no relatório.

**Três bugs pré-existentes achados durante a implementação (confirmados com o leitor real antes de
corrigir, não por suposição) — corrigidos por decisão explícita do usuário:**

1. `_tipoEstruturaVinculo` (TASK-032/033) pegava o dígito da **quantidade**, não do tipo da
   estrutura — o ativo real de uma linha Outros é sempre `"<qtd>-<código>"` (ex.: `"1-U4"`, nunca
   `"U4"` sozinho; confirmado rodando o motor de regras real). Isso quebrava silenciosamente todo
   o casamento de regra por tipo desde que a TASK-032/033 foi mesclada. Corrigido nos dois lados
   (`static/resumo.js` e `services/autonomo/vinculacao.py`).
2. O mesmo problema na geração do ativo composto em `syncTotalizadora` (`CAA2_1-U4` em vez de
   `CAA2_U4`). Mesma causa raiz, mesma correção.
3. `services/dxf_service.py:extract_dxf_content()` calculava `_x`/`_y` só para ordenar a lista de
   saída e depois os **descartava** antes de devolver — nenhum arquivo DXF jamais teve coordenada
   disponível para o vínculo automático (só PDF, desde a captura adicionada no início desta
   sessão). Corrigido: os campos agora ficam no retorno.

**Pendência registrada, não corrigida aqui** (`.ai/tasks/TASK-037-03-10-2026.md`): `static/script.js`
(botão "Processar" da tela Leitor) ainda não repassa `_x`/`_y` para `localStorage['processar_dados']`
— o vínculo automático por coordenada continua não funcional na **tela manual** até isso ser
corrigido (o modo autônomo não depende de `script.js`, por isso já funciona com as correções acima).

Testado: `pytest tests/` completo (679 testes, todos passando — 23 novos em
`test_autonomo_vinculacao.py`, +2 em `test_autonomo_montagem.py`, +2 de pipeline completo em
`test_autonomo_pipeline.py`) + verificação manual em servidor real/Playwright na tela `/resumo`
confirmando o auto-link escolhendo a estrutura mais próxima com os bugs corrigidos.

---

## 2026-10-03 — TASK-035: filtro por projeto no histórico do modo autônomo (+ registro da TASK-036)

**Tipo:** melhoria (backend + frontend) · `.ai/tasks/TASK-035-03-10-2026.md`

Pedido do usuário depois de discutirmos uma estratégia para rodar vários projetos pelo modo
autônomo (TASK-031) de uma vez e só depois ir conferindo. A arquitetura já suporta múltiplos
projetos numa única instância (`entrada/<código-do-projeto>/`), mas o histórico de execuções
(`GET /api/autonomo/execucoes`) só filtrava por status — inviável de triar visualmente num lote
grande com vários projetos misturados.

- `services/autonomo/execucoes.py:listar()` ganhou o parâmetro `projeto` (filtro `AND` com
  `status`); `routers/autonomo.py` repassa.
- `static/autonomo.html`/`autonomo.js`: novo `<select id="filtro-projeto">`, populado a partir de
  `GET /api/projetos` (que devolve `{"projetos": [...]}` — achado durante o teste: a primeira
  versão do JS lia como lista direta).
- Testado: `pytest` (33 testes, 2 novos) + servidor real/Playwright confirmando o seletor populado
  e os dois filtros combinando sem erro de console.

**Também registrada (sem implementar ainda):** `.ai/tasks/TASK-036-03-10-2026.md` — o modo
autônomo foi construído (TASK-031, 2026-10-02) **antes** do vínculo cabo↔estrutura/poste
(TASK-032/033/034, 2026-10-03) e não tem nenhuma referência a ele (confirmado por busca no
código): não gera vínculo automático por coordenada, não roda a validação de vínculo, e não
aplica a origem `VINCULO` das Regras de Conversão ao montar o orçamento. Um arquivo processado
pelo autônomo num projeto que depende de vínculo sai sem essa parte do orçamento. Fica registrada
para decisão/implementação futura.

---

## 2026-10-03 — TASK-032/033/034: vínculo cabo↔estrutura, validação e ativo composto

**Tipo:** nova funcionalidade (backend + frontend) · `.ai/tasks/TASK-032-02-10-2026.md`,
`TASK-033-02-10-2026.md`, `TASK-034-02-10-2026.md`

**TASK-032** — modal "VINCULAR CABOS" na aba Resumo: vincula cada cabo a 1 estrutura/poste
(operação M/\*M) ou 2 (operação I), manualmente ou por um algoritmo opcional de vínculo
automático por coordenada (`_x`/`_y` — já existia no DXF; `services/pdf_service.py` passou a
capturar o `bbox` do PyMuPDF para o PDF também). Nova tabela por projeto `regras_vinculacao`
(quantidade de cabos e compatibilidade — mesmo/qualquer tipo-fase-operação — exigida por tipo de
estrutura, identificado pelo primeiro dígito do ativo). O vínculo vive como um campo
(`vinculoEstruturas`) direto na linha do cabo, o que faz persistência (undo/redo, salvar/carregar
obra) funcionar sem nenhum caminho novo de serialização — só foi preciso estender a `deepClone()`
compartilhada de `static/resumo.js`, que até então descartava qualquer campo fora de
`entidade`/`operacao`/`ativo`/`qtdAtivos` (incluindo a própria coordenada).

**TASK-033** (depende da TASK-032) — ao validar (botão já existente `#btn-validar`), cada
estrutura/poste tem seus vínculos conferidos contra a regra do seu tipo; qualquer violação
(zero vínculos, quantidade errada, incompatibilidade) gera um achado de aviso. Zero vínculos é
perdoado se existir, em qualquer lugar do projeto, um cabo M/\*M ou um ativo `RETCA`.
**Decisão de integração**: em vez de um painel visual novo (como o desenho original previa, nos
moldes do "Resumo da rede" da TASK-008), os achados foram integrados ao sistema de Validação já
existente (`routers/validacao.py`/`services/regras_dominio.py`, construído em sessões anteriores)
— como o vínculo só existe no frontend, a verificação roda localmente
(`avaliarVinculacaoLocal()`) e seus achados são concatenados aos de `/api/validacao/planilhas`
antes de exibir, sem nenhuma mudança no motor/backend de domínio.

**TASK-034** (depende da TASK-032) — cada vínculo gera um ativo composto
`<prefixo_do_cabo>_<ativo_da_estrutura>` (ex.: `CAA2_N4`), que existe só dentro dos dados do
vínculo mas fica disponível como entrada para as Regras de Conversão (TASK-003) via uma nova
origem `VINCULO` — ex.: `origem=VINCULO, ativo_de=CAA2_N4, ação=ADIÇÃO, ativo_para=ALCA2,
fator=3` gera `3-ALCA2` na Totalizadora ao montar o orçamento, pelo mesmo motor ADIÇÃO/SUBST que
já existe. Diferente de `CABOS`/`OUTROS`, um composto sem regra correspondente não fica na
Totalizadora (não é um código de orçamento real).

**Renumeração**: as três tarefas foram inicialmente registradas como TASK-009/010/011, mas esses
números já tinham sido usados por outra sessão (fundidos em `main` enquanto esta sessão
trabalhava) — renumeradas para TASK-032/033/034 antes de implementar, mesma situação já ocorrida
no início desta sessão com TASK-001-004.

---

## 2026-10-08 — TASK-056: exportar e importar obras (`.obra.json`)

**Tipo:** Backend + frontend · `.ai/tasks/TASK-056-08-10-2026.md`

Cada obra do modal Carregar Obra ganhou **Exportar** (baixa `<nome>.obra.json` com Cabos, Outros e Totalizadora; sem regras de conversão, usuário, e-mail ou marcas do autônomo) e o modal ganhou
**Importar obra**: qualquer usuário importa e recebe uma **cópia particular** (id novo, nome com " (importada)" se repetir), **só de obra do projeto selecionado na tela**. O servidor valida formato, versão,
tamanho (5 MB), linhas (5.000) e campos; o conteúdo entra na tela só como texto. Sem mudança de schema. As TASK-057 (obras públicas/particulares) e TASK-058 (modelos com `V`) estão planejadas com as decisões confirmadas.

---

## 2026-10-08 — TASK-057: obras públicas e particulares

**Tipo:** Backend + frontend + schema · `.ai/tasks/TASK-057-08-10-2026.md`

Cada obra tem visibilidade: **particular** (padrão) ou **pública** (visível, para ver/copiar/carregar/adicionar/subtrair/exportar, só a usuários do **mesmo projeto**). Particulares de outros nunca aparecem (lista, índice, por id, exportar, IA).
A lista ganhou filtros (Visibilidade, Origem), selos e o dono (só a parte antes do `@`); só o dono altera a visibilidade, sobrescreve ou apaga (salvar sobre id alheio dá 403). Obras sem dono: só administradores, com "Assumir" (viram particulares do admin).
A IA vê as próprias + públicas do projeto; o modo autônomo grava sempre particular. **Schema Supabase:** `obras.publica` e `obras.dono_nome` (SQL em `scripts/schema_supabase.sql`, a rodar pelo usuário; sem ele tudo segue como antes).

---

## 2026-10-08 — TASK-058: modelos de obra com variável V

**Tipo:** Backend + frontend + schema · `.ai/tasks/TASK-058-08-10-2026.md`

Novo **Salvar como modelo**: o programa acha os `V` / `V(nome)` nas tabelas (Cabos: comprimento, ex. `CAA 2 ABC V m`; Outros: quantidade, ex. `V-U4`, `*V-CFU`; só o token exato — `CAV`, `V1`, `1-VA` não contam), pede rótulo e valor padrão e grava um modelo
**público no projeto** (só o criador edita/exclui; sem Totalizadora). Carregar/Adicionar/Subtrair um modelo — e o comando da IA — abre **"Parâmetros do modelo"**; o servidor (`services/modelos_obra.py`, função pura) gera a obra padrão, validada pela camada 1, **sem gravar**; Ctrl+Z desfaz.
Lista com selo 📐 Modelo, filtro Tipo e "Editar modelo" (só do criador). `V` sobrando bloqueia **Montar Orçamento** e a validação ganhou `C1-VAR`. Exportar/importar aceita modelos (o importado vira modelo de quem importa).
**Schema Supabase:** `obras.tipo` (SQL em `scripts/schema_supabase.sql`, a rodar pelo usuário; sem ele tudo funciona, reconhecendo o modelo pela marca no `dados_json`).

---

## 2026-10-08 — TASK-059 (etapa 1): aprendizado supervisionado do modo autônomo

**Tipo:** Backend + frontend · `.ai/tasks/TASK-059-08-10-2026.md`

Ao salvar uma obra que veio do autônomo (a tela envia `origem_execucao`), o servidor guarda **só as diferenças** para o que o autônomo entregou (troca/remoção/adição de ativo, texto trocado em cabo, linha excluída/adicionada, operação/quantidade alterada). Padrões repetidos (≥ 3 vezes, ≥ 2 obras, ≥ 80% dos casos) viram **propostas de ajuste**,
validadas pelo motor; **Aprovar** cria o ajuste no projeto (com histórico), **Recusar** não volta. Depois de 10 obras corrigidas, "Sugestões da IA" propõe mais (só códigos de ativo e contagens; sempre pendente). Seção "Aprendizado" em `/autonomo`; só administrador, só o próprio usuário, por projeto; armazenamento local (SQLite), sem mudança de schema do Supabase.

---

## 2026-10-09 — TASK-060: modelos com ativo variável `X(opções)`

**Tipo:** Backend + frontend · `.ai/tasks/TASK-060-09-10-2026.md`

Nos modelos (Outros) o **ativo** também pode ser variável: `V(quantidade)-X(U3,U4,U5)` — quem usa o modelo **escolhe num dropdown** entre as opções (2 a 30 escritas). `X(poste:A,B)` + `X(poste:)` repete a mesma escolha; `X(@GRUPO)` usa os ativos do grupo do projeto (expandido na hora do uso). O servidor só aceita uma das opções; opção que não existe na base técnica só **avisa** ao salvar; padrão opcional (sem padrão é preciso escolher).
`X(...)` sobrando bloqueia Montar Orçamento e a validação (`C1-VAR`). Modelos antigos não mudam; sem mudança de schema.

---

## 2026-10-08 — TASK-059 (etapa 2): regras do leitor aprendidas + aprendizado no trabalho manual

**Tipo:** Backend + frontend · `.ai/tasks/TASK-059-08-10-2026.md`

O aprendizado passa a valer também no **trabalho manual**: o botão **Processar** do Leitor envia o que o motor preencheu × o que você deixou nas colunas, e o Resumo guarda a foto das tabelas vindas do Leitor para comparar com a obra salva (com os ajustes ativos aplicados à foto, como o autônomo faria).
No autônomo cada linha agora carrega a origem (texto/cor/camada/coordenada). Com esses itens o programa **propõe regras do leitor** (classificação e processamento): cada regra é **simulada com o motor real** e só aparece se corrigir ≥ 80% do grupo sem estragar nenhum item que já estava certo. Aprovar grava a regra pelo cadastro de regras do leitor (com histórico para reverter).
Mudança de privacidade: texto, cor e camada dos itens observados ficam guardados no SQLite **local** do usuário (apagáveis em "Apagar histórico").

---

## 2026-10-13 — TASK-062: botão "Ler projeto" no Resumo

**Tipo:** Frontend · `static/ler_projeto.js` (novo), `static/resumo.js`, `static/resumo.html`, `app.py` · `.ai/tasks/TASK-062-13-10-2026.md`

O Resumo ganhou **Ler projeto**: escolhe (ou arrasta sobre o botão) um PDF/DXF/DWG; o arquivo vai ao `/upload`, os textos passam pelo mesmo motor do Leitor + a lógica do Resumo, no navegador, e as tabelas são preenchidas **na mesma aba**. Com conteúdo nas tabelas pergunta **Substituir** (padrão) / **Adicionar** / **Cancelar**. Ctrl+Z desfaz; grava `processar_dados` (F5 mantém). Aprendizado: Substituir reinicia o `baselineManual`; Adicionar o zera. Sem mudança de schema nem de API. Paridade com o fluxo antigo provada por `tests/test_ler_projeto.py`.

---

## 2026-10-27 — TASK-084: modal de vinculação em abas e parâmetro `cabo` pelo vínculo

**Tipo:** Backend + Frontend · `services/vinculacao_regras.py`, `static/vinculacao_regras.js`, `static/vinculacao_ui.js` (novos), `routers/vinculacao_config.py` (novo), `data/vinculacao_pernambuco_seed.json`, `static/resumo.js`, `static/resumo.html`, `static/anotacao_outros.js`, `services/anotacao_outros.py`, `services/autonomo/{pipeline,totalizadora,vinculacao,leitor_js}.py`, `database.py`, `config.py`, `app.py`, `scripts/schema_supabase.sql`

O botão **VINCULAR CABOS** abre um modal com cinco abas: **Vínculo** (cabos × estruturas com linhas; preto = sem vínculo, roxo = parcial, verde = completo, vermelho = inválido; o quadrado fica só na estrutura e o poste é só contexto; clicar cabo → estrutura ou o contrário; botão direito na estrutura = **cabo fictício**), **Tabela de Cabos** e **Tabela de Estruturas** (o que cada lado aceita, quantas pontas, lados iguais, amarração, valor do parâmetro `cabo`), **Script** (JSON de vínculo automático por proximidade/sequência, com pré-visualização) e **Mapa**. Estruturas com o parâmetro `cabo` sem valor digitado recebem o do cabo vinculado na Totalizadora (valor digitado vence; N2 usa o maior C > P > A; sem vínculo = padrão). As regras antigas (tipo/qtd/compatibilidade) são migradas para a Tabela de Estruturas; Pernambuco nasce com um ponto de partida. O modo autônomo usa as mesmas tabelas e o 1º script do projeto. Também: a linha compacta digitada (`9200D, N1`) passa a ser classificada pelo texto como escrito (POSTE em Pernambuco). **SQL novo:** tabela `vinculacao_config` em `scripts/schema_supabase.sql`.

---

## 2026-10-27 — TASK-082: poste com itens e linhas do mesmo MTEXT

**Tipo:** Frontend (motor do Leitor) + seed · `static/regras_leitor_engine.js`, `static/script.js`, `routers/regras_leitor.py`, `data/regras_leitor_*_pernambuco_seed.json`

Pernambuco: `9200D` e `9200D, N1` são POSTE; no DXF, até 2 linhas de baixo do mesmo MTEXT são juntadas ao poste (`9200D, N1, N2`) e ignoradas. O motor ganhou opções opcionais de vizinhança (`mesmo_mtext`, `juntar`, `obrigatoria`, `direcao: "ANTES"`, placeholder `{vizinhancas}`) — regras antigas não mudam. Quem já tem as regras do 081 salvas precisa atualizá-las no editor.

---

## 2026-10-27 — TASK-082/083: regras do Leitor de Pernambuco e notação do LABEL na tabela Cabos

**Tipo:** Backend + Frontend · `services/anotacao_cabos.py` e `static/anotacao_cabos.js` (novos), `static/regras_leitor_engine.js`, `static/script.js`, `routers/regras_leitor.py`, `database.py`, `config.py`, `data/regras_leitor_*_pernambuco_seed.json` (novos), `static/resumo.js`, `static/resumo.html`, `services/autonomo/montagem.py`, `services/autonomo/totalizadora.py`, `services/validacao_planilhas.py`, `services/regras_dominio.py`, `services/anotacao_outros.py`, `static/anotacao_outros.js`, `app.py`

Projeto **081 (Pernambuco)** nasce com regras do Leitor próprias: operação pela layer (`01_INSTALANDO`/`01_REMOVER`/`01_REAPLICANDO`) e, sem layer, pela cor exata (#000000 = instalar, #000001 = remover); postes compactos (`11300D`, D/R/F/FV) e cabos pelo LABEL com o comprimento na mesma linha ou na de baixo (a tabela mantém o texto como escrito, inclusive a tensão, que só é ignorada na Totalizadora); ruído ignorado. O motor ganhou **duas condições opcionais** (regras antigas não mudam): `cor_hex_em`/`cor_hex_nao_em` e `vizinhanca.direcao: "DEPOIS"`; o editor de regras mostra os campos. A tabela **Cabos** aceita `3xA35P(7,94)-13,8kV 45 m` (fases = `Nx`, tensão ignorada, neutro à parte: `(7,94)` → `CAZ7,94`, `(16)` em `C16` → `C16`, `S50P(35)` → `S35P`; `(S04)`/`[C16]` soltos, `RAMA25`/`PB35`, `16` = `C16`, `C 3x1x95(95)` = código único); o neutro entra na soma de rede; a conversão é implícita (a célula não muda). Em Outros, poste sozinho (`11300D`) vale 1 poste. Sem SQL novo.

---

## 2026-10-27 — TASK-081: notação compacta na tabela Outros e apelidos de ativo

**Tipo:** Backend + Frontend · `services/anotacao_outros.py` e `static/anotacao_outros.js` (novos), `services/ativo_parametros.py`, `static/ativo_parametros.js`, `services/orcamento_calc.py`, `services/validacao_planilhas.py`, `services/ajustes_planilhas.py`, `services/autonomo/totalizadora.py`, `services/autonomo/leitor_js.py`, `routers/validacao.py`, `static/resumo.js`, `static/resumo.html`, `app.py`

Outros aceita linhas como `11600D, 2 N3, T3` (itens separados por vírgula + espaço; `quantidade ativo`; vírgula sem espaço é decimal). Poste compacto: `D`=DT, `R`=RC, `F`=CV, `FV` (`11300D` = `DT11/300`). Apelidos entram na célula ATIVO da base (`DT11/300, 11300D`) e valem em qualquer item, nos dois formatos. A tradução para o formato antigo é implícita (a célula mantém o que foi digitado); uma linha não mistura formatos (erro `C1-OUT-MISTURA`; vírgula sem espaço fora de decimal = `C1-OUT-VIRGULA`). Estruturas com o parâmetro `poste` recebem o poste mais próximo à esquerda (`N3(11600D)`). **Mudança de comportamento:** na operação `M`/`*M` o poste é ignorado e os demais itens viram `I`/`*I` em TODAS as linhas de Outros (antes entravam na soma de R). Contadores do resumo passam a contar RC e FV como postes. Linhas compactas não são editadas pelos ajustes. Sem SQL novo.

---

## 2026-10-26 — TASK-080: operadores nos parâmetros dos ativos e COMPONENTE automático

**Tipo:** Backend + Frontend · `services/ativo_parametros.py`, `static/ativo_parametros.js`, `services/orcamento_calc.py`, `routers/admin.py`, `routers/orcamento.py`, `static/orcamento.html`

Os grupos de ativos com parâmetros aceitam curinga `%` (`N1(poste=D%)`), lista `|` (`N1(cabo=ALUM|COBRE)`) e exceto `!=` (`N1(cabo!=XLPE, poste=D%)`); o pedido segue com valores concretos (`N1(ALUM,D12)`). O `COMPONENTE` em branco dos grupos é preenchido automaticamente na importação/salvamento (`N1`, `N1.CABO=XLPE`, `N1.POSTE=D%`…), com herança para as linhas seguintes e sem alterar o que já foi digitado. Aviso de sobreposição de códigos mais preciso. Sem mudança no Supabase.

---

## 2026-10-26 — TASK-079: ativos com parâmetros por categoria (ex.: N1(XLPE,DT))

**Tipo:** Backend + Frontend · `services/ativo_parametros.py` e `static/ativo_parametros.js` (novos), `services/orcamento_calc.py`, `services/regras_dominio.py`, `services/validacao_planilhas.py`, `routers/admin.py`, `routers/orcamento.py`, `routers/validacao_regras.py`, `static/resumo.js`, `static/resumo.html`, `static/orcamento.html`, `static/resultado_orcamento.html`, `app.py`

A base de orçamento aceita variantes por categoria no nome do ativo: declaração `N1=N1(cabo=XLPE, poste=DT)` (nomes, ordem e padrões; `N1=N1(cabo, poste)` = sem padrão), grupos `N1(cabo=XLPE)`, `N1(poste=DT)`, `N1(cabo=COBRE, poste=CC)` e o grupo comum `N1`. Em Outros/Totalizadora o pedido usa só os valores (`3-N1(XLPE,DT)`; `3-N1` completa pelo padrão). O cálculo junta o grupo comum e os grupos que valem; erro de variante vai para "não encontrados" com a lista e pedido incompleto gera **avisos**. A importação recusa cadastro inconsistente e avisa sobreposição de códigos. Base sem parâmetros: resultado idêntico ao anterior. Sem SQL novo no Supabase.

---

## 2026-10-25 — TASK-078: linha viva x linha morta (coluna LV da base de orçamento)

**Tipo:** Backend + Frontend + banco · `services/orcamento_lv.py` (novo), `services/orcamento_calc.py`, `database.py`, `routers/admin.py`, `routers/orcamento.py`, `services/sync_service.py`, `services/realtime_sync.py`, `static/orcamento.html`, `data/tabela_seed.csv`, `scripts/schema_supabase.sql`

A base de orçamento ganhou a coluna **LV**: `LV` (só operações em linha viva, as que começam com `*`: `*I`, `*R`, `*M`), `LM` (só sem `*`) ou `A` (ambas). O cálculo só usa as linhas compatíveis com a operação de cada item (em todos os passos da busca e na expansão do componente); item que só existe com o outro tipo aparece em "não encontrados" com a dica. Todas as linhas existentes ficam `A` (o resultado de hoje não muda). CSV sem a coluna = `A`; valor inválido é recusado sem alterar a base. **Rodar `scripts/schema_supabase.sql` no Supabase.**

---

## 2026-10-24 — TASK-077 (etapa 6): histórico e importar/exportar dos modais montáveis

**Tipo:** Backend + Frontend · `routers/modais_montaveis.py`, `static/modal_montavel_arquivo.js` (novo), `static/modal_montavel_ui.js`, `app.py`, `static/resumo.html`

O gerenciador (＋) ganhou **Exportar JSON**, **Importar JSON** (resumo de novos/atualizados/removidos e escolha entre Acrescentar e Substituir tudo antes de gravar) e **Histórico** (últimas 20 versões com data e autor; restaurar guarda a versão atual). Só admin. Sem SQL novo (usa `atalhos_modais`).

---

## 2026-10-24 — TASK-077 (etapa 5): expressões estilo planilha nos modais montáveis

**Tipo:** Backend + Frontend · `routers/modais_montaveis.py`, `static/modal_montavel_engine.js`, `modal_montavel_editor.js`, `modal_montavel_ui.js`

Novo bloco **Cálculo** e linguagem de expressões (SE, E, OU, NAO, MAX, MIN, ABS, ARRED, SOMA, CONCAT, MAIUSCULO, MINUSCULO; comparações; texto e números) usada na quantidade das regras, nos cálculos e em `{{= … }}` nos modelos. Parser próprio, sem eval, validado no servidor. Definições antigas continuam válidas.

---

## 2026-10-24 — TASK-077 (etapa 4): colagem e interpretação dos modais montáveis

**Tipo:** Backend + Frontend · `routers/modais_montaveis.py`, `static/modal_montavel_*.js`, `static/resumo.js`

Grade com modos de colagem (colunas, linha inteira, dividir por separador, ignorar cabeçalho) e botão "Carregar da tabela Outros/Cabos" (modais de processamento). Regras com "todas as ocorrências", normalização de maiúsculas e aviso das linhas que nenhuma regra reconheceu. Sem mudança de banco.

---

## 2026-10-24 — TASK-077 (etapa 3): condicionais dos modais montáveis

**Tipo:** Backend + Frontend · `routers/modais_montaveis.py`, `static/modal_montavel_engine.js`, `static/modal_montavel_editor.js`, `static/modal_montavel_ui.js`

Condições compostas (E / OU / NÃO, até 4 níveis) com operadores `contem`, `maior`, `menor`, `maior_igual`, `menor_igual`; regras de interpretação com `se` (só valem quando a condição dos campos é verdadeira); e `[campo]` na quantidade (`{1}*[fator]`) e no ativo (`P{1}-[regiao]`). O editor visual ganhou a árvore de condições. Definições antigas continuam válidas. Sem mudança de banco.

---

## 2026-10-24 — TASK-077 (etapa 2): editor visual dos modais montáveis

**Tipo:** Frontend · `static/modal_montavel_editor.js` (novo), `static/modal_montavel_ui.js`, `static/resumo.html`, `app.py`

O gerenciador (＋) ganhou a aba **Visual**: monta o modal por blocos com formulários (campos, grade, planilha interna, regras de interpretação, destino, botão, texto para copiar e condição de visibilidade), com reordenar/remover e blocos novos já válidos. A aba **JSON** continua e as duas ficam sincronizadas. Sem mudança de API ou de banco.

---

## 2026-10-24 — TASK-077 (etapa 1): modais montáveis por projeto

**Tipo:** Backend + Frontend · `routers/modais_montaveis.py`, `static/modal_montavel_engine.js`, `static/modal_montavel_ui.js`, `static/resumo.html`, `static/resumo.js`, `app.py` · `.ai/tasks/TASK-077-24-10-2026.md`

A barra de modais do Resumo ganhou o botão **＋** (só admin): cria modais definidos em JSON, por projeto. Blocos: texto, campo, grade (cola várias linhas do Excel), planilha interna, interpretação por regras (regex → ativo + quantidade, expressões aritméticas seguras, sem eval), destino (Outros/Cabos) com "Adicionar ativos", saída de texto com modelo `{{campo}}`/`{{buscar(...)}}` e Copiar, e condição `visivel_se`. Dois exemplos prontos (Ramais colados, Títulos de contrato). Guardado na tabela já existente `atalhos_modais` (sem SQL novo). Os modais atuais (Cabos, Postes, Conexões, Ramais, Vincular) não mudaram.

---

## 2026-10-20 — TASK-076: 401 leva ao login

**Tipo:** Frontend · `static/auth_fetch.js` · `.ai/tasks/TASK-076-20-10-2026.md`

Resposta 401 de qualquer `/api/*` (exceto o login) limpa a sessão do navegador e redireciona para `/login` com o aviso "Sessão expirada. Entre novamente." (uma vez por página). Antes a tela seguia "logada" e as chamadas falhavam em silêncio (causa do "sem regras" online). Validade do token mantida em 24 h. Conferir `JWT_SECRET` fixo nas variáveis do Render.

---

## 2026-10-20 — TASK-074/075: admin online e regras do leitor sincronizadas com a nuvem

**Tipo:** Backend + Frontend · `routers/auth.py`, `routers/regras_leitor.py`, `static/script.js` · `.ai/tasks/TASK-074-20-10-2026.md`, `TASK-075-20-10-2026.md`

Online, `usuarios_nuvem` só tem `role`; o login agora considera admin quem tem `role = admin` (antes `is_admin` vinha sempre falso e os botões de admin sumiam). Leitura das regras do leitor: linha vazia na nuvem não bloqueia mais; as regras do SQLite (inclusive o seed 027/229) são usadas e enviadas ao Supabase quando a nuvem não tem. A tela distingue "falha ao carregar" de "sem regras". Quem já estava logado precisa sair e entrar de novo.

---

## 2026-10-20 — TASK-073: regras do leitor 027/229 semeadas por projeto

**Tipo:** Backend + Frontend · `database.py`, `static/script.js` · `.ai/tasks/TASK-073-20-10-2026.md`

O seed das regras de fábrica do leitor (027/229) deixou de depender de a tabela inteira estar vazia: cada projeto sem linha recebe as suas. Quando a lista está vazia e o usuário não é admin, a mensagem explica que importar/editar regras exige conta administradora.

---

## 2026-10-16 — TASK-072: SAP copiar sem descrição, seleção de células e marcação do copiado

**Tipo:** Frontend · `static/resultado_orcamento.html`, `static/orcamento_sap.js` · `.ai/tasks/TASK-072-16-10-2026.md`

Os botões **Copiar** ODI/ODD passam a levar só TIPO MDO, CÓDIGO, QUANTIDADE e CENTRO (sem descrição). Dá para **selecionar células** (arrastar ou Shift+clique) e **copiar várias de uma vez** (Ctrl+C ou "Copiar seleção"). O que foi copiado fica **marcado em teal com ✓** (some se o valor for editado; "Limpar marcações" por tabela). Exportar Excel não mudou (mantém a descrição).

---

## 2026-10-16 — SAP: nova ordem das colunas

**Tipo:** Frontend + exportação · `static/resultado_orcamento.html`, `static/orcamento_sap.js`, `services/exportar_xlsx.py`

A visão SAP (tela, Copiar e Excel) passou à ordem **TIPO MDO · CÓDIGO · QUANTIDADE · CENTRO · DESCRIÇÃO** (a REFERÊNCIA do modo comparação continua antes da quantidade).

---

## 2026-10-16 — TASK-071: modo comparação com um REC de referência

**Tipo:** Frontend · `static/resultado_orcamento.html`, `static/orcamento_comparar.js` (novo), `static/orcamento_sap.js`, `app.py` · `.ai/tasks/TASK-071-16-10-2026.md`

Além de somar e subtrair, é possível **comparar** o REC atual com outro (do banco ou de arquivo): coluna **REFERÊNCIA** (só leitura) antes de QUANTIDADE, nas visões REC e SAP; 🟢 iguais · 🟡 diferentes · 🔴 só no atual (referência = 0) · 🔵 só na referência (atual = 0, em linha-fantasma intercalada, com botão ＋ para adicionar ao atual). Barra com liga/desliga, trocar, remover, filtro por status e legenda com contagens. Salvar, exportar `.rec`, Copiar e Excel continuam só com os dados atuais. Internamente a tabela REC ganhou a coluna REFERÊNCIA (QUANTIDADE e FILTRO passaram aos índices 5 e 6).

---

## 2026-10-15 — TASK-070: visão SAP (ODI/ODD) no Resultado do Orçamento + Excel

**Tipo:** Backend + frontend · `static/resultado_orcamento.html`, `static/orcamento_sap.js` (novo), `services/exportar_xlsx.py` (novo), `routers/orcamento.py`, `requirements.txt`/`build.bat` (**nova dependência: openpyxl**) · `.ai/tasks/TASK-070-15-10-2026.md`

A tela ganhou o alternador **REC | SAP**. Na SAP há duas tabelas lado a lado, **ODI** (instalando) e **ODD** (removendo), com TIPO MDO em **M/N**, mesmos campos editáveis, filtros dinâmicos, ordenação, **Copiar** por tabela, ⇄ para mover entre elas e **Exportar Excel** (planilhas ODI e ODD; `POST /api/orcamento/exportar-xlsx`). A SAP é só uma projeção do REC: salvar, exportar `.rec`, somar/subtrair e importar PDFs seguem o fluxo do REC, e o filtro da SAP não esconde linhas do REC. Instalar a dependência: `pip install openpyxl` (já no `requirements.txt`). Coluna **CENTRO** na SAP (depois de CÓDIGO): 1600 para material, 1000 para mão de obra, derivada do TIPO MDO; vai também no Copiar e no Excel.

---

## 2026-10-15 — TASK-069: Entidade de volta em Outros e classificar/agrupar por coluna

**Tipo:** Frontend · `static/resumo.js/.html`, `static/grade_tabela.js` · `.ai/tasks/TASK-069-13-10-2026.md`

A coluna **Entidade** volta à tabela **Outros** (clique para trocar; escolher à mão trava a linha 🔒; filtro por entidade de volta; não entra em colar/copiar). Cada coluna de Outros e de Cabos ganha o menu **⇅**: Classificar A→Z / Z→A (Shift = critério secundário), Agrupar com separadores e Limpar agrupamento. A ordenação reordena de verdade, é estável (Ativo em ordem natural, Operação em ordem lógica), tem um passo de Ctrl+Z e **remapeia os vínculos cabo↔estrutura**; em Cabos a Qtd Ativos não é recalculada (aviso). A ação de IA "ordenar" usa a mesma rotina. Selects de Operação/Entidade carregam a lista só ao focar (desempenho). **Ordenar é o próprio título da coluna** (clique = crescente, de novo = decrescente; Shift = 2º critério; botão direito = agrupar) e a coluna Entidade de Outros tem um olho discreto para ocultar/mostrar.

---

## 2026-10-13 — TASK-068: modais de Cabos e Postes com atalhos, favoritos e linhas estilo Excel

**Tipo:** Backend + frontend + schema · `routers/atalhos.py`, `database.py` (tabela `atalhos_modais`), `static/atalhos_modal.js`, `static/gerador_modal.js`, `static/grade_tabela.js`, `static/resumo.js/.html`, `app.py` · `.ai/tasks/TASK-068-13-10-2026.md`

Os geradores de **Cabos** e de **Postes e Estruturas** ganharam painéis de atalhos (operação cria linha; ativo/fase/comprimento alteram a linha em destaque), botão **＋ Novo** (guarda no computador ou no projeto, na nuvem), **favoritos ★** (linha completa), modo de edição (ocultar, editar, excluir) e linhas estilo Excel com copiar/colar e Ctrl+Z próprio. Atalhos de projeto: qualquer usuário cria e altera os seus, só administrador exclui. **Supabase:** criar a tabela `atalhos_modais` (SQL na task). No Postes a entidade inserida em Outros passa a ser reclassificada (antes `OUTRO`).

---

## 2026-10-13 — TASK-063 a 067: seletor cinza, Entidade escondida, RAMAIS e grade estilo Excel

**Tipo:** Frontend · `static/grade_tabela.js` (novo), `static/resumo.js`, `static/resumo.html`, `static/styles.css`, `static/autonomo.html`, `app.py` · `.ai/tasks/TASK-063…067-13-10-2026.md`

- **063:** lista aberta dos seletores (projeto etc.) com fundo cinza `#1c1f26`.
- **064/065:** a coluna **Entidade** sai das tabelas Cabos e Outros (o dado continua nas linhas; aparece ao passar o mouse no nº da linha). Em Outros a entidade passa a ser **recalculada pelo motor do Leitor sempre que o ativo muda** (edição ou colagem); a correção manual é pelo menu de contexto (**Definir entidade…**) e **trava** a linha (`entidadeManual`, preservado em Ctrl+Z e obras salvas). Entidade que vira CABO/RAMAIS por edição não move a linha.
- **067:** Cabos e Outros viram **grade contínua**: seleção de intervalo, Ctrl+C em TSV, Ctrl+V de várias linhas/colunas do Excel (sobrescreve e cria linhas; um único Ctrl+Z), Delete, Excluir por intervalo, Enter/setas. Lógica pura e testada em `grade_tabela.js`.
- **066:** modal **RAMAIS** aceita colar várias linhas (uma por linha, tipo RAMAIS), recalcula INSTALANDO/REMOVENDO na hora, tem + Linha, seleção/cópia/Delete de linhas e guarda as edições.
Sem mudança de schema nem de API.

---

## 2026-10-09 — TASK-061: regras do leitor e ajustes em JSON (modo avançado, exportar, importar)

**Tipo:** Backend + frontend · `.ai/tasks/TASK-061-09-10-2026.md`

**Exportar JSON** / **Importar JSON** nas duas tabelas de regras do leitor e nos ajustes, e **Modo avançado (JSON)** por regra do leitor (os ajustes já tinham). A importação valida (formato, versão, limites, regras e regex, no servidor), mostra o resumo (novas / iguais / a remover) e deixa escolher **Acrescentar** ou **Substituir**; **nada é gravado**: vai para o rascunho e só vale no Salvar de sempre (com histórico e reversão).
Arquivo de outro projeto é aceito com aviso (copiar regras entre projetos); da outra tabela é recusado. Sem mudança de schema. Manual atualizado.

---

## 2026-10-09 — TASK-059: Montar Orçamento também alimenta o aprendizado

**Tipo:** Backend + frontend · `.ai/tasks/TASK-059-08-10-2026.md`

Além do **Salvar obra**, o **Montar Orçamento** passa a registrar as correções da sessão (obra do autônomo ou trabalho manual) para o aprendizado, sem salvar obra: `POST /api/aprendizado/sessao`. Só depois de passar a checagem de contrato; salvar e montar na mesma sessão contam uma vez só.

---

## 2026-10-09 — TASK-059 (etapa 3): nível de confiança do autônomo

**Tipo:** Backend + frontend · `.ai/tasks/TASK-059-08-10-2026.md`

Novo painel "Confiança do autônomo" em `/autonomo`: por tipo de arquivo, das últimas N (10) obras que você revisou, quantas ficaram sem correção (confiável a partir de 90%; recomeça quando você aprova uma proposta). Opcional e **desligada por padrão**: com a opção ligada num projeto e o tipo confiável, as exclusões que sempre pediam confirmação passam a ser confirmadas sozinhas
(exceto se forem anormalmente muitas — mais de 25% das linhas); fica registrado no relatório e dá para reverter. Janela e taxa configuráveis na tela de configuração. Sem mudança de schema.

---

## 2026-10-02 — TASK-031: achados do primeiro teste real (projeto 027/229)

**Tipo:** Backend (modo autônomo)

(1) Ajustes habilitados do projeto 027 somavam **53 ações** e o limite de 50 (válido nas rotas/tela) barrava o arquivo inteiro no autônomo → `ajustar(..., limite_acoes)` e `validar_acoes(..., limite)` ganharam parâmetro;
o pipeline usa `LIMITE_ACOES_AUTONOMO = 300`; rotas e tela continuam com 50. (2) A mensagem de "tipo de arquivo não suportado" agora mostra a extensão encontrada e o nome do arquivo.

---

## 2026-10-02 — INCIDENTE: testes rodaram contra o Supabase real (corrigido) — `docs/INCIDENTE_TESTES_NA_NUVEM_2026-10-02.md`

**Tipo:** Testes / segurança de dados

Ao rodar `pytest` no PC do usuário (com `.env` do Supabase), os testes usaram a nuvem REAL e gravaram regras/ajustes/prompts/obras de teste (DEFAULT, 229, 027; usuários `@x.com`, `u1`, `u2`).
Causa: os testes assumiam "sem credenciais" (Linux de desenvolvimento) e nada garantia isso. Correção: `tests/conftest.py` zera `SUPABASE_*`/chaves de IA, usa banco temporário por teste e recusa rodar
(`pytest.exit`) se houver cliente Supabase; `tests/test_isolamento.py` confere. Roteiro de verificação/recuperação (via histórico) no documento do incidente. **Regra nova: teste nunca toca a nuvem; precisa de nuvem → cliente falso.**

---

## 2026-10-02 — TASK-031: preparação para o Windows (trabalhador por socket, arquivo em uso, roteiro de testes)

**Tipo:** Backend + scripts + docs · `.ai/tasks/TASK-031-02-10-2026.md`

O processo auxiliar do leitor passou a falar por **socket local com token** (no `.exe` sem console não há stdin/stdout) e abre sem janela no Windows; `app.py` repassa `<porta> <token>`.
A vigia **adia** o movimento de arquivo em uso (PermissionError) em vez de reprocessar/errar. `build.bat` inclui `quickjs`. Novos: `docs/ROTEIRO_TESTE_WINDOWS_MODO_AUTONOMO.md`,
`scripts/autonomo_verificar_ambiente.py`, `scripts/autonomo_arquivos_teste.py`.

---

## 2026-10-02 — TASK-031 fase D: tela de controle do modo autônomo (`/autonomo`)

**Tipo:** Frontend + correções no backend do modo autônomo · `.ai/tasks/TASK-031-02-10-2026.md`

Página `/autonomo` (só admin; link no Painel Admin): estado, ligar/desligar, varrer agora, **confirmação de exclusões (Sim / Não / Sim para todos / Não para todas)**, histórico com
detalhes, reverter e reprocessar, e configuração (dono das obras, pastas, intervalo, espera). Sem biblioteca externa; só `textContent`. Correções achadas no teste real: varreduras
simultâneas (agora serializadas) e falhas pré-pipeline passam a aparecer no histórico. O modo continua **desligado por padrão**.

---

## 2026-10-02 — TASK-031 fase C: pasta monitorada, fila e API de controle do modo autônomo

**Tipo:** Backend + `app.py` · `.ai/tasks/TASK-031-02-10-2026.md`

`services/autonomo/pasta.py` (vigia em thread: `entrada/<projeto>/` → pipeline, um por vez, só arquivos estáveis; originais para `processados/` ou `erros/` + `.erro.txt`),
`config_autonomo.py` (chave `autonomo_config` em `configuracoes`) e `routers/autonomo.py` (`/api/autonomo/*`, só admin: config, ligar/desligar, status, varrer, histórico,
confirmar/rejeitar/reverter/reprocessar). `app.py` ganhou: argumento `--trabalhador-leitor-js` (subprocesso do QuickJS no `.exe`), include do router e hooks `startup`/`shutdown`
que religam/param a vigia conforme a configuração. Desligado por padrão: nada muda até um admin ligar.

---

## 2026-10-02 — TASK-031 fase B: pipeline autônomo (aplicador de diff, Totalizadora, orçamento, obra e pasta)

**Tipo:** Backend (código novo, ainda sem rota/tela) · `.ai/tasks/TASK-031-02-10-2026.md`

`services/autonomo/` ganhou `aplicar.py` (porte de `aplicarOperacoesAjuste`), `totalizadora.py` (porte de `syncTotalizadora` + payload do orçamento),
`pipeline.py` (ler → montar → validar → ajustar → confirmação de exclusões → Totalizadora → orçamento → obra + pasta), `execucoes.py` (tabela local
`execucoes_autonomas`, criada sob demanda; sem Supabase) e `saida.py` (JSON + CSV). Exclusões (`excluir_linhas`, `remover_ativo`) deixam o arquivo
"aguardando_confirmacao"; `confirmar` (Sim / Sim para todos), `rejeitar`, `reverter`. O leitor roda em subprocesso (`trabalhador_js.py`).
`ramais.py` gera os ativos dos RAMAIS (decisão: entram no orçamento autônomo; etapa `ramais` do pipeline). **Alterar `aplicarOperacoesAjuste`, o modal RAMAIS, `syncTotalizadora`, `obterPayloadCalculo` ou a montagem em `resumo.js` exige rodar `tests/test_autonomo_*.py`** (o oráculo fatia o código real).

---

## 2026-10-02 — TASK-031 fase A: leitor via QuickJS e montagem em Python (modo autônomo)

**Tipo:** Backend (código novo, não carregado pelo app por padrão) · `.ai/tasks/TASK-031-02-10-2026.md`

`services/autonomo/leitor_js.py` executa o **mesmo** `static/regras_leitor_engine.js` num QuickJS em **processo separado** (a regex catastrófica
não é interrompida pelo limite de tempo do QuickJS; o processo principal mata e recria o trabalhador). `services/autonomo/montagem.py` porta a
montagem de Cabos/Outros/Ramais e o cálculo de quantidades de `resumo.js`. Paridade provada contra o JS real (Node) em `tests/test_autonomo_*.py`.
**Mudar `regras_leitor_engine.js` ou as funções de montagem em `resumo.js` exige rodar esses testes** (marcadores fatiados em `tests/oraculo_tela.py`).
Dependência nova: `quickjs`. Nenhuma mudança no modo manual.

---

## 2026-10-02 — TASK-030: manual de uso de regras e ajustes

**Tipo:** Documentação do usuário (backend leve + frontend) · `.ai/tasks/TASK-030-02-10-2026.md`

Manual em Markdown (`data/manual_regras_e_ajustes.md`) servido por `GET /api/manual` (login, qualquer perfil) e renderizado em `/manual`;
botão **Manual** na gaveta Regras de validação. Os exemplos JSON do manual são validados em `tests/test_manual_exemplos.py`
(ao mudar a linguagem de regras/ajustes, o teste acusa exemplo desatualizado — atualize o manual junto).

---

## 2026-10-02 — TASK-029: botão Ajustar executa todos os ajustes habilitados

**Tipo:** Ajustes (backend + frontend) · `.ai/tasks/TASK-029-02-10-2026.md`

`Ajustar` deixou de ser menu: abre um modal com um **cartão por ajuste habilitado que muda algo** (caixa por mudança, **Executar correção**)
e, acima, **Executar todas as correções** (em cadeia, ordem do cadastro, um passo de histórico). Nova rota `POST /api/validacao/ajustes/preview-lote`
(só simula; uma chamada devolve cartões, ajustes sem mudança, configs inválidas e a cadeia; limites 30 ajustes / 50 ações na cadeia).
Exclusões pedem confirmação extra; guarda da camada 1 e proteção de linha alterada mantidas. Após executar, os cartões são recalculados.

---

## 2026-10-01 — TASK-028: IA do chat consulta e comanda as obras salvas

**Tipo:** IA do chat (backend + frontend) · `.ai/tasks/TASK-028-01-10-2026.md`

Quando o pedido menciona **obra/obras**, o backend acrescenta ao prompt um **índice enxuto** (id | nome | data, até 50, só do usuário e
do projeto selecionado, sem `dados_json`) e o bloco `data/prompt_obras.txt` ao system prompt; o **conteúdo** de uma obra só vai
se o usuário pedir ("o que tem na obra X?") e só dela (até 2, teto de linhas). **Pedidos sem a palavra "obra" seguem idênticos
a antes** (nenhuma consulta ao banco, mesmo prompt). Comandos: a IA responde `{"acao_ui":"obra","acao":"adicionar|subtrair|carregar",
"obra_id":…}`; o frontend valida o id contra as obras reais (`GET /api/obras/{id}`), exige que a obra seja do projeto selecionado e
**sempre abre um popup de confirmação** (carregar avisa que substitui tudo); aplicar é um passo de histórico. A IA não salva nem
exclui obras. O Fast-Path não intercepta frases com "obra". Refatoração: `adicionarObraAoProjeto`/`subtrairObraDoProjeto`
compartilhadas pelo modal e pela IA; o modal Carregar Obra deixou de usar `innerHTML` com o nome da obra (XSS armazenado).

---

## 2026-10-01 — TASK-027: microfone no chat de IA do Resumo

**Tipo:** frontend (Resumo) · `.ai/tasks/TASK-027-01-10-2026.md`

Novo `static/voz.js`: botão 🎤 ao lado de enviar dita em pt-BR pelo reconhecimento de voz do **navegador** (Web Speech API); o
texto aparece ao vivo no campo (anexado ao que já estava) e **nunca é enviado sozinho**; para por clique, Esc ou 4 s de silêncio;
erros (permissão, sem microfone, sem rede, sem fala) com mensagem clara. **Leitura da resposta em voz alta** (`speechSynthesis`),
**desligada por padrão** (🔇/🔊, `localStorage` `chat_voz_resposta`), com "🔊 Ouvir" por resposta e ⏹ para parar (não lê blocos de
código nem tabelas). Sem suporte (ex.: janela de desktop): botão desabilitado com explicação; o chat digitado segue igual. Sem
backend, sem custo, sem guardar áudio.

---

## 2026-10-01 — TASK-026: controles de validação embutidos na gaveta lateral

**Tipo:** frontend (Resumo) · `.ai/tasks/TASK-026-01-10-2026.md`

A barra do Resumo ficou só com `Regras de validação`, `Ajustar ▾` e `Validar` (além de Montar Orçamento). As caixas de modo
(Determinística, Regras, IA, Pular IA se houver erro de contrato) e o **Validar ao montar** foram para a nova aba **Execução**,
a primeira da gaveta e visível também para o operador (são preferências locais, mesmas chaves de `localStorage`, valem sem a
gaveta aberta). Ao lado do `Validar` há um resumo do que vai rodar (ex.: "Det + Regras + IA · ao montar", com dica) e um
atalho ⚙ que abre a gaveta na aba Execução. `resumo.js` expõe `window.validacaoPrefs`; sem mudança de backend.

---

## 2026-10-01 — TASK-025: ajustes por IA com ações estruturadas

**Tipo:** camada 3 (IA) + frontend · `.ai/tasks/TASK-025-01-10-2026.md`

Novo prompt semente **`ajustar-planilhas`** (editável; entra sozinho na próxima inicialização, sem sobrescrever prompts já
editados) e rota `POST /api/validacao/ajustes-ia`: a IA propõe **ações** da linguagem de ajustes (inclusive **excluir**),
o backend valida cada ação (schema + grupos do projeto), **recusa `operacao_nova`**, deduplica, limita a 10 e roda o motor da
TASK-022 nas tabelas completas, devolvendo `acoes` (com frase em português) + `diff` + `destrutivo`. O botão **Pedir ajuste à IA**
do painel de validação agora abre a pré-visualização com aviso em vermelho quando há exclusão, confirmação extra ao aplicar e
**Cadastrar como ajuste** (admin): abre a aba Ajustes com as ações em rascunho. A rota antiga `/corrigir` (substituição de linha
inteira) continua existindo, mas a tela não a usa mais. Sem mudança de banco.

---

## 2026-10-01 — TASK-024: fluxo unificado Validar + Ajustar

**Tipo:** frontend (Resumo) · `.ai/tasks/TASK-024-01-10-2026.md`

O painel de validação agora oferece **Ajustar** por achado (ajustes cadastrados e ligados que corrigem a regra do achado),
**Ajustar tudo (determinístico)** e **Pedir ajuste à IA** (só para achados sem ajuste cadastrado), com filtro
Todos · Determinístico · IA. Todo ajuste passa por uma **pré-visualização** (editar / inserir / EXCLUIR / reordenar, com caixa
por mudança; nada é aplicado sem aceite). Aplicar é **um passo de histórico** (um Ctrl+Z desfaz o lote), ignora mudança cuja
linha foi alterada desde a pré-visualização e **revalida automaticamente**, mostrando "antes do ajuste". No Validar ao
montar, o ciclo painel → ajuste → revalidação continua até o usuário decidir; se sobrar só info, segue para o orçamento.
Novo botão **Ajustar ▾** na barra roda um ajuste cadastrado avulso. "Não, vou corrigir" deixou de chamar a IA sozinho.

---

## 2026-10-01 — TASK-023: cadastro de ajustes recorrentes por projeto (receitas)

**Tipo:** persistência + frontend admin · `.ai/tasks/TASK-023-01-10-2026.md` · **exige SQL no Supabase** (`scripts/schema_supabase.sql`)

Receitas de ajuste (ações da TASK-022 com nome, ligada/desligada e vínculo a regras de validação) agora são **cadastradas**:
`DEFAULT` guarda o container completo e cada projeto só o overlay (adicionadas/sobrescritas/ocultas), como as regras
(`services/ajustes_camadas.py`). Tabelas novas `ajustes_planilhas` e `ajustes_planilhas_historico` no SQLite e no Supabase
(nuvem primeiro, fallback local; falha de sincronização avisada) — gravação **só no botão Salvar**, com histórico e
reversão. Semente (todas desligadas): normalizar postes, `SUP-L→SUPL`, adicionar `1-SUPL` à linha com CFU (vinculada a
`C2-CFU-SUPL`), excluir linhas vazias, ordenar Outros. Rotas `GET/POST /api/validacao/ajustes` (+histórico, reverter,
restaurar-semente); `preview` aceita `receitas: [ids]`. Nova aba **Ajustes** na gaveta (`painel_ajustes.js`): selos, −/+,
voltar ao padrão, editor visual das ações, modo JSON e **pré-visualização do diff nas tabelas atuais** (nada é aplicado:
isso é a TASK-024). Operador vê a lista e pré-visualiza; só admin edita.

---

## 2026-10-01 — TASK-022: motor de ajustes determinísticos (ADR-006)

**Tipo:** novo motor de regra de negócio (backend) · `.ai/tasks/TASK-022-01-10-2026.md`

Novo `services/ajustes_planilhas.py`: 7 ações declarativas sobre Cabos/Outros (substituir texto/item, normalizar, ordenar,
excluir linhas, adicionar linha, adicionar/remover ativo na linha, mesclar duplicadas), reaproveitando condições, seletores e
grupos da linguagem de regras v2. Devolve um **diff** (editar/inserir/excluir/mover) sem alterar a entrada; a camada 1 descarta
ajuste que gere erro de contrato novo; a operação da linha só muda com `operacao_nova`; idempotente. Rotas
`POST /api/validacao/ajustes/preview` e `/descrever`. Sem mudança de banco. Telas e cadastro: TASK-023/024.

---

## 2026-10-01 — TASK-019: painel de regras de validação vai do Admin para o Resumo

**Tipo:** frontend (reorganização) · `.ai/tasks/TASK-019-01-10-2026.md`

O editor das regras de domínio (TASK-017/018) e os prompts de validação saíram de `/admin` e passaram a uma **gaveta
lateral** no Resumo (botão **Regras de validação**), com abas Regras · Grupos · Prompts da IA · Testar, que segue o projeto
de `#select-projeto` (o padrão de todos é uma opção da gaveta). **Só admin edita**; operador vê Regras e Grupos em leitura
(frases, selos). Reestilizado sem Tailwind (`painel_regras.css`); JS carregado sob demanda. Mesmos dados e rotas, nada perdido;
o Admin mostra só um aviso de mudança. Prompts e Testar são abas só de admin.

---

## 2026-10-01 — TASK-020: validação em modos independentes (determinística e IA)

**Tipo:** comportamento do frontend · `.ai/tasks/TASK-020-01-10-2026.md`

Barra do Resumo ganhou quatro caixas (preferência **local**, `localStorage` `validacao_modos`): **Determinística**
(contrato + regras), **Regras** (liga/desliga a camada 2; o contrato fica sempre com a determinística), **IA** e **Pular IA se
houver erro de contrato** (padrão ligado). Valem para o botão Validar e para "Validar ao montar". A IA roda **depois** da
determinística; com IA desligada nenhuma chamada a `/api/validacao/ia` é feita; com tudo desligado o Validar avisa "nada a
validar" e o orçamento segue. O painel diz o que rodou. "Corrigir com IA" só aparece com a IA ligada. Sem mudança de backend
(usa `incluir_dominio`).

---

## 2026-10-01 — TASK-021: desfazer/refazer sem desvio

**Tipo:** correção de comportamento (frontend) · `.ai/tasks/TASK-021-01-10-2026.md`

`static/resumo.js`: `undo()` passa a gravar o estado ao vivo quando ele está à frente do ponteiro do histórico e
`pushHistory()` ignora estado repetido (`historyDirty` controla o botão Desfazer). Funciona para os chamadores que
empilham antes de mudar e para os que empilham depois. Duas edições seguidas agora exigem dois Ctrl+Z; a correção
assistida deixou de precisar do "push duplo". Alterações em lote (ajustes das próximas tasks) serão um passo cada.

---

## 2026-09-30 — TASK-018 + TASK-017: regras em camadas por projeto e editor visual

**Tipo:** armazenamento/regra de negócio + frontend admin · `.ai/tasks/TASK-018-30-09-2026.md`, `TASK-017-30-09-2026.md`

**Camadas (TASK-018):** o `DEFAULT` guarda regras e grupos completos; cada projeto guarda só o **overlay**
(`adicionadas`, `sobrescritas`, `ocultas`, `grupos`) no mesmo `regras_json` — **sem SQL novo**. Efetivo = padrão + overlay,
vale a sobrescrita do projeto; o que muda no padrão chega aos projetos. Cópias integrais antigas (lista v1 ou container v2)
são lidas como overlay equivalente (o que faltava no padrão vira "oculta"), sem mudar resultados; virarão overlay no
próximo salvamento. Novo `services/regras_camadas.py` (funções puras) e `GET /api/validacao/regras` devolve `origem`,
`oculta`, `frase`, `grupos_origem`, `grupos_usados`, `avisos`. `restaurar-semente` num projeto = descartar os ajustes;
`adicionar-novas` passou a valer só no `DEFAULT`. Nova `GET /api/validacao/regras/ativos` (autocomplete).

**Editor visual (TASK-017):** novo `static/regras_editor.js` substitui o editor JSON da TASK-013: cartão em frase (vinda
do backend), selos *padrão / ajustada neste projeto / só deste projeto*, botão **−/+** (oculta/reexibe regra herdada),
"voltar ao padrão", construtor SE/ENTÃO/EXCETO SE com E/OU/NÃO e comparador de quantidade, autocomplete de ativos e
`@GRUPO`, editor de grupos com "usado por", assistente com 4 modelos, modo avançado (JSON) e painel de teste que explica
também por que uma linha **não** disparou.

**Correção:** os seletores de projeto do painel admin (prompts e regras) nunca listavam projetos, porque `/api/projetos`
devolve `{projetos: [...]}` e o código esperava uma lista; agora aceitam os dois formatos.

---

## 2026-09-30 — TASK-016: motor de regras de domínio v2 (ADR-005)

**Tipo:** mudança de motor de regras (backend) · `.ai/tasks/TASK-016-30-09-2026.md`

Regras de domínio passam a uma linguagem declarativa v2: condições sobre **quantidade** (`=, ≠, ≥, ≤, >, <, entre`), grupos
de ativos nomeados (`@ESTRUTURA_MT`), combinadores E/OU/NÃO, valores dinâmicos, regras sobre a **planilha inteira** (P50 e
contagens) e `explicacao` em português em cada achado; `descrever(regra)` gera a frase. Regras v1 salvas continuam
funcionando (conversão determinística); o resultado das 11 regras da semente é **idêntico** ao do motor anterior, provado
contra uma cópia congelada dele em milhares de cenários. Sem SQL novo (mesmo `regras_json`, agora um container com grupos).
Painel de achados mostra "Por quê". Camadas por projeto e editor visual: TASK-018 e TASK-017.

---

## 2026-09-30 — Regras de domínio pendentes (`C2-TR-EF`, `C2-BT-EXT`) e "adicionar regras novas da semente"

**Tipo:** regras de negócio (decisão do usuário) + ação do admin · `.ai/tasks/TASK-013-30-09-2026.md` (acréscimo)

`C2-TR-EF`: trafo com elo fusível só é aceito com chave `CFU` ou `CFUR` no poste (`CFA`, `CL` e as de reinstalação
não liberam). `C2-BT-EXT`: estrutura `SI` exige `RA2`, salvo estrutura `S#`. Ambas na semente, desligadas, aviso,
operações I/*I. O tipo `proibe` ganhou `exceto_se_regex`. Nova rota admin `POST /api/validacao/regras/adicionar-novas`
(botão no painel): acrescenta à versão do projeto só as regras da semente que faltam, sempre desligadas, sem tocar nas
existentes — necessária porque a semente só carrega quando não há versão DEFAULT. **Fecha o catálogo de validações.**

---

## 2026-09-30 — TASK-015: correção assistida por IA com aceite explícito

**Tipo:** nova funcionalidade (backend + frontend) · `.ai/tasks/TASK-015-30-09-2026.md`

`POST /api/validacao/corrigir`: a IA propõe correções só para as linhas citadas nos achados; o servidor reconfere
cada proposta (linha existente, muda algo, **passa de novo na camada 1**) e devolve também as descartadas com o
motivo. Na tela, "Não, vou corrigir" e o botão "Corrigir com IA" abrem as propostas (antes → depois) com aceite por
linha ou "Aceitar todas"; aplicar é um único passo no histórico e ignora linhas que mudaram. A operação nunca é
alterada. Refatorações mínimas: laço de lotes compartilhado (`executar_em_lotes`) e preâmbulo da IA (`_preparar_ia`).
O prompt `corrigir-planilhas` perdeu a frase sobre alterar a operação (instalações existentes: "Restaurar semente").
**Fecha o conjunto ADR-004 (TASK-009 a TASK-015).**

---

## 2026-09-30 — TASK-014: validação na tela (botão, opção automática) e revisão por IA (camada 3)

**Tipo:** nova funcionalidade (backend + frontend) · `.ai/tasks/TASK-014-30-09-2026.md`

Botão **Validar** e opção **Validar ao montar** (radio Não/Sim, `localStorage`, padrão Não) na aba Resumo, com painel
de achados das 3 camadas e o diálogo "continuar mesmo assim?" (só erro/aviso interrompem). `POST /api/validacao/ia`
usa o prompt salvo (TASK-012), lotes de 40 linhas, resposta JSON validada (descarta linha inventada) e nunca derruba
as camadas 1-2: falha da IA vira aviso. Refatoração mínima: `obterPayloadCalculo()` extraída do handler de "Montar
Orçamento" (comportamento igual). `localStorage['processar_dados']` e `orcamentoPayload` inalterados. O botão
"Não, vou corrigir" ainda só fecha o painel (a correção é a TASK-015).

---

## 2026-09-30 — TASK-013: regras de domínio da validação (camada 2), editáveis pelo admin

**Tipo:** nova funcionalidade (backend, banco, painel admin) · `.ai/tasks/TASK-013-30-09-2026.md`

Motor `services/regras_dominio.py` com 5 tipos de regra declarativa, avaliadas por linha da tabela Outros
(P50 pelo total da planilha). Regras por projeto com fallback `DEFAULT`, histórico e reversão, painel de teste
do rascunho e editor no painel admin. Semente com 9 regras derivadas do `prompt_rede_eletrica.txt §5`
(intocado), **todas desligadas**; valem só para operações I/*I. As 3 regras de P50 do catálogo viraram uma
(a exigência é uma soma). `POST /api/validacao/planilhas` passa a incluir a camada 2.
**Ação necessária:** rodar o trecho novo de `scripts/schema_supabase.sql` no Supabase real.

---

## 2026-09-30 — TASK-012: prompts de validação editáveis pelo admin

**Tipo:** nova funcionalidade (backend, banco, painel admin) · `.ai/tasks/TASK-012-30-09-2026.md`

Prompts de validação (camada 3 do ADR-004) viram dados: semente em `data/validacoes/*.md`, tabelas
`prompts_validacao`/`prompts_validacao_historico` (SQLite + Supabase), rotas `/api/validacao/prompts`
(leitura autenticada, escrita admin), validação de cabeçalho e placeholders antes de salvar, histórico
com reversão e restauração da semente. Versão por projeto com fallback `DEFAULT`. UI no painel admin.
**Ação necessária:** rodar o trecho novo de `scripts/schema_supabase.sql` no Supabase real. Contratos da
Regra 5: nenhum alterado; `prompt_rede_eletrica.txt` intacto.

---

## 2026-09-30 — TASK-011: camada 1 da validação das planilhas (primeiros testes do projeto)

**Tipo:** nova funcionalidade (backend) + refatoração mínima · `.ai/tasks/TASK-011-30-09-2026.md`

`POST /api/validacao/planilhas` valida o contrato de Cabos e Outros (formato, operação, quantidade,
duplicidade e, opcionalmente, existência do ativo na base sobre o payload de cálculo). Para reaproveitar
o parser sem duplicá-lo, `orcamento_calc.py` teve o parsing por linha extraído em
`extrair_ativo_cabo`/`extrair_ativos_outros`/`tokenizar_outros` — saída do cálculo comprovadamente
idêntica. **Achado:** a tabela Cabos da tela (`CAA 2 ABC 35 m`) e o payload de cálculo (`CAA2 1 35`) têm
formatos diferentes; a validação de Cabos espelha a normalização do frontend (nova duplicação
conhecida). `ADR-004` passa a ACEITA. Suíte `pytest` criada (`tests/`, dependências em `requirements-dev.txt`).

---

## 2026-09-30 — TASK-009: chave de IA padrão do sistema

**Tipo:** nova funcionalidade + correção (backend/frontend/config) · `.ai/tasks/TASK-009-30-09-2026.md`

O chat de IA passa a funcionar sem o usuário digitar chave. `resolver_credencial()` define a
precedência **usuário > salva > padrão** (`GEMINI_API_KEY`/`GOOGLE_API_KEY`). Corrige um bug: a UI
enviava o texto `SAVED_IN_BACKEND` quando o campo estava vazio, e por ser "verdadeiro" em Python ele
impedia que a variável de ambiente fosse consultada — resultado 401 mesmo com chave no servidor.
No desktop, `config.py` carrega o `.env` ao lado do `.exe`. Rate-limit por usuário só com a chave
padrão. `/api/health` ganha `ai_key_source`. Contratos da Regra 5: nenhum alterado.

## 2026-09-30 — TASK-010 / ADR-004: base do sistema de validação das planilhas

**Tipo:** documentação e sementes (sem lógica) · `.ai/tasks/TASK-010-30-09-2026.md`

Decisão de validar Cabos e Outros em três camadas (contrato em código, regras de domínio editáveis
pelo admin, IA com prompts salvos editáveis) — `ADR-004` (PROPOSTA). Criados o catálogo
`.ai/VALIDACOES.md`, os prompts iniciais em `data/validacoes/` e a semente das regras de domínio,
todas inativas até confirmação. `prompt_rede_eletrica.txt` não foi alterado. Achado documentado: o
parser de `orcamento_calc.py` assume valores e descarta tokens em silêncio.

---

## 2026-09-29 — Correção: regra de CERCA ("FIOS") sempre saía com operação R quando a cor era cinza

**Tipo:** correção de dado (seed) · fora do escopo de uma TASK, reportado pelo usuário via edição
manual no admin

A regra de Classificação que atribui `entidade: CERCA` a partir do texto "FIOS"
(`data/regras_leitor_classificacao_seed.json`, ordem 40) estava sem o campo
`operacao_ajustada: "I"` — o mesmo campo que todas as regras irmãs (RAMAIS, APOIO, IP) usam para
forçar a operação. Sem ele, a operação de uma linha de cerca ficava com o que a Tabela de
Processamento decidisse (a regra de "FIOS" lá também não define operação), caindo no fallback por
cor (`operacaoPelaCor`: vermelho→I, **cinza→R**, resto→M) — cercas desenhadas em cinza saíam como
"R" em vez de sempre "I". A própria `.ai/tasks/TASK-006-25-09-2026.md` já documentava que CERCA
deveria forçar operação I junto com RAMAIS/APOIO/IP; ficou de fora só dessa regra no seed, uma
omissão de transcrição.

Corrigido o seed (`operacao_ajustada: "I"` adicionado). O usuário já havia corrigido o dado no
banco em produção diretamente pelo editor de regras (TASK-007) antes desta correção do seed —
esta mudança só evita que o mesmo bug volte a aparecer em uma instalação nova/reset do banco.

---

## 2026-09-29 — TASK-007: regras do leitor ficam editáveis pela UI (admin)

**Tipo:** nova funcionalidade (backend + frontend) · **Tarefa:** `.ai/tasks/TASK-007-29-09-2026.md`

O modal "Regras do Leitor" (TASK-006, antes só visualização) ganha edição completa das duas
tabelas de regras, restrita a usuários com role `admin`: formulário estruturado por regra (campos
condicionais por modo — `DEFINIR`/`SUBSTITUIR`/`SUBSTITUIR_TOTAL` — e fase), reordenação por
drag-and-drop, painel de teste que roda `RegrasLeitorEngine` no navegador contra o conjunto de
regras em edição (ainda não salvas, sem round-trip ao backend), validação de schema no backend
antes de salvar (toda regex precisa compilar, campos obrigatórios por modo — nunca aceita um
payload que quebraria o motor), e histórico de versões com reversão (cada salvamento preserva a
versão anterior; reverter também vira uma nova entrada de histórico, a cadeia nunca perde uma
versão). Rotas `POST` protegidas por `Depends(require_role("admin"))`; `GET` continua aberto a
qualquer usuário autenticado.

**Correção de planejamento descoberta na implementação:** o refinamento inicial da tarefa assumia
que a ordem de execução das regras era a posição delas no array/JSON — na verdade o motor
(`regras_leitor_engine.js`) ordena por um campo `ordem` explícito dentro de cada fase. O
drag-and-drop reescreve esse campo (múltiplos de 10, mesmo espaçamento do seed) em vez de só
reordenar o array.

Novas tabelas `regras_leitor_processamento_historico`/`_classificacao_historico` (SQLite +
Supabase). Duplicação de lógica registrada como risco conhecido: a validação de schema do backend
precisa se manter em sincronia com o que `regras_leitor_engine.js` de fato interpreta.

---

## 2026-09-29 — TASK-008: painel "Resumo da rede" na aba Resumo

**Tipo:** nova funcionalidade (frontend) · **Tarefa:** `.ai/tasks/TASK-008-29-09-2026.md`

Adicionado um painel com 4 contagens em tempo real na aba Resumo, logo abaixo da barra de botões
dos modais geradores (Postes/Cabos/Conexões/Ramais/...): **postes instalando** (soma de quantidade
dos tokens `qtd-ativo` da tabela Outros que começam com `DT`/`CV`), **rede de média instalando**
(soma do comprimento bruto dos Cabos que começam com `CAA`/`P`/`CAL`), **rede de baixa instalando**
(idem para `M2X`/`M3X`) e **equipamentos instalando** (soma de quantidade dos tokens Outros que
começam com `TR`/`CFU`/`CFA`). Todas as 4 consideram só linhas com operação `I`.

**Duplicação de parsing (risco registrado):** `static/resumo.js` ganhou uma porta em JavaScript do
parsing `qtd-ativo` (Outros) e `ATIVO FASE COMPRIMENTO` (Cabos) que hoje só existia em Python
(`services/orcamento_calc.py`), necessária para recalcular em tempo real no navegador sem round-trip
ao backend. **O comprimento usado nunca é multiplicado pela quantidade de ativos/fases** —
diferente do cálculo de orçamento, que multiplica — decisão explícita do usuário.

Não persiste em banco — é só um cálculo em cima de `tableStates.cabos.data`/`outros.data` já
carregados na tela, atualizado a cada render e a cada edição direta (mudança de operação/entidade,
edição do texto do ativo).

Junto: registrada (não implementada) `.ai/tasks/TASK-007-29-09-2026.md` — tornar as regras do
leitor (TASK-006) editáveis via UI para usuários `admin`, em stand-by até refinamento posterior.

---

## 2026-09-25 — TASK-006: classificação do leitor vira motor de regras editável por projeto

**Tipo:** nova arquitetura (motor de regras) · **Tarefa:** `.ai/tasks/TASK-006-25-09-2026.md`

A classificação automática do leitor (entidade/operação/ativo a partir de texto/cor/layer),
antes embutida em `processAtivoFormula`/`computeRowLogic`/`updateRowLogic`/`autoClassifyEntidade`
(duplicada e já divergente entre `script.js` e `resumo.js`), virou duas tabelas de regras por
projeto — **Processamento** (texto/cor/layer → operação/ativo, em 3 fases: seleção de ramo,
pós-processamento, overrides por cor/layer) e **Classificação** (operação/ativo/cor/layer/texto →
operação/entidade, reutilizada também em `resumo.js`) — interpretadas por um motor genérico
(`static/regras_leitor_engine.js`).

**Decisão de unificação:** o comportamento das duas abas foi unificado usando `script.js` (o mais
completo) como canônico. Dois bugs reais do código original foram corrigidos por decisão do
usuário em vez de replicados: elo fusível nunca classificava como `CHAVE` (variável `uAtivo`
desatualizada) e a detecção de cabo multiplexado `M3x1` nunca funcionava (regex case-sensitive
contra texto já maiúsculo).

**Alterado/criado:**
- `static/regras_leitor_engine.js` (novo) — motor de regras, roda em navegador e Node
- `data/regras_leitor_{processamento,classificacao}_seed.json` (novos) — transcrição fiel de
  `script.js` (45 + 22 regras), semeada automaticamente para os projetos `027`/`229`
- `database.py`, `scripts/schema_supabase.sql` — duas tabelas novas por projeto
- `routers/regras_leitor.py` (novo) — CRUD nuvem-primeiro, mesmo padrão de `regras_conversao`
- `static/script.js`, `static/resumo.js` — cascata antiga removida, religada ao motor; botão
  "Regras do Leitor" (visualização) + aviso/escolha de projeto alternativo quando o projeto
  selecionado não tem regras cadastradas

**Validado:** porta fiel do código original para Node.js ("oráculo"), comparada contra o motor
novo em 66 casos sintéticos (63 idênticos, 3 divergências aprovadas — os bugs corrigidos, 0
falhas não explicadas); servidor real + navegador real via Playwright confirmando que o motor
carrega as regras pela API de verdade e produz os mesmos resultados dentro do navegador.

**Não corrigido nesta etapa:** `.ai/CONTEXT.md`/`.ai/ARCHITECTURE.md` ainda descrevem a
classificação como lógica fixa em código (desatualizado); nenhuma UI de edição das regras foi
construída (só visualização, que era o que o critério de aceite pedia).

---

## 2026-09-25 — TASK-005: obras, RECs e regras de IA isolados por projeto

**Tipo:** schema + backend + frontend · **Tarefa:** `.ai/tasks/TASK-005-25-09-2026.md`

Obras salvas, RECs e regras de IA (aprendizado do chat) apareciam em todos os projetos,
independentemente de onde foram salvos. Adicionada a coluna/campo `projeto`/`projeto_codigo` em
`obras`, `historico_rec` e `regras` (SQLite + Supabase), com filtro em todas as rotas de
leitura/escrita relevantes. Regras de IA, que nunca tinham sincronização com o Supabase (só
SQLite local de cada instância), ganharam esse suporte pela primeira vez, no mesmo padrão de
`regras_conversao` (TASK-003).

**Correção de convenção:** o valor de backfill/`DEFAULT` usado é o **código** do projeto (`"229"`
para RONDÔNIA), não o nome — corrigido depois de descoberto, durante o teste end-to-end da
TASK-006, que o nome não é a chave técnica usada em nenhum outro lugar do app para agrupar por
projeto (`regras_conversao.projeto_codigo`, `localStorage['projeto_selecionado_codigo']`).

**Validado:** simulado um banco "antigo" sem as colunas novas e confirmado o backfill via
`ALTER TABLE ... DEFAULT`; testado via HTTP real que salvar/listar obras, RECs e regras de IA em
projetos diferentes fica isolado, e que a listagem sem filtro preserva o comportamento anterior.

---

## 2026-09-19 — TASK-004-19-09-2026: correção de cache de JS no navegador (Revisão)

**Tipo:** correção de infraestrutura · **Tarefa:** `.ai/tasks/TASK-004-19-09-2026.md`

Alterações nos arquivos JS só carregavam em aba privativa (incognito). Causa raiz: o navegador
armazenava um "hard cache" dos arquivos `.html` (ex: `/static/resumo.html`). Por causa
do cache local, o navegador nunca pedia o arquivo ao servidor, e assim nunca recebia os
cabeçalhos de `NO_CACHE` criados. Para invalidar o hard cache local, era preciso alterar 
a URL de navegação (ex: de `/static/resumo.html` para `/resumo`).

**Alterado:**
- `app.py` — Novas rotas `@app.get("/resultado_orcamento")` e `@app.get("/orcamento")` servindo as páginas sem cache.
- `static/index.html`, `static/resumo.html`, `static/script.js`, `static/resumo.js` — Alteradas todas as navegações `window.location` e `window.open` que apontavam para `/static/*.html`, alterando para as rotas limpas do `app.py` (ex: `/resumo`, `/resultado_orcamento`).
- `middleware/nocache_middleware.py` — Injeta `NO_CACHE_HEADERS` em respostas `/static/*.html` (implementado no passo anterior, mantido por precaução).

**Validado:** aguardando confirmação do usuário em aba normal após reinício do servidor.

---


**Tipo:** melhoria de UX + segurança de processo · **Tarefa:** `.ai/tasks/TASK-002-19-09-2026.md`

Ao iniciar o programa, o navegador abria diretamente na página principal com dados em cache,
sem exigir novo login. Além disso, o processo Python na porta 8000 ficava ativo mesmo após
fechar todas as abas do navegador, sem forma simples de encerrá-lo pela UI.

**Alterado:**
- `app.py` — URL de abertura mudada de `/` para `/login?new_session=1`; novo endpoint
  `GET /api/shutdown` que envia `SIGTERM` ao processo após 500 ms (bloqueado com HTTP 403
  em `APP_MODE == "server"`).
- `middleware/auth_middleware.py` — `/api/shutdown` adicionado a `PUBLIC_ROUTES`.
- `static/login.html` — script IIFE que, ao detectar `?new_session=1`, limpa `auth_token`,
  `user_id`, `user_email` e `is_admin` do `localStorage` e pré-preenche o campo email com
  o último email usado (conveniência — senha nunca armazenada).
- `static/index.html` — botão "Encerrar Servidor" adicionado ao header; ao confirmar o
  diálogo, chama `GET /api/shutdown` e exibe mensagem "Servidor encerrado. Pode fechar esta aba."

**Validado:** servidor real reiniciado; navegador abriu em `/login?new_session=1` ✅;
login realizado com dados da nuvem carregados corretamente ✅; botão "Encerrar Servidor"
acionado com confirmação — log: `Encerramento solicitado pelo usuário via /api/shutdown.` ✅

---


**Tipo:** melhoria de UI · **Tarefa:** `.ai/tasks/TASK-001-19-09-2026.md`

A página inicial (`index.html`) exibia apenas dois caminhos de entrada: "Procurar Arquivo"
e "Montar Projeto". O usuário não tinha como navegar diretamente para a tela de resultado
de orçamento (`resultado_orcamento.html`) sem antes passar por outra tela.

**Alterado:**
- `static/index.html` — adicionados `<p>ou acesse os orçamentos</p>` e
  `<button id="btn-manipular-orcamentos">Manipular Orçamentos</button>` no `upload-card`,
  após o botão "Montar Projeto". O botão navega para `/static/resultado_orcamento.html`
  via `window.location.href`, no mesmo padrão já usado pelo botão "Montar Projeto".

**Não alterado:** nenhuma funcionalidade das abas, nenhuma rota de API, nenhum banco
de dados, nenhum contrato de `localStorage`.

**Validado:** servidor já em execução; alteração visível ao recarregar a página inicial.

---


**Tipo:** nova tabela + correção de código · **Tarefa:** `.ai/tasks/TASK-003-18-09-2026.md`

O botão "Salvar na Nuvem" das regras de conversão (CABOS/OUTROS → totalizadora, em
`resumo.html`/`resumo.js`) nunca escrevia no Supabase — `routers/regras.py` lia e gravava apenas
na tabela `configuracoes` do SQLite local. Cada instância do backend (desktop, Render) tem seu
próprio arquivo de banco, sem nenhum compartilhamento; por isso as regras "somem" ao trocar de
ambiente ou reiniciar o servidor.

**Alterado:**
- `scripts/schema_supabase.sql` — nova tabela `regras_conversao` (`projeto_codigo` PK,
  `regras_json`), bloco idempotente no mesmo padrão das migrações anteriores.
- `routers/regras.py:get_regras_conversao` — passa a consultar o Supabase primeiro (fonte de
  verdade quando configurado/alcançável), com fallback para o SQLite local, no mesmo padrão de
  `routers/obras.py:get_obras`.
- `routers/regras.py:save_regras_conversao` — passa a gravar local **e** fazer upsert no
  Supabase; diferente de `obras.py` (que engole erro de Supabase em silêncio — problema 21 do
  `STATE.md`, não corrigido aqui por estar fora do escopo pedido), aqui a falha de sincronização
  vira `HTTPException` 500 visível, no padrão já aprovado na TASK-002.
- `static/resumo.js:salvarRegrasNuvem` — passa a mostrar `data.detail` (mensagem real de erro) em
  vez de um texto genérico, no mesmo padrão já usado em `orcamento.html`.

**Validado:** servidor real subido a partir de uma cópia isolada do projeto (stubs para `ezdxf`,
`pymupdf` e `supabase`, técnica já usada nas tarefas anteriores), autenticado como admin:
- `POST` sem Supabase configurado → `200`, grava só local (comportamento preservado)
- `POST` com Supabase configurado porém falhando → `500` com mensagem clara, e SQLite local
  confirmado gravado mesmo assim
- `GET` de um `projeto_codigo` que só existia no Supabase "fake" (nunca gravado localmente) →
  retornou o dado da nuvem corretamente, confirmando que a integração é real e não um fallback
  disfarçado
- `POST` seguido de `GET` do mesmo `projeto_codigo`, com Supabase "fake" funcionando → o `GET`
  retornou exatamente o que foi salvo, vindo da nuvem

**Mesclado em `main`** via PR #5 (commit de merge `cac4b85`). Usuário rodou a migração no
Supabase real, redeployou e **confirmou**: regras salvas na nuvem persistem tanto no desktop
quanto na versão online. **TASK-003 encerrada.**

---

## 2026-09-18 — TASK-002 (continuação): `origem` era removida do payload antes do envio ao Supabase

**Tipo:** correção de código · **Tarefa:** `.ai/tasks/TASK-002-18-09-2026.md`

Depois que o usuário rodou a migração `ALTER TABLE ... ADD COLUMN origem` no Supabase real e
reimportou a base master completa, `origem` continuava `null` em todas as linhas, sem nenhum erro
sendo reportado. Achado um trecho em `routers/admin.py` (`upload_master_csv` e `sync_master_all`)
que filtrava `origem` para fora do dicionário antes de cada `INSERT`/`DELETE` no Supabase
(`rows_supabase = [{k: v for k, v in r.items() if k != 'origem'} for r in rows_to_insert]`), com o
comentário "A coluna 'origem' não existe no Supabase - removemos antes do insert". Esse trecho
veio de um commit anterior à TASK-002 (`08cd044`, 17/09), feito quando a coluna de fato ainda não
existia — e sobreviveu a um merge de `main` de volta para a branch de trabalho sem gerar conflito,
continuando a remover `origem` mesmo depois do schema já ter a coluna. Por isso o `INSERT` nunca
falhava (sem erro visível) e `origem` nunca chegava na nuvem, independentemente de qualquer outro
fix já aplicado.

**Alterado:** removida a filtragem `rows_supabase = [...]` nos dois handlers; ambos voltam a
enviar `origem` ao Supabase como qualquer outra coluna, no mesmo padrão já usado por
`add_master_row`.

**Mesclado em `main`** via PR #3 (commit de merge `36de3aa`). Usuário redeployou/reconstruiu o
ambiente, reimportou a base master e **confirmou**: `ORIGEM` aparece corretamente preenchida em
`GET /api/orcamento/dados` após recarregar a tela. **TASK-002 encerrada.**

---

## 2026-09-18 — TASK-002: coluna `origem` restabelecida na tabela master de orçamento

**Tipo:** correção de schema + bug de código + diagnóstico · **Tarefa:**
`.ai/tasks/TASK-002-18-09-2026.md`

Investigado por que a coluna `ORIGEM`, ao ser importada num CSV para a tabela master, não ficava
disponível depois de recarregar a tela. Dois problemas distintos, mesma raiz de família da
TASK-001: schema do Supabase desatualizado em relação ao código.

**Causa raiz:** `tabela_orcamento_master` no Supabase real nunca teve a coluna `origem`
(`scripts/schema_supabase.sql` não a declarava). `admin.py:upload_master_csv` lia e gravava
`origem` certo no CSV e no SQLite local, mas o `INSERT` espelhado no Supabase falhava com o erro
**engolido em silêncio** — a resposta ao admin continuava dizendo sucesso. Na sequência, o próximo
`GET /api/orcamento/dados` puxava a master antiga da nuvem e sobrescrevia o local, apagando o dado
que tinha acabado de ser importado.

**Bug de código independente:** `admin.py:sync_master_all` (botão "Sincronizar Tudo com a Nuvem")
nunca lia nem gravava `origem` em lugar nenhum, apesar do frontend já enviar o campo — não
dependia do schema do Supabase, era um bug puro no Python.

**Alterado:**
- `scripts/schema_supabase.sql` — coluna `origem` adicionada ao `CREATE TABLE
  tabela_orcamento_master` + migração idempotente para instalações existentes
- `routers/admin.py:sync_master_all` — passou a ler/gravar `origem`, alinhado com
  `upload_master_csv`/`add_master_row`
- `routers/admin.py:upload_master_csv` — falha de sincronização com o Supabase agora vira
  `HTTPException` 500 visível ao admin (`"Tabela local atualizada, mas falhou ao sincronizar com a
  nuvem (Supabase): {erro}"`), em vez de log silencioso; adicionado `except HTTPException: raise`
  para essa exceção não ser reembrulhada pelo handler genérico da função (que trocaria o status
  500 por 400 e duplicaria a mensagem) — **melhoria aprovada explicitamente pelo usuário**
- `static/orcamento.html` — o botão "Sobrescreve a tabela mestre" passou a mostrar a mensagem real
  de erro (`data.detail`), no mesmo padrão já usado pelo botão "Sincronizar Tudo com a Nuvem"

**Validado:** servidor real subido localmente; `POST /api/admin/sync-master-all` com `origem`
preenchida confirmado persistindo no SQLite; `POST /api/admin/upload-master` com Supabase
configurado mas inválido confirmado retornando `500` com mensagem clara (não mais `400`
reembrulhado) e, mesmo assim, salvando `origem` corretamente no SQLite local.

**Não corrigido nesta etapa** (depende de ação do usuário, fora do alcance deste ambiente): rodar
o `ALTER TABLE` no Supabase real.

**Achados registrados, não corrigidos** (fora do escopo pedido): o botão "Importar CSV Local"
chama um endpoint (`/api/upload/csv-orcamento`) que não existe no backend; a tabela pessoal do
usuário (`routers/orcamento.py`) também não trata `origem`, mas nunca sincroniza com a nuvem.

---

## 2026-09-18 — TASK-001: login/cadastro via Supabase restabelecidos (desktop + Render.com)

**Tipo:** correção de schema (dados) + diagnóstico (código) · **Tarefa:**
`.ai/tasks/TASK-001-18-09-2026.md` · **Status final: CONCLUÍDA**

Investigado com o usuário por que login e cadastro de usuário (Supabase) haviam parado de
funcionar. RLS descartado (estava desabilitado). Confirmado que a tabela real `usuarios_nuvem`
do usuário tinha `is_admin` mas não tinha `role` — coluna que `routers/admin.py` grava em
`POST /api/admin/users` (falha visível, com mensagem enganosa de "e-mail duplicado") e em
`PUT /api/admin/users/{id}/role` (falha silenciosa).

**Alterado no código:**
- `scripts/schema_supabase.sql` — migração idempotente adicionando `role` a `usuarios_nuvem`
- `routers/health.py` — `GET /api/health` ganhou `supabase_configured`, `supabase_reachable` e
  `usuarios_nuvem_has_rows`, sem expor dados de usuário (rota é pública)

**Validado:** servidor real subido localmente (dependências pesadas substituídas por stubs, já
que não influenciam este diagnóstico); `/api/health` testado com Supabase não configurado e com
credenciais configuradas porém inválidas — os dois casos respondem corretamente, sem quebrar a
rota.

**Feito pelo usuário, fora deste ambiente (sem acesso a Supabase/Render a partir daqui):**
1. `ALTER TABLE ... ADD COLUMN role` no Supabase real → destravou login e cadastro
2. Efeito colateral descoberto e corrigido: o `DEFAULT 'operador'` do passo 1 fez backfill em
   **todas** as linhas existentes, inclusive contas admin — derrubando o acesso ao painel admin
   até um `UPDATE` reconciliando `role` com `is_admin` (ver detalhe em STATE.md, problema 8, e no
   arquivo da tarefa)
3. Configuradas `SUPABASE_URL`/`SUPABASE_KEY` nas env vars do serviço Render.com (nunca haviam
   sido definidas lá — ambiente de deploy adicional, não documentado antes desta tarefa) e forçado
   um redeploy manual para pegar o código já mesclado em `main`

**Resultado confirmado pelo usuário:** login, painel admin e cadastro de usuário funcionando tanto
no desktop quanto na aplicação online (Render.com).

**Não alterado** (fora do escopo pedido): qualquer lógica de `orcamento_calc.py` ou das tabelas de
orçamento.

---

## 2026-09-18 — Estrutura de contexto IA-First criada

**Tipo:** documentação · **Alterações de código: nenhuma**

Criada a pasta `.ai/` como fonte única de contexto do projeto para desenvolvimento assistido por IA.

### Estado do projeto no momento do levantamento
- Versão `2.0.0`, branch `claude/beautiful-pasteur-2tdk18`, commit base `530d5cc`
- Backend FastAPI com 11 routers e 9 services; frontend HTML/CSS/JS vanilla
- SQLite local (11 tabelas) + Supabase (5 tabelas) em arquitetura offline-first
- Sem nenhum teste automatizado

### Arquivos analisados para o levantamento
**Raiz:** `app.py`, `config.py`, `database.py`, `models.py`, `websocket_manager.py`, `trigger.py`,
`requirements.txt`, `.env.example`, `.gitignore`, `build.bat`, `README.md`

**Routers:** `auth.py`, `upload.py`, `orcamento.py`, `regras.py`, `recs.py`, `obras.py`,
`projetos.py`, `ai_chat.py`, `admin.py`, `health.py`, `update.py`

**Services:** `pdf_service.py`, `dxf_service.py`, `orcamento_calc.py`, `supabase_client.py`,
`sync_service.py`, `offline_queue.py`, `realtime_sync.py`, `connectivity_monitor.py`,
`auto_updater.py`

**Middleware:** `auth_middleware.py`

**Frontend:** `script.js`, `resumo.js` (seções de lógica de negócio, cálculo de `qtdAtivos`, motor de
regras e camada unificada), `resultado_orcamento.html` (cálculo e REC), `resumo.html`,
`orcamento.html`, `index.html`, `login.js`, `admin.js`, `auth_fetch.js`

**Dados e infra:** `data/tabela_seed.csv`, `prompt_rede_eletrica.txt`,
`scripts/schema_supabase.sql`, `scripts/release.py`,
`.github/workflows/deploy-backend.yml`, `.github/workflows/deploy-frontend.yml`

### Estrutura criada
```
.ai/
├── CONTEXT.md        identificação, tecnologias, fluxo, entradas/saídas, 17 regras de negócio
├── ARCHITECTURE.md   camadas, pipeline, módulos, banco, API, exemplos entrada→regra→saída
├── GLOSSARY.md       termos do sistema e ativos de rede elétrica, com origem de cada definição
├── RULES.md          12 regras de trabalho para agentes + protocolo de 6 fases
├── STATE.md          implementado, não conectado, 24 problemas conhecidos, limitações
├── CHANGELOG.md      este arquivo
├── tasks/TEMPLATE.md modelo de tarefa
└── decisions/        TEMPLATE.md + ADR-001, ADR-002, ADR-003 (retroativos)
```

### Documentado a partir do código
- Pipeline completo: extração → classificação → separação → camada unificada → motor de regras →
  cálculo → resultado
- 17 regras de negócio (RN-01 a RN-17) com referência a arquivo e linha
- Cascata de 5 passos de resolução de ativo e expansão ativo → componente → códigos
- Precedência da base técnica (usuário > master > seed) e prioridade de projeto
- Contratos de `localStorage` entre as telas

### Registrado como pendente ou divergente (não corrigido)
- Três workers de background escritos e nunca iniciados (`offline_queue`, `realtime_sync`,
  `connectivity_monitor`)
- `Dockerfile` e `fly.toml` ausentes, embora exigidos pelo workflow de deploy
- Divergências de schema entre SQLite e Supabase (`origem`, `is_admin`)
- Três conjuntos inconsistentes de roles (`viewer` vs. `operador` vs. `editor`)
- Lógica de classificação duplicada entre `script.js` e `resumo.js`
- Ausência total de testes

### Pontos que necessitam confirmação do usuário
Registrados em GLOSSARY.md §3 e na seção final do relatório de inicialização: `K1`, `L1`, `L2`,
`TR+1/+2/+3`, `U23`, `U32`, significado de `REC`, `RS M`/`RS T`, `FLY`, `CV`, e o propósito da
coluna `filtro`.

---

## Antes de 2026-09-18

Histórico anterior não registrado neste formato. As mensagens de commit existentes
(`"comit changes"`, `"melhorias"`, `"bug fixes"`, `"mudanças"`) não permitem reconstruir com
segurança quais alterações de arquitetura ou de regra de negócio ocorreram em cada ponto.

`[NÃO DETERMINADO PELO CÓDIGO]` — para reconstruir o histórico anterior seria necessário o
relato do usuário.

Um marco é identificável pelo próprio código: o salto para a versão `2.0.0`, descrito em
`app.py:2` como *"Ponto de entrada da aplicação FastAPI (v2.0 - Profissionalizado)"*, introduziu a
separação em `routers/`, `services/` e `middleware/`, os dois modos de operação e a autenticação JWT.
