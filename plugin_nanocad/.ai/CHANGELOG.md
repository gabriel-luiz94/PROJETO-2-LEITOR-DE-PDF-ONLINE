# CHANGELOG DO PLUGIN

> Alterações relevantes do `plugin_nanocad/` (arquitetura, regras, contrato, comportamento). Mais recente primeiro.

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
