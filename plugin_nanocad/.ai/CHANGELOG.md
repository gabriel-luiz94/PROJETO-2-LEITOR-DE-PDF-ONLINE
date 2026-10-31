# CHANGELOG DO PLUGIN

> Alterações relevantes do `plugin_nanocad/` (arquitetura, regras, contrato, comportamento). Mais recente primeiro.

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
