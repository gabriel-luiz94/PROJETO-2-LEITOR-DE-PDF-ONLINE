# ADR-003 — Obra no DWG (dicionário de objetos nomeados) + arquivo ao lado; marcas transitórias

## Contexto

Briefing §3 itens 9–10 e §5: persistir "dentro do DWG e em `<nome>.obra.json`", "sem camada de marcas", vínculos como
gráficos transitórios. Ponto em aberto 5: formato dentro do DWG.

## Decisão

- DWG: `NamedObjectsDictionary` → `DBDictionary "ORCAMENTO_LEITOR"` → `Xrecord "OBRA"` com o MESMO JSON do arquivo,
  em pedaços de 250 caracteres (`DxfCode.Text`). Um só formato para os dois lugares; reabrir = continuar.
- Arquivo: `<desenho>.obra.json` na pasta do desenho (UTF-8 sem BOM). Desenho ainda sem arquivo: só no DWG, com aviso.
- Ao abrir: o DWG tem prioridade; senão o arquivo.
- Números e vínculos na tela: `TransientManager` (não gravam). Fallback `SEM_TRANSIENTES` (vetores do Editor) para nanoCAD sem essa API.
- O desenho só muda em dois casos: gravar a obra e trocar a layer da anotação (opção do pop-up).

## Alternativas consideradas

- XData nas entidades: espalha a obra pelo desenho e se perde ao copiar/apagar entidades — descartada.
- Extension dictionary por entidade: idem, e "Reler do zero" teria de limpar entidade por entidade.
- Camada de marcas: descartada pelo briefing.

## Evidência no código

`src/Plugin/PersistenciaDwg.cs`, `src/Plugin/Marcas.cs`, `Comandos.TrocarLayerDoTexto`.

## Data

2026-10-31

## Status

**Atual:** ACEITA
