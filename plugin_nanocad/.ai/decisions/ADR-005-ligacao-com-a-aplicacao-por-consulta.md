# ADR-005 — Ligação com a aplicação: API própria do plugin + consulta periódica de versão

## Contexto

O usuário quer deixar a aplicação aberta (localhost ou online), mudar Base de Orçamento / regras de vínculo etc. e ver
no plugin "na mesma hora" — inclusive estruturas não cadastradas na base.

## Alternativas consideradas

- **Reusar as rotas existentes** (`/api/regras-leitor/*`, `/api/vinculacao-config`, `/api/orcamento/dados`): 4+ chamadas, a base
  inteira (milhares de linhas) a cada consulta, e o plugin teria de repetir a resolução de ativos (RN-06/RN-08, parâmetros,
  conversão) para saber o que "não está na base" — mais uma cópia de lógica.
- **WebSocket `/ws`**: avisaria na hora, mas o `.NET 4` do nanoCAD 5 não tem cliente WebSocket pronto, passa mal por proxy no
  modo online e a aplicação não emite eventos para mudanças de regras/base.
- **Escolhida: rotas `/api/plugin/*` + consulta de versão a cada N s.** `versao` devolve só hashes (leve); `conferir` roda na
  aplicação o MESMO caminho do modo autônomo (Totalizadora → cálculo), então "não está na base" é a resposta da própria aplicação.

## Decisão

Rotas `GET /api/plugin/versao`, `GET /api/plugin/regras`, `POST /api/plugin/conferir` (só leitura, com login). O plugin consulta a
versão a cada 5 s (ajustável, mín. 2 s); muda → baixa/reconfere em segundo plano; aplica no thread do nanoCAD.

## Consequências

- "Ao vivo" = até um intervalo de atraso. No modo online, cada consulta lê regras/base (Supabase): intervalo maior se pesar.
- O desenho em si não é relido sozinho (precisa do contexto de comando); a obra é refeita do último desenho lido.
- Tokens no perfil do usuário, cifrados com DPAPI; senha nunca guardada.

## Evidência no código

`routers/plugin_cad.py`; `src/Nucleo/ClienteAplicacao.cs`; `src/Plugin/Sincronizador.cs`; `tests/test_plugin_cad_api.py`;
`tests/Nucleo.Testes/ClienteAplicacaoTests.cs`.

## Data

2026-10-31

## Status

**Atual:** ACEITA
