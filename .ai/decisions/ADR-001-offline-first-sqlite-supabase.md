# ADR-001 — Arquitetura offline-first: SQLite local com Supabase como nuvem

> ⚠️ **ADR retroativo.** Esta decisão já estava implementada quando a documentação `.ai` foi criada
> (2026-09-18). O contexto e as consequências abaixo foram **reconstruídos a partir do código**.
> Não existe registro histórico da deliberação original.

## Contexto

O sistema é usado por projetistas de redes de distribuição que trabalham em campo e em escritório.
O código trata explicitamente o cenário sem internet: `services/offline_queue.py` fala em
"operações CRUD feitas quando o sistema está sem internet", `routers/auth.py` documenta
"Login com fallback offline", e o app é distribuído também como executável desktop (`build.bat`,
`pywebview`).

Ao mesmo tempo, existe necessidade de dados compartilhados entre usuários: a base técnica oficial
(tabela master), os projetos/concessionárias e o histórico de RECs.

## Problema

Como manter o sistema plenamente utilizável sem conexão, sem abrir mão de uma base técnica
centralizada e de dados compartilhados entre os usuários.

## Alternativas consideradas

`[NÃO DETERMINADO PELO CÓDIGO]` — não há registro de quais alternativas foram efetivamente
avaliadas. As opções abaixo são o espaço de decisão plausível, apresentadas para contexto, **não**
como histórico:

### Alternativa A — Somente nuvem (Supabase como único banco)
- Prós: uma fonte de verdade, sem sincronização
- Contras: inutilizável sem internet — incompatível com o uso em campo

### Alternativa B — Somente local (SQLite por instalação)
- Prós: simples, sempre disponível
- Contras: sem base técnica compartilhada, sem histórico entre usuários

### Alternativa C — Local como fonte primária + nuvem como espelho (a adotada)

## Decisão

**SQLite local é a fonte de verdade operacional; o Supabase é a camada de compartilhamento e
respaldo.** Toda leitura funciona a partir do banco local; a nuvem é consultada quando disponível e
qualquer falha dela é degradada silenciosamente para o local.

A precedência de leitura da base técnica é explícita e hierárquica
(`services/sync_service.py:get_merged_orcamento`):

```
tabela_orcamento WHERE user_id = <atual>   ← edições do próprio usuário
      ↓ vazio
tabela_orcamento_master                    ← master sincronizada do Supabase
      ↓ vazio
tabela_orcamento WHERE user_id IS NULL     ← seed do CSV
```

Para autenticação, a ordem é invertida — tenta a nuvem primeiro e cai para o local
(`routers/auth.py:login`) — porque as credenciais são geridas centralmente pelo admin.

## Consequências

### Positivas
- O app funciona integralmente sem internet, inclusive login
- A base técnica pode ser atualizada centralmente pelo admin e propagada aos clientes
- Nenhuma operação do usuário é bloqueada por indisponibilidade da nuvem
- O seed embutido garante que uma instalação nova já nasce utilizável

### Negativas / custos aceitos
- **Duplicação de schema:** toda tabela compartilhada existe duas vezes (SQLite e Postgres), e as
  duas definições podem divergir — e divergiram (ver STATE.md, problemas 5 a 7: colunas `origem` e
  `is_admin`)
- **Falhas de nuvem são invisíveis:** o padrão `try: … except Exception: pass` aparece em
  `obras.py`, `recs.py`, `projetos.py` e `admin.py`. O usuário não sabe se salvou na nuvem
- **Sem resolução de conflito:** não há merge nem detecção de escrita concorrente. `sync_tabela_master`
  apaga a master local inteira e reinsere
- **Sincronização de volta (local → nuvem) não está ativa:** `offline_queue.py` foi escrito para
  isso, mas `enqueue_operation()` e `start_offline_queue_worker()` nunca são chamados

### Implicações para implementações futuras
1. **Toda nova entidade compartilhada precisa existir nos dois bancos** — `database.py:init_db` e
   `scripts/schema_supabase.sql` — e os dois devem ser atualizados na mesma alteração.
2. **Nunca torne uma leitura dependente da nuvem.** O padrão é: tenta a nuvem, cai para o local.
3. **Respeite a precedência da base técnica.** Alterar a ordem de `get_merged_orcamento` muda qual
   base o orçamento usa — é mudança de regra de negócio (RN-11), não otimização.
4. Ao conectar a fila offline, decida antes a política de conflito; hoje não existe nenhuma.

## Evidência no código

- `services/sync_service.py:76-104` — precedência da base técnica
- `services/sync_service.py:9-73` — download paginado da master
- `routers/auth.py:38-60` — login nuvem → local
- `database.py:15-27` — conexões SQLite com WAL
- `services/supabase_client.py` — singleton que devolve `None` quando não configurado
- `routers/obras.py`, `recs.py`, `projetos.py` — padrão de fallback silencioso
- `services/offline_queue.py` — fila escrita, não conectada

## Data

2026-09-18 (documentação retroativa; data da decisão original desconhecida)

## Status

**ACEITA** — em vigor, com a parte de sincronização de volta ainda não conectada.
