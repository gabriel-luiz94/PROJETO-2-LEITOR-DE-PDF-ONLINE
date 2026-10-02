# Incidente: o `pytest` rodou contra o Supabase REAL (2026-10-02)

## O que aconteceu
Ao rodar `python -m pytest tests -q` no computador com `.env` (Supabase configurado), os testes usaram o **Supabase de verdade**: leram **e gravaram** nele. Os testes sempre foram escritos supondo "sem nuvem" (como no Linux de desenvolvimento) e **não havia nada que impedisse** o uso da nuvem. Os 18 testes que falharam mostram isso: encontraram versões de regras/ajustes/prompts do seu usuário (`gabriel.sales@…`) e dados gravados por outros testes.

**Já corrigido na `main` (veja o PR):** `tests/conftest.py` zera as credenciais antes de qualquer teste, usa um banco temporário e **recusa rodar** se achar um cliente Supabase. Enquanto esse PR não estiver na sua cópia, **não rode o `pytest` nesse computador.**

## O que pode ter sido alterado (somente tabelas que os testes gravam)
Os testes usam os projetos **`DEFAULT`**, **`229`** e **`027`** e usuários de mentira (`admin@x.com`, `operador@x.com`, `u1`, `u2`).

| Tabela na nuvem | O que os testes gravaram | Risco |
|---|---|---|
| `regras_dominio` e `regras_dominio_historico` | regras de validação de teste em DEFAULT / 229 / 027 | **a versão atual desses projetos pode ser a de teste** |
| `ajustes_planilhas` e `ajustes_planilhas_historico` | ajustes de teste | idem |
| `prompts_validacao` e `prompts_validacao_historico` | prompts de teste (validar/corrigir/ajustar) | idem |
| `obras` | obras de teste com `user_id` `u1`/`u2` e ids `obra_1`, `obra_2`…, `auto_<12 letras/números>` | lixo; não aparece para os seus usuários |

Não há gravação de teste em: usuários, projetos, tabela de orçamento, regras do leitor, regras de conversão (os testes não chamam essas rotas). **Cada gravação preservou a versão anterior no histórico** (`*_historico`), então os dados reais são recuperáveis.

## Passo 1 — Só olhar (nada é alterado)
No **SQL Editor do Supabase**, rode (somente `select`):

```sql
-- Versões gravadas pelos testes (o autor termina em @x.com)
select 'regras'  as tabela, id, projeto_codigo, criado_por, criado_em from regras_dominio_historico   where criado_por like '%@x.com'
union all
select 'ajustes' as tabela, id, projeto_codigo, criado_por, criado_em from ajustes_planilhas_historico where criado_por like '%@x.com'
union all
select 'prompts' as tabela, id, projeto_codigo || '/' || prompt_id, criado_por, criado_em from prompts_validacao_historico where criado_por like '%@x.com'
order by criado_em;

-- Obras de teste
select id, nome, user_id, projeto, data from obras where user_id in ('u1', 'u2', 'u3');
```

Anote o **horário** da primeira linha: é quando os testes começaram.

## Passo 2 — Fazer um backup antes de mexer
No Supabase: **Table Editor → cada tabela acima → Export CSV** (as 6 tabelas de regras/ajustes/prompts + `obras`).

## Passo 3 — Recuperar as regras, ajustes e prompts (pela tela, sem SQL)
Como o histórico guarda **a versão que existia antes de cada gravação**, a versão real de cada projeto é a que ficou salva na **primeira** gravação de teste.

Para **cada** aba (Regras, Ajustes, Prompts da IA) e para **cada** projeto (`DEFAULT`, `229`, `027`):

1. **Resumo → Regras de validação** (como administrador) → escolha o projeto em **Editando**.
2. **Projeto 229 e 027:** se, **antes dos testes, esse projeto não tinha versão própria** (herdava o padrão), use **"Descartar ajustes do projeto"** (volta ao padrão e guarda a versão atual no histórico). Se tinha versão própria, use o passo 3.
3. Abra **Histórico** e escolha a versão **mais antiga entre as gravadas por `…@x.com`** (a primeira após o horário do passo 1; as mais novas são de teste) → **Reverter**.
4. Confirme que as regras/ajustes/prompts voltaram ao que você conhece (nomes, ligado/desligado, mensagens).

> Dúvida em algum projeto? **Não reverta às cegas**: me envie o resultado do Passo 1 (print) e eu indico, linha a linha, qual versão restaurar.

## Passo 4 — Limpar as obras de teste (depois de conferir o `select`)
```sql
delete from obras where user_id in ('u1', 'u2', 'u3');
```
Confira antes que o `select` do passo 1 só mostrou obras de teste (nomes como "Obra 2", "Rede Norte", "[Autônomo] …").

## O que NÃO foi afetado
- O seu **`banco_resumo.db` local**: os testes de rotas criam um banco temporário próprio. Mesmo assim, se fez o backup do roteiro (0.1), guarde-o até confirmar que tudo está normal.
- Chaves de IA: nenhum teste usa a chave real (as chamadas de IA são simuladas); nada foi enviado ao Gemini.
- Usuários, projetos, orçamento base, regras do leitor e regras de conversão.

## Prevenção
`tests/conftest.py` (rodando antes de qualquer teste): zera `SUPABASE_*` e as chaves de IA, aponta o banco para uma pasta temporária e executa `pytest.exit` se ainda existir cliente Supabase; `tests/test_isolamento.py` confere isso. Testes que precisem de nuvem usam um cliente **falso** (monkeypatch).
