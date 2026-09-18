# REGRAS DE TRABALHO PARA AGENTES DE IA

> Este documento define como qualquer agente de IA deve trabalhar neste projeto.
> Ele vale para todas as tarefas futuras, até ser explicitamente alterado pelo usuário.

---

## Regra 1 — Entender antes de alterar

Antes de modificar qualquer código, consulte nesta ordem:

1. `.ai/CONTEXT.md` — o que o sistema é, entradas, saídas, regras de negócio
2. `.ai/RULES.md` — este documento
3. `.ai/STATE.md` — o que está pronto, o que está quebrado, o que é limitação conhecida
4. `.ai/ARCHITECTURE.md` — módulos, fluxo, banco, API, dependências
5. `.ai/GLOSSARY.md` — termos de domínio

Depois disso, leia os arquivos específicos da tarefa. **Não leia o projeto inteiro** — a
documentação existe justamente para evitar isso.

**A documentação nunca substitui o código.** Quando houver conflito entre o que está documentado e
o que o código faz, **o código vence**. Aponte o conflito ao usuário e registre a correção no
arquivo `.ai` correspondente — nunca ajuste silenciosamente o código para casar com a documentação.

---

## Regra 2 — Não inventar comportamento

Quando uma regra de negócio for ambígua:

- **não assuma** o comportamento pretendido;
- **não invente** uma definição plausível;
- identifique explicitamente a ambiguidade;
- apresente a interpretação necessária e **pergunte** antes de implementar.

Isto vale especialmente para:

- termos marcados como `[DEFINIÇÃO NECESSITA CONFIRMAÇÃO DO USUÁRIO]` no GLOSSARY.md;
- as regras técnicas que existem apenas em `prompt_rede_eletrica.txt` (elas **não são validadas por
  código** — não presuma que o sistema as aplica, nem as implemente sem pedido explícito);
- códigos de ativo cujo significado não está no código nem no seed.

---

## Regra 3 — Alteração mínima

Modifique somente o necessário para atender à tarefa.

- Nenhuma refatoração não solicitada.
- Nenhuma "limpeza" oportunista de código vizinho.
- Nenhuma abstração criada para necessidades hipotéticas.
- Nenhuma reorganização de arquivos junto com uma correção de bug.

Se, ao resolver a tarefa, você identificar outro problema, **registre-o** (em `.ai/STATE.md` ou como
uma tarefa nova em `.ai/tasks/`) em vez de corrigi-lo no mesmo trabalho.

---

## Regra 4 — Não duplicar lógica

Antes de criar uma função, classe ou módulo:

1. procure uma implementação existente;
2. avalie a reutilização;
3. evite a duplicação.

⚠️ **Atenção específica deste projeto:** `static/script.js` e `static/resumo.js` já contêm cópias
independentes de `isGray()`, `processAtivoFormula()` e da cascata de classificação de entidade
(`updateRowLogic` vs. `computeRowLogic` vs. `autoClassifyEntidade`).

**Ao alterar qualquer regra de classificação, verifique todas as cópias.** Alterar só uma faz a aba
principal e a aba de resumo discordarem entre si — um bug silencioso e difícil de rastrear.
Unificar essas cópias é uma refatoração legítima, mas só mediante pedido explícito (Regra 3).

---

## Regra 5 — Preservar compatibilidade

Não altere interfaces, estruturas, nomes ou comportamentos existentes sem necessidade explícita.

Contratos que **quebram outras partes do sistema** se alterados:

| Contrato | Quem escreve | Quem lê |
|---|---|---|
| `localStorage['processar_dados']` | `script.js` | `resumo.js` |
| `localStorage['orcamentoPayload']` | `resumo.js` | `resultado_orcamento.html` |
| Payload de `POST /api/orcamento/calcular` | `resultado_orcamento.html` | `orcamento_calc.processar_calculo` |
| Formato de ativo `[ATIVO] [FASE] [COMPRIMENTO]` (CABOS) | frontend e IA | `orcamento_calc.py:46-96` |
| Formato de ativo `<qtd>-<ativo>` (OUTROS) | frontend e IA | `orcamento_calc.py:99-122` |
| Colunas do CSV da base técnica | usuário/admin | `orcamento.py`, `admin.py`, `database.py` |
| Colunas do CSV de regras de conversão | usuário | `resumo.js:importarRegrasCSV` |
| Schema do Supabase | `scripts/schema_supabase.sql` | `sync_service`, `auth`, `admin`, `obras`, `recs` |

Mudanças de schema no SQLite devem seguir o padrão já existente em `database.py:init_db`:
`ALTER TABLE … ADD COLUMN` dentro de `try/except sqlite3.OperationalError`, mantendo a
idempotência. **Não** apague nem recrie tabelas com dados de usuário.

---

## Regra 6 — Analisar antes de implementar

Para qualquer tarefa que altere código:

1. compreenda o problema;
2. identifique os arquivos afetados;
3. identifique as funções/componentes afetados;
4. identifique os riscos e efeitos colaterais;
5. **apresente o plano**;
6. aguarde aprovação quando o processo exigir;
7. implemente;
8. teste.

Formato do plano:

```
OBJETIVO          o que muda, em uma frase
ARQUIVOS AFETADOS caminho:linha de cada ponto de alteração
ALTERAÇÕES        o que será feito em cada ponto
RISCOS            o que pode quebrar (use a tabela da Regra 5)
TESTES            como a mudança será verificada
```

---

## Regra 7 — Testes

⚠️ **O projeto não possui testes automatizados hoje** (ver STATE.md). Não existe framework de teste
configurado, nem CI que rode testes.

Enquanto isso não mudar:

- Toda alteração relevante deve vir acompanhada de uma **verificação explícita** — o que foi
  executado e qual foi o resultado.
- Se você criar testes, o alvo natural e de maior retorno é `services/orcamento_calc.py`: é uma
  função pura, sem I/O e sem dependências do projeto, e concentra as regras RN-03 a RN-10.
- Nunca declare que algo "funciona" sem ter executado. Se não foi possível testar
  (por exemplo, fluxo de UI sem navegador disponível), **diga isso explicitamente** em vez de
  afirmar sucesso.
- Verificações mínimas antes de entregar: o app sobe (`python app.py`), a rota alterada responde, e
  o fluxo end-to-end afetado continua produzindo o mesmo resultado para a mesma entrada.

Depois de implementar: execute o que for verificável, informe os resultados e corrija as falhas
causadas pela sua alteração.

---

## Regra 8 — Preservar regras de negócio

As regras de negócio existentes têm **prioridade sobre preferências genéricas de programação**.

Não substitua uma regra específica por uma implementação "mais elegante" sem autorização. Boa parte
do que parece estranho neste código é intencional e reflete a realidade do domínio. Exemplos reais:

- a cascata de 5 passos de resolução de ativo (RN-06) parece redundante, mas cada passo cobre um
  formato diferente de preenchimento da base técnica;
- `hasApoioConflict` parece defensivo demais, mas evita classificar errado textos ambíguos;
- o recálculo de `qtdAtivos` **de baixo para cima** existe porque linhas standalone herdam a fase da
  linha seguinte;
- a regex `_RE_PARA_FMT` em `dxf_service.py` é deliberadamente sensível a maiúsculas: `\p` é
  formatação de parágrafo e `\P` é quebra de parágrafo. O comentário no código explica; **não o remova**.

Na dúvida sobre o porquê de uma regra: pergunte, não "corrija".

---

## Regra 9 — Documentação

Quando uma alteração modificar:

- arquitetura → atualize `ARCHITECTURE.md`
- regra de negócio → atualize `CONTEXT.md` (seção 7) e, se houver termo novo, `GLOSSARY.md`
- fluxo de dados → `CONTEXT.md` (seção 3) e `ARCHITECTURE.md` (seção 2)
- estrutura de banco → `ARCHITECTURE.md` (seção 4)
- comportamento importante → `CONTEXT.md`
- decisão arquitetural relevante → novo ADR em `.ai/decisions/`

Não modifique a documentação sem motivo. Documentação atualizada "por garantia" gera ruído e
desatualização cruzada.

---

## Regra 10 — Estado do projeto

Após concluir uma alteração, atualize `.ai/STATE.md`:

- mova itens entre "Implementado" / "Em desenvolvimento";
- remova problemas resolvidos;
- registre problemas descobertos;
- atualize "Última atualização" com data e motivo.

Alterações relevantes também entram em `.ai/CHANGELOG.md`.

---

## Regra 11 — Segurança e dados sensíveis

- **Credenciais só por variável de ambiente.** Nunca escreva `SUPABASE_KEY`, `JWT_SECRET`,
  `ADMIN_PASSWORD` ou chaves de IA em código, commit, log ou documentação. O padrão correto já está
  em `services/supabase_client.py`.
- **Não versione dados de usuário.** `.gitignore` já cobre `*.pdf`, `*.dxf`, `*.db`, `.env`.
  Arquivos desse tipo na raiz do repositório são erro, não conteúdo.
- **Não afrouxe a trava do `/extract-local`.** A checagem `IS_FROZEN` impede leitura arbitrária de
  disco quando o app está hospedado (RN-15).
- Ao mexer em autenticação, lembre que o middleware protege `/api/*` por padrão: adicionar uma rota
  a `PUBLIC_ROUTES` ou `PUBLIC_PREFIXES` é uma **decisão de segurança**, não um detalhe — justifique.
- Não registre `senha_hash`, tokens ou conteúdo de `dados_json` em log.

---

## Regra 12 — Idioma e estilo

- Código, comentários, mensagens de commit e documentação: **português**, seguindo o padrão já
  existente no projeto.
- Docstrings de módulo no formato já usado: `"""caminho/arquivo.py — descrição curta."""`
- Não adicione comentários que apenas repitam o que o código faz. Comente o **porquê**, quando ele
  não for óbvio — como já ocorre em `dxf_service.py:10-15`.
- Não introduza dependências novas sem necessidade clara e sem avisar o usuário.

---

## Protocolo de desenvolvimento (fases)

```
FASE 1  CONTEXTO       ler os arquivos .ai + localizar apenas os arquivos da tarefa
FASE 2  ENTENDIMENTO   comportamento atual → desejado; arquivos, funções,
                       dependências, efeitos colaterais
FASE 3  PLANO          OBJETIVO / ARQUIVOS AFETADOS / ALTERAÇÕES / RISCOS / TESTES
FASE 4  IMPLEMENTAÇÃO  somente o solicitado
FASE 5  VALIDAÇÃO      funcionalidade nova + funcionalidades afetadas + regressões
FASE 6  DOCUMENTAÇÃO   STATE.md, CHANGELOG.md e o que a Regra 9 exigir
```

---

## Economia de contexto

Esta estrutura existe para que o usuário **não precise reexplicar o projeto** a cada conversa.

- Não peça ao usuário informação que já está documentada aqui.
- Quando a tarefa puder ser entendida pela documentação + código, use essas fontes.
- Uma tarefa futura precisa conter apenas: **objetivo, problema, comportamento desejado, restrições,
  critérios de aceite e exemplos específicos** (ver `.ai/tasks/TEMPLATE.md`).

### Fonte única de verdade

Os arquivos `.ai` são **a** fonte de contexto do projeto. Não crie arquivos paralelos com
informação duplicada. Nunca crie `context_final.md`, `context_novo.md`, `regras_v2.md` e afins —
atualize o arquivo existente.

---

## Segurança contra alterações acidentais

Antes de qualquer alteração estrutural:

1. identifique o impacto;
2. informe os arquivos afetados;
3. verifique as dependências (ARCHITECTURE.md §10);
4. preserve o comportamento existente;
5. teste.

**Nunca apague código ou arquivos por parecerem não utilizados sem verificar as referências.**

Neste projeto isso é especialmente importante: existem módulos que parecem mortos mas são
funcionalidade planejada e já escrita (`offline_queue.py`, `realtime_sync.py`,
`connectivity_monitor.py` — definidos e nunca chamados). **Eles não são lixo; são funcionalidade
não conectada.** Ver STATE.md antes de tocar neles.
