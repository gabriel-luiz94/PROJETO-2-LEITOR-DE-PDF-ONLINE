# CONTEXTO DO PROJETO

> Memória permanente de alto nível. Descreve o sistema **como ele existe hoje no código**,
> não como deveria ser. Informações não determináveis pelo código estão marcadas como
> `[NÃO DETERMINADO PELO CÓDIGO]`.
>
> Estado atual de implementação: ver [STATE.md](STATE.md).
> Detalhamento técnico de módulos e fluxo: ver [ARCHITECTURE.md](ARCHITECTURE.md).
> Termos de domínio: ver [GLOSSARY.md](GLOSSARY.md).

---

## 1. Identificação do projeto

| Item | Valor |
|---|---|
| **Nome (repositório)** | `PROJETO-2-LEITOR-DE-PDF-ONLINE` |
| **Nome (aplicação)** | "Leitor de Projetos Online Pro" (`app.py`), título da UI: "Leitor de Projetos" |
| **Versão** | `2.0.0` (`config.py:APP_VERSION`) |
| **Objetivo** | Ler projetos de rede elétrica de distribuição em PDF/DXF, extrair os ativos desenhados, classificá-los, aplicar regras de conversão e produzir um **orçamento** (lista de códigos com quantidades) para a concessionária |
| **Finalidade prática** | Substituir o trabalho manual de ler a prancha do projeto e montar a planilha de orçamento/REC item a item |
| **Usuário** | Projetistas/orçamentistas de redes de distribuição de energia. Dois papéis no sistema: `admin` e `operador` |
| **Problema que resolve** | Transcrição manual de ativos de projeto (postes, cabos, estruturas, trafos, chaves) para códigos de orçamento da concessionária — trabalho repetitivo e sujeito a erro |

O nome "Leitor de PDF" descreve apenas a porta de entrada. O sistema real é um
**orçamentista de redes de distribuição**: a leitura do PDF/DXF é a primeira das cinco etapas.

### Concessionárias / projetos atendidos

Cadastrados como *projetos* com código (`database.py:init_db`, `scripts/schema_supabase.sql`):

- `PARAIBA` → código `027`
- `RONDONIA` → código `229`

Novos projetos podem ser cadastrados por um admin (`POST /api/projetos`).
O seed `data/tabela_seed.csv` (2.328 linhas) contém a base técnica de `RONDONIA`.

---

## 2. Tecnologias

### Backend
| Tecnologia | Uso |
|---|---|
| **Python 3** | Linguagem |
| **FastAPI** | Framework HTTP + WebSocket |
| **Uvicorn** | Servidor ASGI |
| **PyMuPDF (`pymupdf`)** | Extração de texto, fonte, cor e flags de PDF |
| **ezdxf** | Extração de TEXT/MTEXT/ATTRIB/INSERT de arquivos DXF |
| **SQLite** | Banco local (`banco_resumo.db`), fonte de verdade offline |
| **Supabase** (`supabase-py`) | Banco Postgres na nuvem + sincronização |
| **PyJWT + bcrypt** | Autenticação (JWT HS256, 24 h) e hash de senha |
| **pywebview** | Janela de aplicativo desktop nativo |
| **python-dotenv** | Carga de `.env` |

### IA
| Tecnologia | Uso |
|---|---|
| **google-genai** | Provider `gemini` (padrão) |
| **openai** | Provider `openai` e compatíveis (Ollama, OpenRouter) via `base_url` customizável |

### Frontend
HTML/CSS/JavaScript **vanilla** (sem framework, sem build step), servido como estático
pelo próprio FastAPI. Estado de trabalho trafega entre páginas via `localStorage`.

### Empacotamento e deploy
- **PyInstaller** (`build.bat`) → `.exe` desktop Windows (`--onefile --windowed`)
- **GitHub Actions** → backend para **Fly.io**, frontend para **Cloudflare Pages**
  (⚠️ ver STATE.md: `Dockerfile` e `fly.toml` referenciados pelo workflow não existem no repositório)

---

## 3. Fluxo geral

```
ENTRADA                    PDF de projeto  |  DXF de projeto  |  PDFs Orçamento+Lista (REC)
    ↓
EXTRAÇÃO                   pdf_service / dxf_service
                           → linhas {pagina, texto, cor, fonte, tamanho, flags, layer}
    ↓
CLASSIFICAÇÃO (CONTEXTO)   script.js + resumo.js: computeRowLogic()
                           cor + layer + texto → {entidade, operação, ativo}
    ↓
SEPARAÇÃO EM 2 TABELAS     CABOS (entidade = CABO)  |  OUTROS (demais entidades)
    ↓
CAMADA UNIFICADA           syncTotalizadora(): rawItems com {baseId, origem, ativo, qtd, operacao}
    ↓
MOTOR DE REGRAS            Tabela de Regras de Conversão (ADIÇÃO / SUBSTITUIÇÃO, fator,
                           arredondamento, val_min, val_max) → Tabela Totalizadora (editável)
    ↓
CÁLCULO                    POST /api/orcamento/calcular → services/orcamento_calc.processar_calculo()
                           ativo → componente → todos os códigos → soma_i / soma_r por código
    ↓
RESULTADO                  Lista {operacao, mdo, codigo, desc_codigo, filtro, total}
                           + lista de ativos não encontrados na base técnica
    ↓
ARMAZENAMENTO              historico_rec (SQLite + Supabase), exportação .rec / CSV
```

Diagramas por componente e detalhamento de cada etapa: ver [ARCHITECTURE.md](ARCHITECTURE.md).

---

## 4. Principais conceitos

> Definições completas de todos os termos: ver [GLOSSARY.md](GLOSSARY.md).

- **ATIVO** — item físico da rede identificado no projeto (`DT11/300`, `CAA 2 ABC 35 m`, `1-CFU`).
  É a chave de entrada do cálculo.
- **ENTIDADE** — classificação do ativo, usada para separar as tabelas e como observação.
  Valores fixos: `0`, `CABO`, `CHAVE`, `TRAFO`, `ESTRUTURA`, `APOIO`, `IP`, `POSTE`, `RAMAIS`, `CERCA`.
- **OPERAÇÃO** — `I` (instalar), `R` (remover), `M` (manter); prefixo `*` = linha viva/energizado.
- **ORIGEM** — `CABOS` ou `OUTROS`. Define o formato de escrita do ativo, quais linhas da base
  técnica podem ser usadas e a quais itens uma regra de conversão se aplica.
- **COMPONENTE** — agrupador na base técnica. Um ativo aponta para um componente, e o componente
  expande para **todos** os códigos que ele consome (mão de obra + materiais).
- **CÓDIGO** — código de serviço/material da concessionária (saída final do orçamento).
- **FATOR I / FATOR R** — multiplicadores por código para instalação e remoção, respectivamente.
- **MDO** — `MAO-DE-OBRA` ou `MATERIAL`; classifica a linha do orçamento e define a ordenação.
- **REC** — registro da obra orçada, salvo por `numero_obra` em `historico_rec`.

---

## 5. Entradas

| # | Entrada | Onde entra | Processamento |
|---|---|---|---|
| 1 | **PDF de projeto** | `POST /upload` | `pdf_service.extract_pdf_content` — texto por span, com fonte, tamanho, **cor** e flags |
| 2 | **DXF de projeto** | `POST /upload` | `dxf_service.extract_dxf_content` — TEXT/MTEXT/ATTRIB + blocos INSERT, com resolução de cor e **layer** |
| 3 | **Arquivo local por caminho** | `GET /extract-local?path=` | Só permitido quando `IS_FROZEN` (`.exe`); bloqueado em servidor |
| 4 | **PDFs de Orçamento + Lista** | `POST /api/importar-rec-pdf` | Parser dedicado que reconhece códigos de 6 dígitos e reconstrói itens de REC |
| 5 | **CSV da base técnica** | `POST /api/orcamento/upload` (admin), `POST /api/admin/upload-master` (admin) | Substitui a tabela de orçamento do usuário ou a tabela master |
| 6 | **CSV de regras de conversão** | `importarRegrasCSV()` no frontend | Importa regras `ORIGEM;OP_DE;ATIVO_DE;ACAO;OP_PARA;ATIVO_PARA;FATOR;ARREDONDAMENTO;VAL_MIN;VAL_MAX` |
| 7 | **Descrição em linguagem natural / imagem** | Chat de IA (`/api/gemini/*`) | Prompt de domínio (`prompt_rede_eletrica.txt`) instrui o modelo a devolver tabela `COMANDO/ID/AÇÃO/ATIVOS` |
| 8 | **Digitação manual** | Tabelas Cabos/Outros/Totalizadora e modais geradores (Cabos, Postes e Estruturas, Ramais) | Entrada direta do usuário |

**Cor e layer são dados de negócio, não estilo.** A cor do texto no projeto determina a
operação (vermelho = instalar, cinza = remover) e o layer determina linha viva e retensionamento.

---

## 6. Processamento

### 6.1 Extração
- **PDF** (`services/pdf_service.py`): percorre páginas → blocos → linhas → spans. Converte a cor
  para hex (`color_to_hex`) e decompõe as flags de fonte (`flags_decomposer`).
- **DXF** (`services/dxf_service.py`): resolve a cor real por hierarquia
  (true_color da entidade → ACI da entidade → RGB no nome do layer `TXT_RRGGBB` → true_color do
  layer → ACI do layer → preto). MTEXT é quebrado por parágrafo (`\P`) com **herança de cor** entre
  parágrafos; cor dentro de `{ }` é local e não é herdada. Saída ordenada por Y decrescente, depois X.

### 6.2 Classificação (`computeRowLogic`, `static/resumo.js`)
Converte uma linha extraída em `{entidade, operação, ativo}`:
1. Se entidade e ativo já vieram preenchidos da aba principal, **respeita a escolha do usuário** e não re-deriva.
2. `processAtivoFormula()` normaliza o texto livre para a fórmula de ativo (ex.: `"... 35 METROS"` → `35-ROCO`).
3. Operação automática pela cor: `#FF0000` → `I`; cinza → `R`; caso contrário `M`.
4. Entidade automática por uma cascata de padrões de texto, layer e cor.
5. Ajustes finais: `IP/APOIO/CERCA/RAMAIS` forçam `I`; layer `01_LV` prefixa a operação com `*`.

### 6.3 Camada unificada + motor de regras (`syncTotalizadora`, `static/resumo.js`)
As tabelas Cabos e Outros são consolidadas em uma lista única (`rawItems`), cada item com
`baseId` (`TOT-n`, mantém a rastreabilidade da linha de origem) e `origem`. Sobre essa lista roda o
motor de regras; o resultado é a **Tabela Totalizadora**, ainda editável manualmente.

### 6.4 Cálculo (`services/orcamento_calc.py:processar_calculo`)
Backend puro, sem I/O. Recebe `cabos`, `outros`, `projeto` e as linhas da base técnica.
Ver a cascata completa de resolução de ativo e a fórmula de soma em
[ARCHITECTURE.md](ARCHITECTURE.md) e as regras numeradas na seção 7 abaixo.

---

## 7. Regras de negócio identificadas no código

> Lista resumida com a origem de cada regra. O detalhamento com exemplos
> entrada → regra → saída está em [ARCHITECTURE.md](ARCHITECTURE.md), seção 8.

### RN-01 — Cor determina a operação
`#FF0000` → `I`; cinza (|R−G|<5, |G−B|<5, 20<R<230) → `R`; demais → `M`.
📍 `resumo.js:computeRowLogic`, `script.js:isGray`

### RN-02 — Layer determina linha viva e retensionamento
Layer `01_LV` → operação recebe prefixo `*`. Layers `01_RETENS`, `01_RETENS_LV`, `01_LV` com texto
preto disparam as regras de reinstalação (`TR-n` → `1-RTRn`, `n-100A` → `n-RCFU`).
📍 `resumo.js:computeRowLogic`

### RN-03 — Formato de ativo por origem
- `CABOS`: `[ATIVO] [FASE] [COMPRIMENTO]` (ex.: `CAA 2 ABC 35 m`)
- `OUTROS`: `<quantidade>-<ativo>`, separados por espaço (ex.: `3-IP 2-RECAL`)
- `*` no lugar da quantidade representa valor negativo
📍 `orcamento_calc.py:44-122`, `prompt_rede_eletrica.txt:66-72`

### RN-04 — Quantidade de ativos de um cabo (`qtdAtivos`)
Padrão: `len(fase)`. Exceções: prefixo `M…` ou `CAZ…` → sempre `1`; `CAA2` com fase de 1 caractere
→ `len(fase)+1`; `CA4` → sempre `len(fase)+1`. Linha "standalone" (só o ativo) **herda a fase da
linha seguinte**; o recálculo percorre a tabela de baixo para cima.
📍 `resumo.js:calcularQtdAtivos`, `recalcAllQtdAtivos`

### RN-05 — Motor de regras de conversão
Para cada item da camada unificada, toda regra cujo `origem` + `op_de` + `ativo_de` casam é aplicada
(campo vazio = "qualquer"; `%` = LIKE; senão regex ancorada). A regra pode trocar operação e ativo,
multiplicar por `fator`, arredondar (`NORMAL` 2 casas / `PARA CIMA` / `PARA BAIXO` / `INTEIRO`) e
limitar por `val_min`/`val_max`. **`ADIÇÃO` mantém o item original; `SUBST` o remove.**
📍 `resumo.js:2918-3006`

### RN-06 — Resolução de ativo na base técnica (cascata de 5 passos)
`ATIVO` exato → `COMPONENTE` exato → `DESC_ATIVO` exato → `DESC_ATIVO` parcial → `CÓDIGO` direto
(neste último caso os fatores são forçados para `1.0`). Sem match → o ativo entra em
`nao_encontrados`. Todo passo respeita o filtro de `origem`: linha com `origem` vazia serve para
qualquer origem; linha com `origem` preenchida só serve para a origem igual.
📍 `orcamento_calc.py:165-231`

### RN-07 — Expansão ativo → componente → códigos
Achado o ativo, toma-se o `componente` do **primeiro** match e expandem-se **todos** os códigos
daquele componente. Havendo várias linhas para o mesmo código, vence a que tem `mdo` preenchido.
📍 `orcamento_calc.py:206-231`

### RN-08 — Prioridade de projeto na base técnica
Match exato do projeto > linha genérica (`projeto` vazio) > qualquer outra linha. O campo `projeto`
aceita várias concessionárias separadas por `/`.
📍 `orcamento_calc.py:17-41`, `126-135`

### RN-09 — Soma por código (equivalente ao SOMASE)
Chave de agrupamento `(codigo, mdo)`. `I`/`*I` somam `qtd × fator_i` em `soma_i`;
qualquer outra operação soma `qtd × fator_r` em `soma_r`. Só entram no resultado as somas `> 0`,
como linhas separadas de `I` e `R`, arredondadas em 2 casas.
📍 `orcamento_calc.py:233-268`

### RN-10 — Ordenação do resultado
`MAO-DE-OBRA` (rank 1, inclui vazio e `-`) antes de `MATERIAL` (rank 2), depois operação A–Z
(`I` < `R`), depois descrição do código A–Z.
📍 `orcamento_calc.py:270-280`

### RN-11 — Precedência da base de orçamento
Edições do próprio usuário (`tabela_orcamento` com `user_id`) > tabela master sincronizada do
Supabase (`tabela_orcamento_master`) > seed inicial (`tabela_orcamento` com `user_id IS NULL`).
📍 `sync_service.py:get_merged_orcamento`

### RN-12 — Preservação de REC de outro usuário
Se um **operador** salva um REC cujo `numero_obra` pertence a outro usuário, o original é preservado
e a versão dele é gravada como `"<numero> (Cópia - <alias do e-mail>)"`. Admin sobrescreve.
Exclusão: operador só apaga REC próprio.
📍 `recs.py:save_rec`, `delete_rec`

### RN-13 — Login offline-first
Tenta Supabase (`usuarios_nuvem`) primeiro; se indisponível ou sem match, valida contra o SQLite
local (`usuarios_locais`). Senhas SHA-256 legadas são aceitas uma vez e **migradas para bcrypt**
automaticamente.
📍 `auth.py:login`, `database.py:verify_password`, `_migrate_legacy_password`

### RN-14 — Operações restritas a admin
Alterar a base de orçamento (`/api/orcamento/upload`, `/salvar`), cadastrar projeto
(`POST /api/projetos`) e todo o `/api/admin/*` exigem `role == "admin"`.
📍 `orcamento.py`, `projetos.py`, `admin.py:_require_admin`

### RN-15 — Leitura de arquivo local só no executável
`GET /extract-local` retorna "Acesso negado" quando `IS_FROZEN` é falso, impedindo leitura
arbitrária de disco quando hospedado.
📍 `upload.py:extract_local`

### RN-16 — Regras de conversão são por projeto
Persistidas com a chave `regras_conversao_<projeto_codigo>` em `configuracoes` (nuvem/servidor) e
`regras_orcamento_<projeto_codigo>` no `localStorage` (local). A nuvem tem precedência na leitura.
📍 `regras.py`, `resumo.js:carregarRegras`

### RN-17 — Itens com quantidade zero ou vazia não vão para o cálculo
📍 `resumo.js:1838-1840`

### Regras de domínio que existem **apenas** no prompt da IA
`prompt_rede_eletrica.txt` contém regras técnicas (poste de 10 m não pode ser usado em MT; CFU exige
`1-SUPL`; trafo exige `1-PR15` e 2 ou 3 `PR220`; P50 mínimo de 2 m mono / 6 m tri + 1 m por PR15;
estruturas `3` não podem existir sozinhas no poste; etc.). **Essas regras não são validadas por
código** — são apenas instruções ao modelo de linguagem. Não assuma que o sistema as aplica.

---

## 8. Saídas

| Saída | Formato | Onde |
|---|---|---|
| **Orçamento calculado** | `{resultado: [{operacao, mdo, codigo, desc_codigo, filtro, total}], nao_encontrados: [...]}` | `POST /api/orcamento/calcular` |
| **Tabela Totalizadora** | Tabela editável na tela, com destaque vermelho para ativo não encontrado na base | `static/resumo.html` |
| **REC salvo** | JSON em `historico_rec` (SQLite + Supabase), recuperável por `numero_obra` | `/api/recs`, `/api/rec/*` |
| **Arquivo `.rec` / CSV** | Download pelo navegador | `static/resultado_orcamento.html`, `resumo.js:exportarRegrasCSV` |
| **Backup** | JSON com `tabela_orcamento` + `regras` + `projetos` | `GET /api/backup/export` |
| **Diagnóstico** | Contagens do banco, caminhos, flags de ambiente | `GET /api/health` |
| **Eventos em tempo real** | WebSocket `/ws`: `load_file`, `connectivity` | `websocket_manager.py` |

---

## 9. Estado atual

Ver **[STATE.md](STATE.md)** — implementado, limitações, problemas conhecidos e pontos que precisam
de confirmação do usuário.
