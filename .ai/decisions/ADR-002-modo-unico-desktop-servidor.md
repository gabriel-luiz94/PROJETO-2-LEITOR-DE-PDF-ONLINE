# ADR-002 — Base de código única para desktop e servidor, chaveada por `APP_MODE`

> ⚠️ **ADR retroativo.** Decisão já implementada quando a documentação `.ai` foi criada
> (2026-09-18), reconstruída a partir do código. Sem registro histórico da deliberação original.

## Contexto

O sistema é distribuído de duas formas simultâneas:

1. **Aplicativo desktop** — executável Windows gerado por PyInstaller (`build.bat`), abrindo uma
   janela nativa via `pywebview`, com banco SQLite ao lado do `.exe` e capacidade de ler arquivos
   do disco local.
2. **Serviço web** — hospedado na nuvem (workflow para Fly.io), acessado pelo navegador, sem acesso
   ao disco do usuário e com CORS restrito.

O cabeçalho de `app.py` declara isso explicitamente:
*"Suporta dois modos de operação: desktop … server …"*.

## Problema

As duas formas de distribuição exigem comportamentos incompatíveis — CORS, bootstrap do processo,
resolução de caminhos e, principalmente, permissão para ler arquivos locais — sem manter dois
códigos separados.

## Alternativas consideradas

`[NÃO DETERMINADO PELO CÓDIGO]` — as alternativas efetivamente avaliadas não estão registradas.

## Decisão

**Uma única base de código, com o comportamento chaveado em tempo de execução pela variável de
ambiente `APP_MODE`** (`desktop` | `server`), complementada pela flag `IS_FROZEN`
(`sys.frozen`, verdadeira dentro do executável PyInstaller).

Os dois eixos têm papéis distintos:

- **`APP_MODE`** decide política: CORS, bootstrap (janela nativa vs. uvicorn headless), origem da
  informação de versão, e se o monitor de conectividade deve rodar.
- **`IS_FROZEN`** decide capacidade física e segurança: resolução de caminhos
  (`_MEIPASS` vs. diretório do fonte) e a **trava de leitura de arquivo local**.

A separação é deliberada: `GET /extract-local` é bloqueado por `IS_FROZEN`, não por `APP_MODE`
(`routers/upload.py:46`). Um servidor mal configurado com `APP_MODE=desktop` continua incapaz de ler
o disco, porque não é um executável congelado.

## Consequências

### Positivas
- Uma única base de código, um único conjunto de rotas e regras de negócio
- Correção de regra de negócio vale para as duas distribuições automaticamente
- A trava de segurança não depende de configuração correta de ambiente

### Negativas / custos aceitos
- Condicionais de modo espalhados por `app.py`, `config.py`, `upload.py`, `update.py` e
  `connectivity_monitor.py`
- O comportamento em produção depende de variáveis de ambiente corretas — e o default falha para o
  lado do desktop (ver abaixo)
- Dependências de desktop (`pywebview`) ficam em `requirements.txt` mesmo no deploy servidor

### Armadilhas conhecidas (verificadas no código)
1. `config.py:84` — `APP_MODE = os.environ.get("APP_MODE", "desktop" if IS_FROZEN else "desktop")`:
   **os dois ramos do ternário são idênticos**. O default é sempre `"desktop"`, inclusive quando
   hospedado. O modo servidor exige `APP_MODE=server` explícito no ambiente.
2. `app.py:38-46` — em modo servidor sem `CORS_ORIGINS` definido, o CORS cai para `["*"]` **com**
   `allow_credentials=True`. O próprio código chama isso de "fallback temporário".
3. `app.py:148` — o bootstrap headless é acionado por `HEADLESS=1` **ou** `APP_MODE == "server"`,
   dando duas formas de obter o mesmo efeito.

### Implicações para implementações futuras
1. **Toda funcionalidade nova deve declarar em qual modo faz sentido.** Se for exclusiva de um,
   proteja com a flag correta.
2. **Capacidade física e segurança → `IS_FROZEN`. Política e ambiente → `APP_MODE`.** Não troque
   um pelo outro; a trava de `/extract-local` depende dessa distinção.
3. Dependências só-desktop devem ser importadas **dentro** da função que as usa (como
   `app.py:168` faz com `webview`), nunca no topo do módulo.
4. Ao mexer em CORS ou no bootstrap, teste os dois modos.

## Evidência no código

- `app.py:1-7` — docstring declarando os dois modos
- `app.py:38-54` — CORS por modo
- `app.py:146-179` — bootstrap headless vs. janela nativa com fallback para navegador
- `config.py:27-68` — resolução de caminhos por `IS_FROZEN` / `_MEIPASS`
- `config.py:80-84` — definição de `IS_FROZEN` e `APP_MODE`
- `routers/upload.py:42-47` — trava de `/extract-local` por `IS_FROZEN`
- `routers/update.py:27-51` — origem da versão por modo
- `services/connectivity_monitor.py:63-64` — não inicia em modo servidor

## Data

2026-09-18 (documentação retroativa; data da decisão original desconhecida)

## Status

**ACEITA**
