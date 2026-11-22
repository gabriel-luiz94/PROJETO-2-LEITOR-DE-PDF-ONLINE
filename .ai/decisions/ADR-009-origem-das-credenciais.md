# ADR-009 — Origem das credenciais

## Contexto

O Assistente usa APIs de vários provedores. OpenAI, DeepSeek, Qwen, OpenRouter e Ollama compartilham o formato da API, mas cada serviço recebe uma credencial diferente. A autenticação do OrçaCore também precisa funcionar em servidores e em instalações desktop.

## Problema

Uma URL compatível aceita do cliente podia receber a credencial compartilhada da OpenAI. O servidor também podia operar com a chave JWT publicada no projeto.

## Decisão

- A chave de ambiente `OPENAI_API_KEY` só autentica no endereço oficial da OpenAI. Credenciais antigas do SQLite, gravadas sem origem, ficam fora dessa seleção. A chave pessoal enviada pelo navegador vale para o pedido.
- Os destinos OpenAI-compatíveis são validados antes da conexão e o cliente segue apenas o endereço selecionado. Presets públicos constam em `services/provedores_ia.py`. O Ollama local é permitido no desktop e endereços adicionais dependem da configuração da instalação.
- O administrador grava a chave do reserva junto do endereço e do provedor. A troca de destino exige a credencial correspondente.
- O servidor exige `JWT_SECRET` configurado. O desktop usa esse segredo quando configurado e gera um segredo persistente por perfil nos demais casos.

## Consequências

Instalações que usavam o JWT padrão exigem novo login. Para usar uma chave OpenAI antiga, o usuário pode informá-la na interface ou o administrador pode configurar uma chave de ambiente. Novas integrações que aceitam URLs compatíveis devem usar a mesma validação de destino e uma chave específica para o provedor.

## Evidência no código

`config.py`, `middleware/auth_middleware.py`, `services/provedores_ia.py`, `routers/ai_chat.py` e `routers/assistente.py`.

## Data e status

2026-11-21 · **ACEITA** · [TASK-217](../tasks/TASK-217-21-11-2026.md).
