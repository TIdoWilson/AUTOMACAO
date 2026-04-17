# UTalk Rebuild - Fase 1 (inicio do modulo Conversas)

Este documento descreve o que ja foi implementado no backend para iniciar o modulo de conversas com persistencia real.

## Regras de negocio ja cobertas

- entrada de mensagem do cliente por endpoint interno
- triagem inicial por chatbot com selecao de departamento
- conversa vai para fila (`queued`) apos selecao
- assumir conversa muda para `in_progress`
- sair da conversa remove apenas o participante atual
- se nao houver participantes ativos e o cliente responder, conversa volta para fila do ultimo departamento
- finalizar conversa muda para `finalized`
- nova mensagem apos finalizacao reinicia fluxo em `bot_triage`
- `waiting_customer` tratado como estado operacional real

## Endpoints implementados

Autenticacao e base:

- `GET /health`
- `POST /auth/bootstrap-admin`
- `POST /auth/login`
- `POST /auth/refresh`
- `POST /auth/logout`
- `GET /auth/me`

Admin e acesso:

- `GET /settings/modules` (admin-only)
- `GET /departments`
- `POST /departments` (admin-only)
- `GET /users` (admin-only)
- `POST /users` (admin-only)

Contatos:

- `GET /contacts`
- `POST /contacts`
- `PATCH /contacts/:contactId`

Conversas:

- `POST /conversations/incoming-message` (token interno)
- `GET /conversations`
- `GET /conversations/:conversationId/messages`
- `POST /conversations/:conversationId/assume`
- `POST /conversations/:conversationId/leave`
- `POST /conversations/:conversationId/finalize`
- `POST /conversations/:conversationId/waiting-customer`
- `POST /conversations/:conversationId/transfer`
- `POST /conversations/:conversationId/messages`

Dashboard inicial:

- `GET /dashboard/summary`

## Seguranca aplicada nesta etapa

- RBAC backend com perfis `admin` e `operator`
- modulo de configuracoes bloqueado para operador
- token interno para endpoint de entrada de mensagem do canal
- auditoria inicial para eventos criticos

## Observacoes operacionais

- limite de anexo ja refletido no schema: `100 MB` por arquivo
- URLs publicas sao validadas para impedir `localhost` e `127.0.0.1`
- dados de exemplo permanecem mascarados

## Proximo passo da Fase 1

1. Integrar frontend de conversas com estes endpoints.
2. Implementar upload real de anexos (storage e validacao MIME).
3. Adicionar atribuir responsavel explicito na conversa.
4. Expandir fluxo de chatbot para menu configuravel por admin.
