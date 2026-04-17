# UTalk Rebuild - Fase 0 (Fundacao tecnica)

Este documento descreve o que foi iniciado em codigo para a Fase 0 e como executar localmente.

## Entregas iniciadas

- API interna em `apps/api`
- autenticacao base com `Admin` e `Operador`
- autorizacao por perfil no backend
- schema inicial de banco PostgreSQL
- migracao versionavel por script
- trilha de auditoria inicial (`audit_logs`)
- validacao de ambiente com bloqueio de URLs publicas em `localhost/127.0.0.1`
- modulo de departamentos e usuarios administrativos iniciais

## Estrutura criada

- `apps/api/src/config/env.js`
- `apps/api/src/db/schema.sql`
- `apps/api/src/db/migrate.js`
- `apps/api/src/middleware/auth.js`
- `apps/api/src/routes/auth.js`
- `apps/api/src/routes/departments.js`
- `apps/api/src/routes/users.js`
- `apps/api/src/routes/settings.js`
- `apps/api/src/routes/health.js`
- `apps/api/src/app.js`
- `apps/api/src/index.js`

## Variaveis de ambiente

Baseie-se em `apps/api/.env.example`.

Regra importante:

- `APP_PUBLIC_URL`, `WEB_PUBLIC_URL`, `API_PUBLIC_URL` e `WS_PUBLIC_URL` nao podem apontar para `localhost` nem `127.0.0.1`.

## Execucao

No backend:

```bash
cd apps/api
npm install
npm run migrate
npm run dev
```

No frontend (em outro terminal):

```bash
npm run dev:web
```

Infraestrutura (na raiz do projeto):

```bash
npm run infra:up
```

Portas usadas pela stack de desenvolvimento:

- PostgreSQL: `127.0.0.1:55432`
- Redis: `127.0.0.1:56379`
- API: `127.0.0.1:4000`
- Frontend Vite: `127.0.0.1:5173`

## Bootstrap inicial de admin

Depois de subir a API, execute:

`POST /auth/bootstrap-admin`

Payload exemplo sintetico:

```json
{
  "organizationName": "Organizacao Interna",
  "departmentName": "Geral",
  "name": "Admin Interno",
  "email": "admin@empresa.local",
  "password": "SenhaSegura123"
}
```

## Validacoes recomendadas da Fase 0

1. `GET /health` retorna status `ok`.
2. `POST /auth/login` retorna `accessToken` e cookie de refresh.
3. `GET /auth/me` funciona com token Bearer.
4. `GET /settings/modules`:
- retorna 200 para `Admin`
- retorna 403 para `Operador`
5. `POST /departments` funciona para `Admin` e bloqueia `Operador`.
6. `POST /users` permite criar operador e vinculo a departamentos.

## Proximo passo apos Fase 0

Iniciar Fase 1 (`Conversas`) com persistencia real de:

- filas por departamento
- participantes de conversa
- mensagens e anexos por conversa
- fluxo de finalizacao/retorno
- chatbot de triagem com roteamento por departamento
