# UTalk Rebuild

Projeto interno para reconstruir um sistema de atendimento estilo WhatsApp multiatendentes, com foco em:

- frontend de alta fidelidade em relacao aos prints mapeados
- backend funcional interno com persistencia em banco
- atendimento por departamentos
- chatbot real para triagem inicial
- controle de acesso por perfil `Admin` e `Operador`
- documentacao forte desde o inicio para desenvolver modulo por modulo

## Estado atual

- frontend React/Vite com varias telas-base navegaveis
- boa cobertura visual em `Conversas`, `Contatos`, `Atendentes`, `Etiquetas`, `Respostas rapidas`, `Canais`, `Templates` e `Campanhas`
- ainda sem backend real acoplado ao frontend atual
- mapeamento funcional e arquitetura ja iniciados em `docs/`

## Documentos principais

- [Especificacao Mestre](</W:/DOCUMENTOS ESCRITORIO/INSTALACAO SISTEMA/python/utalk_rebuild/docs/master-project-blueprint.md:1>)
- [Mapeamento Funcional de Botoes](</W:/DOCUMENTOS ESCRITORIO/INSTALACAO SISTEMA/python/utalk_rebuild/docs/functional-button-mapping.md:1>)
- [Inventario de Telas](</W:/DOCUMENTOS ESCRITORIO/INSTALACAO SISTEMA/python/utalk_rebuild/docs/screen-inventory.md:1>)
- [Notas de Implementacao](</W:/DOCUMENTOS ESCRITORIO/INSTALACAO SISTEMA/python/utalk_rebuild/docs/implementation-notes.md:1>)
- [Fase 0 - Fundacao tecnica](</W:/DOCUMENTOS ESCRITORIO/INSTALACAO SISTEMA/python/utalk_rebuild/docs/phase-0-foundation.md:1>)
- [Fase 1 - Inicio Conversas](</W:/DOCUMENTOS ESCRITORIO/INSTALACAO SISTEMA/python/utalk_rebuild/docs/phase-1-conversations-start.md:1>)

## Regras do projeto

- nunca copiar dados reais dos prints para o codigo
- mascarar sem excecao nomes, telefones, e-mails e identificadores
- evitar `localhost` hardcoded em callbacks, redirects e links internos
- usar URL/IP publico da maquina hospedeira via configuracao
- quando uma funcionalidade afetar mais de uma tela, atualizar todos os pontos relacionados

## Proxima direcao

Seguir desenvolvimento por modulo fechado:

1. arquitetura base e backend interno
2. conversas
3. contatos
4. atendentes e controle de acesso
5. dashboards
6. modulos administrativos restantes

## Comandos uteis

- Frontend: `npm run dev:web`
- API (desenvolvimento): `npm run dev:api`
- API (migracao): `npm run api:migrate`
- Infra (postgres/redis): `npm run infra:up`
