# UTalk Rebuild - Especificacao Mestre

## 1. Objetivo do projeto

Construir um sistema interno de atendimento multiatendentes, com experiencia visual de alta fidelidade em relacao aos prints capturados, backend funcional interno, persistencia em banco, chatbot real para triagem por departamento, dashboards operacionais e controle de acesso por perfil.

Este documento passa a ser a referencia principal para arquitetura, fases de entrega, regras de negocio, seguranca, LGPD, conformidade e criterios de aceite.

## 2. Decisoes fechadas com o responsavel do projeto

- O projeto final precisa ter `backend funcional interno`.
- O criterio de conclusao e `fluxo completo ate salvar em banco`.
- Prioridade de desenvolvimento:
- `Conversas`
- `Contatos`
- `Atendentes`
- Fidelidade visual: `alta`.
- Perfis iniciais:
- `Admin`
- `Operador`
- Restricao de acesso:
- configuracoes ficam bloqueadas para `Operador`
- somente `Admin` acessa modulos administrativos
- Fluxo operacional principal:
- cliente envia mensagem
- chatbot responde
- cliente escolhe departamento
- conversa entra na fila do departamento
- operadores vinculados ao departamento conseguem ver a conversa
- um operador pode assumir, adicionar responsavel ou transferir
- `Finalizar conversa`:
- move para `Finalizados`
- proxima mensagem do cliente reinicia o bot e a selecao de departamento
- `Sair da conversa`:
- remove apenas o operador atual da conversa
- nao finaliza o atendimento
- se ainda houver outro operador participante, a conversa continua com ele
- se nao houver mais operador e o cliente escrever novamente, a conversa volta para a fila do ultimo departamento selecionado
- `Marcar como esperando` significa `Esperando cliente`
- Anexos:
- todos os tipos permitidos
- limite assumido: `100 MB por arquivo`
- Biblioteca de midia:
- `por conversa`
- Respostas rapidas:
- suporte a variaveis dinamicas
- Chatbot:
- execucao real de regras
- Campanhas:
- fora do escopo atual
- Dashboards:
- devem refletir corretamente os indicadores operacionais reais
- Todas as telas dos prints sao obrigatorias
- Todos os dados reais permanecem mascarados
- Ha interesse em conexao real por QR code desde o inicio

## 3. Restricao importante sobre o canal WhatsApp

### Diretriz tecnica

O dominio do produto nao deve ficar acoplado a um unico tipo de integracao de canal. Mesmo que a primeira conexao real seja por QR code, a arquitetura precisa isolar esse canal em um adaptador proprio.

### Risco operacional e juridico

Uma conexao baseada em sessao web/QR code nao oficial pode sofrer:

- bloqueios da plataforma
- invalidez repentina da sessao
- mudancas nao anunciadas na interface/protocolo
- risco de indisponibilidade do atendimento
- possivel conflito com termos de uso da plataforma

### Regra de projeto

Por isso, a integracao por QR code deve nascer como `modulo substituivel`, nunca como regra central da aplicacao. O sistema precisa continuar organizado para futura troca por um conector oficial ou homologado, sem reescrever o core.

### Observacao de conformidade

Este documento traz diretrizes tecnicas e de governanca. Validacoes juridicas finais devem ser revisadas com apoio juridico/contabil da operacao antes da entrada em producao.

## 4. Arquitetura alvo do projeto

## 4.1 Visao geral

Arquitetura recomendada em modulos:

- `web-app`
- frontend React responsavel por UI, navegacao, estado de tela e consumo da API
- `api-app`
- backend interno responsavel por autenticacao, regras de negocio, persistencia e autorizacao
- `realtime-gateway`
- canal de atualizacao em tempo real para mensagens, filas, status e dashboards
- `chatbot-engine`
- execucao dos fluxos de triagem, menus, regras e roteamento
- `channel-connector`
- servico isolado para conexao do canal por QR code e entrega/recebimento de mensagens
- `jobs-worker`
- processamento assincrono de tarefas como anexos, reprocessamentos, metricas e notificacoes
- `postgres`
- banco principal transacional
- `redis`
- fila, cache, locks e eventos temporarios
- `object-storage`
- armazenamento de anexos por conversa, com metadados em banco

## 4.2 Regras de comunicacao

- frontend nunca acessa banco diretamente
- frontend conversa apenas com a API
- API publica eventos para gateway realtime
- channel connector nao conhece telas nem regras de interface
- chatbot engine nao depende da implementacao visual
- dashboards leem fatos operacionais persistidos, nao estado efemero do frontend

## 4.3 Estrutura sugerida de pastas

```text
utalk_rebuild/
  apps/
    web/
    api/
    worker/
    connector/
  packages/
    shared-types/
    ui/
    domain/
    config/
  docs/
```

## 4.4 Regra de URLs e redirecionamentos

- nao usar `localhost` hardcoded em rotas, callbacks, redirects ou URLs de websocket
- toda URL publica deve nascer de configuracao
- variavel obrigatoria sugerida:
- `APP_PUBLIC_URL=http://IP_DA_MAQUINA:PORTA`
- variaveis derivadas:
- `API_PUBLIC_URL`
- `WS_PUBLIC_URL`
- `WEB_PUBLIC_URL`
- callbacks de login, QR code, arquivos e deep links devem usar a URL publica configurada

## 5. Modelo funcional central

## 5.1 Entidades principais

- `Organization`
- `Department`
- `User`
- `UserDepartment`
- `Role`
- `Contact`
- `ContactTag`
- `Conversation`
- `ConversationParticipant`
- `ConversationRouting`
- `Message`
- `Attachment`
- `ChatbotFlow`
- `ChatbotNode`
- `ChatbotExecution`
- `QuickReply`
- `ChannelSession`
- `AuditLog`
- `MetricSnapshot`

## 5.2 Estados principais da conversa

- `bot_triage`
- `queued`
- `in_progress`
- `waiting_customer`
- `finalized`

## 5.3 Regras de negocio obrigatorias

- uma conversa pertence a um departamento atual
- varios operadores do mesmo departamento podem visualizar a fila
- uma conversa pode ter zero, um ou varios participantes ativos
- `assumir conversa` adiciona o operador atual como participante
- `sair da conversa` remove apenas esse participante
- `finalizar conversa` muda o estado para `finalized`
- nova mensagem em conversa `finalized` cria nova rodada de triagem com chatbot
- nova mensagem em conversa nao finalizada sem participantes ativos devolve a conversa para a fila do ultimo departamento
- `esperando cliente` significa pausa operacional aguardando retorno do contato

## 6. Fases e modulos de entrega

Regra de governanca:

- um modulo so avanca quando estiver implementado, integrado, testado e documentado
- qualquer impacto colateral em outras telas deve ser corrigido no mesmo ciclo

## Fase 0 - Fundacao tecnica

### Objetivo

Preparar a base para impedir retrabalho nos modulos prioritarios.

### Entregas

- estrutura monorepo ou separacao clara `web/api/worker/connector`
- autenticacao base com `Admin` e `Operador`
- autorizacao por middleware e guards
- base de dados e migracoes
- padrao de eventos realtime
- object storage para anexos
- padrao de configuracao por IP/URL publica
- trilha de auditoria
- observabilidade inicial

### Criterio de aceite

- login funcional
- sessao persistida
- roles aplicadas
- nenhuma dependencia de `localhost`
- banco versionado por migracao
- ambiente sobe usando IP da maquina hospedeira

## Fase 1 - Conversas

### Objetivo

Fechar o modulo mais critico do produto do inicio ao fim.

### Escopo

- inbox por departamento
- filas `Entrada`, `Esperando cliente`, `Finalizados`
- assumir conversa
- sair da conversa
- adicionar responsavel
- transferir conversa
- finalizar conversa
- detalhes da conversa
- detalhes do contato
- mensagens em tempo real
- anexos por conversa
- biblioteca de midia por conversa
- respostas rapidas com variaveis
- chatbot real de triagem
- execucao de regras por departamento

### Regras especificas

- toda mensagem recebida precisa persistir no banco
- todo evento operacional precisa gerar `AuditLog`
- anexos ficam associados a conversa e mensagem
- limite de anexo: `100 MB por arquivo`
- ao finalizar, proxima mensagem reinicia bot
- ao sair sem finalizar, a conversa continua viva

### Dependencias

- Fase 0 completa

### Criterio de aceite

- fluxo completo real:
- entrada do cliente
- chatbot
- selecao de departamento
- entrada na fila
- operador assume
- conversa recebe mensagens
- operador transfere ou sai
- operador finaliza
- nova mensagem do cliente respeita a regra correta
- tudo salvo em banco
- tudo refletido em realtime

## Fase 2 - Contatos

### Objetivo

Tornar os contatos uma base real e utilizavel pelo atendimento.

### Escopo

- cadastro, edicao e consulta
- historico consolidado por contato
- etiquetas
- observacoes internas
- busca rapida
- deduplicacao por identificador do canal
- relacao contato <-> conversas

### Criterio de aceite

- contato criado e editado persiste em banco
- busca funciona por nome, identificador e etiqueta
- detalhes do contato refletem historico real
- mudancas impactam corretamente a tela de conversas

## Fase 3 - Atendentes, departamentos e acesso

### Objetivo

Fechar o controle administrativo minimo do sistema.

### Escopo

- cadastro de atendentes
- ativacao/desativacao
- vinculo por departamento
- papel `Admin` e `Operador`
- restricao de configuracoes por perfil
- filas visiveis conforme departamento
- politicas de permissao por rota e acao

### Criterio de aceite

- operador nao acessa configuracoes administrativas
- admin gerencia departamentos e atendentes
- visibilidade de filas respeita vinculos reais
- todas as acoes sensiveis ficam auditadas

## Fase 4 - Dashboards

### Objetivo

Transformar os relatorios dos prints em paineis reais do sistema.

### Indicadores minimos

- conversas atendidas
- tempo medio de resposta
- contatos ativos
- produtividade por atendente
- volume por departamento
- tempo em fila
- conversas finalizadas
- conversas aguardando cliente

### Regra

Dashboard so entra depois que conversas, contatos e atendentes estiverem gerando dados reais confiaveis.

## Fase 5 - Modulos administrativos restantes

### Escopo

- etiquetas avancadas
- respostas rapidas administrativas
- configuracoes de canal
- templates
- bases de conhecimento
- agentes de IA
- demais telas obrigatorias dos prints

## 7. Matriz de modulos x impacto cruzado

- `Conversas` impacta `Contatos`, `Dashboards`, `Atendentes`, `Chatbot`, `Anexos`
- `Contatos` impacta `Conversas`, `Dashboards`, `Etiquetas`
- `Atendentes` impacta `Conversas`, `Filas`, `Seguranca`, `Dashboards`
- `Departamentos` impacta `Chatbot`, `Conversas`, `Permissoes`, `Dashboards`
- `Chatbot` impacta `Conversas`, `Filas`, `Dashboards`

Regra: toda entrega deve revisar os modulos relacionados acima antes de ser considerada concluida.

## 8. Seguranca

## 8.1 Controle de acesso

- RBAC minimo com `Admin` e `Operador`
- principio do menor privilegio
- checagem no backend, nunca apenas no frontend
- sessao expirada automaticamente por inatividade configuravel

## 8.2 Credenciais e sessao

- cookies `HttpOnly`, `Secure` e `SameSite`
- rotacao de segredo e invalidez de sessao em logout
- armazenamento seguro de segredos em variaveis de ambiente

## 8.3 Anexos

- validacao de tamanho, MIME e extensao
- nome fisico aleatorio, nunca confiar no nome enviado pelo usuario
- hash do arquivo
- escaneamento antimalware recomendado antes de liberar download
- controle de download por permissao

## 8.4 Logs e auditoria

- toda acao critica gera `AuditLog`
- nao registrar corpo completo de mensagem em logs tecnicos por padrao
- mascarar identificadores sensiveis em logs operacionais

## 8.5 Realtime e integracoes

- autenticar conexoes websocket
- isolamento do conector de canal
- retentativa controlada e circuit breaker no canal QR
- alarmes para queda de sessao do canal

## 9. LGPD e conformidade

## 9.1 Principios que o sistema deve respeitar

- finalidade
- adequacao
- necessidade
- seguranca
- prevencao
- responsabilizacao

## 9.2 Medidas tecnicas obrigatorias

- mascaramento de dados sensiveis em ambiente de desenvolvimento e documentacao
- trilha de auditoria de acesso a conversas e contatos
- segregacao por perfil
- exclusao logica e politica de retencao
- mecanismo de anonimizar ou excluir dados quando aplicavel
- exportacao estruturada de dados por contato/conversa quando necessario
- criptografia em transito
- criptografia em repouso recomendada para banco, backups e storage

## 9.3 Governanca recomendada

- definir base legal por tipo de uso dos dados com apoio juridico
- manter politica interna de retencao
- documentar incidentes e resposta a incidente
- revisar contrato e politica do canal integrado

## 10. Testes e qualidade

## 10.1 Piramide minima

- testes unitarios de regras de negocio
- testes de integracao da API
- testes end-to-end dos fluxos criticos
- testes de permissao por perfil

## 10.2 Fluxos criticos que devem virar testes obrigatorios

- login admin
- login operador
- entrada de nova mensagem
- execucao do chatbot
- roteamento para departamento
- assumir conversa
- sair da conversa
- transferir conversa
- finalizar conversa
- nova mensagem apos finalizacao
- nova mensagem apos abandono sem finalizacao
- upload de anexo
- consulta da biblioteca de midia
- bloqueio de configuracao para operador

## 10.3 Regra de passagem de fase

Um modulo so pode ser considerado concluido quando:

- backend funcional estiver pronto
- tela correspondente estiver pronta
- fluxo salvo em banco
- regras de permissao cobrirem o modulo
- testes essenciais passarem
- documentacao do modulo estiver atualizada

## 11. Fora do escopo atual

- modulo de campanhas
- dependencia estrutural de conectores especificos no core
- exposicao de dados reais em exemplos, fixtures ou docs

## 12. Proxima execucao recomendada

1. travar a arquitetura da Fase 0 em codigo
2. definir schema inicial do banco
3. implementar autenticacao e roles
4. implementar modulo de departamentos
5. iniciar modulo `Conversas` com backend real
