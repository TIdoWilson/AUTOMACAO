# UTalk Rebuild - Inventario de Telas

Base usada:
- Espelhamento parcial do frontend original em `W:\DOCUMENTOS ESCRITORIO\INSTALACAO SISTEMA\python\site_espelhado\app-utalk.umbler.com`
- 85 prints extraidos em `W:\DOCUMENTOS ESCRITORIO\INSTALACAO SISTEMA\python\site_espelhado\prints_extraidos`

Politica de dados sensiveis:
- os prints reais ficam fora do projeto frontend
- nomes, telefones, e-mails e identificadores reais nao devem ser copiados para o codigo
- qualquer placeholder visual no projeto deve ser sintetico e neutro

Objetivo desta documentacao:
- mapear telas, menus, botoes e estados visuais
- orientar o refinamento do frontend tela por tela
- separar "estrutura visual" de "funcionalidade real"

## Rotas Base Ja Estruturadas

As seguintes rotas ja possuem base visual em `src/App.jsx`:

- `/login`
- `/`
- `/contacts`
- `/chatbots`
- `/bulksend`
- `/dashboard`
- `/settings`
- `/settings/agents`
- `/settings/templates`

## Mapa Visual Por Grupo

### 1. Login

Prints principais:
- `imagem - 2026-04-14T102236.652.png`

Elementos identificados:
- logo `umbler talk`
- login social Google
- login social Facebook
- campos `E-mail` e `Senha`
- checkbox `Manter-me conectado`
- link `Esqueci minha senha`
- CTA `Entrar`
- link `Cadastre-se`
- painel promocional lateral com mockup do produto

Status no projeto:
- base visual implementada
- falta refinar espacos, proporcoes, tipografia e mockup lateral

### 2. Conversas / Inbox

Prints principais:
- `imagem (1).png`
- `imagem (30).png` a `imagem (56).png`

Subestados identificados:
- lista de conversas
- conversa aberta
- filtros de lista
- menu de distribuicao/transferencia
- drawer/modal de contato
- composer expandido
- anexos
- notas
- agendamento
- acoes da topbar da conversa

Elementos recorrentes:
- sidebar vertical azul
- busca por nome ou telefone
- abas `Entrada`, `Esperando`, `Finalizados`
- banner trial
- cards de conversa com avatar, horario e categoria
- topbar da conversa com acoes
- composer com anexos, biblioteca, emoji e canais

Botoes confirmados por print:
- detalhes do contato
- detalhes da conversa
- transferir conversa
- ativar atendimento
- agendar envio de mensagem
- finalizar conversa
- ocultar essa janela

Status no projeto:
- base forte implementada
- falta refinar overlays, popovers, drawer lateral, estados hover e topbar completa

### 3. Contatos

Prints principais:
- `imagem (57).png`
- `imagem (58).png`

Elementos identificados:
- cabecalho `Contatos`
- busca
- filtros `Todos`, `Etiquetas`
- ordenacao
- CTA `Adicionar contato`
- tabela/lista de contatos
- acao por linha

Status no projeto:
- base implementada
- falta construir estados vazios, detalhes do contato e menu por linha

### 4. Chatbots

Prints principais:
- `imagem (59).png`
- `imagem (60).png`
- `imagem (61).png`
- `imagem (62).png`

Elementos identificados:
- listagem de chatbots
- estado vazio
- canvas visual de fluxo
- cards de blocos conectados
- CTA lateral/superior

Status no projeto:
- tela-base implementada
- falta reconstruir canvas de fluxo e cards do builder

### 5. Envio de Campanhas

Prints principais:
- `imagem (63).png`
- `imagem (86).png`

Elementos identificados:
- titulo `Envio de campanha de marketing`
- estado vazio
- CTA principal
- entrada dedicada no menu lateral

Status no projeto:
- base implementada
- falta tela completa de criacao/edicao

### 6. Relatorios / Dashboard

Prints principais:
- `imagem (64).png` a `imagem (77).png`
- `imagem (87).png`

Subestados identificados:
- cards metricos
- graficos de barras
- graficos de linha
- grafico de pizza
- filtros por periodo
- lista de itens/atividades
- drawers/paineis laterais

Status no projeto:
- base visual implementada
- falta fidelidade dos cards, tabs, filtros e tipos de grafico

### 7. Configuracoes - indice

Prints principais:
- `imagem (88).png`
- `imagem (89).png`
- `imagem (90).png`

Elementos identificados:
- lista vertical de modulos
- cabecalho `Configuracoes`
- acesso para paginas filhas

Itens confirmados:
- Canais de atendimento
- Financeiro
- Setores
- Agentes de IA
- Bases de conhecimento
- Atendentes
- Etiquetas
- Chatbots
- Respostas rapidas
- Templates WhatsApp Business API

Status no projeto:
- base implementada
- falta hierarquia visual mais fiel e icones mais proximos do original

### 8. Canais, Setores, Bases, Atendentes, Etiquetas

Prints principais:
- `imagem (91).png`
- `imagem (92).png`
- `imagem (93).png`
- `imagem (94).png`
- `imagem (95).png`
- `imagem (96).png`
- `imagem (97).png`
- `imagem (98).png`
- `imagem (99).png`

Subtelas confirmadas:
- canais de atendimento
- setores
- base de conhecimento
- atendentes
- dropdown de permissao
- etiquetas
- importacao rapida
- modal de resposta rapida

Status no projeto:
- `Atendentes` e `Templates` ja iniciados
- demais subtelas ainda nao implementadas

### 9. Configuracao de Organizacao

Prints principais:
- `imagem (79).png`
- `imagem (80).png`
- `imagem - 2026-04-14T102213.175.png`
- `imagem - 2026-04-14T102215.663.png`
- `imagem - 2026-04-14T102218.584.png`
- `imagem - 2026-04-14T102221.084.png`
- `imagem - 2026-04-14T102223.291.png`
- `imagem - 2026-04-14T102225.513.png`
- `imagem - 2026-04-14T102227.916.png`
- `imagem - 2026-04-14T102230.175.png`
- `imagem - 2026-04-14T102233.882.png`

Elementos identificados:
- dados da organizacao
- avatar/logo
- campos de formulario
- secoes com abas
- passos/onboarding
- listas administrativas

Status no projeto:
- ainda nao implementado

## Inventario Inicial de Acoes Confirmadas

### Sidebar

Mapeadas visualmente:
- Conversas
- Contatos
- Chatbots
- Campanhas
- Relatorios
- Configuracoes

### Conversa aberta

Acoes confirmadas por prints:
- abrir detalhes do contato
- abrir detalhes da conversa
- transferir conversa
- ativar/privar atendimento
- agendar mensagem
- finalizar conversa
- fechar painel

### Configuracoes

Acoes confirmadas:
- abrir modulo
- criar canal
- convidar atendente
- alterar permissao
- alterar reatribuicao
- criar etiqueta
- criar/importar resposta rapida

## Proxima Sequencia Recomendada

1. Refino completo da tela de conversas, incluindo topbar e overlays.
2. Refino da tela de contatos e drawer de detalhes.
3. Reconstrucao da tela de configuracoes indice.
4. Subtelas administrativas uma a uma:
   - atendentes
   - etiquetas
   - respostas rapidas
   - canais
   - setores
   - bases de conhecimento
5. Dashboard/relatorios com todos os estados de filtro e graficos.
6. Configuracao de organizacao e onboarding final.
