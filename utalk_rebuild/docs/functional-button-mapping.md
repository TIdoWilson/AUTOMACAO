# UTalk Rebuild - Mapeamento Funcional de Telas e Botoes

## Fontes inspecionadas

- 85 prints do pacote `W:\DOCUMENTOS ESCRITORIO\INSTALACAO SISTEMA\python\site_espelhado\site_espelhado.zip`
- 85 imagens extraidas em `W:\DOCUMENTOS ESCRITORIO\INSTALACAO SISTEMA\python\site_espelhado\prints_extraidos`
- Codigo atual em `src/` (componentes, rotas e estados locais)

## Regra de privacidade

- Nenhum dado real dos prints foi copiado para o codigo.
- Todos os exemplos continuam sinteticos (`empresa.local`, nomes genericos, telefone oculto).

## Legenda de status

- `Conectado (mock local)`: botao executa fluxo no frontend com estado React.
- `Parcial`: abre UI/estado, mas sem fluxo completo esperado do produto final.
- `Nao implementado`: existe no print esperado, mas ainda nao existe funcionalmente no codigo.

## Resposta direta da auditoria

- Nao, **nem todos os botoes de todas as telas estao conectados**.
- Nao, **nem todas as funcionalidades estao desenvolvidas**.
- A base atual esta forte em `Conversas`, `Contatos`, `Atendentes`, `Etiquetas`, `Respostas rapidas`, `Canais`, `Templates` e `Campanhas`, mas ainda com dependencia de mock local e sem backend real.

## Decisoes de escopo que substituem o mapeamento anterior

- este projeto agora deve evoluir para `backend funcional interno`
- prioridade pratica:
- `Conversas`
- `Contatos`
- `Atendentes`
- `Campanhas` saem do foco atual
- dashboards continuam obrigatorios, mas entram depois dos tres modulos prioritarios
- configuracoes administrativas devem ficar restritas a `Admin`
- anexos e biblioteca de midia passam a ser tratados como recursos reais por conversa
- chatbot deixa de ser apenas visual e passa a ser modulo funcional obrigatorio para triagem
- redirecionamentos, callbacks e links operacionais nao podem depender de `localhost`

## Mapeamento por tela

### 1) Login (`/login`)

- `Entrar com conta da equipe` -> Nao implementado (botao sem acao).
- `Entrar com o Facebook` -> Nao implementado (botao sem acao).
- `Entrar` -> Nao implementado (botao sem submit/login real).
- `Esqueci minha senha` -> Parcial (link aponta para `/`).
- `Ainda nao possui conta? Cadastre-se` -> Conectado (mock local, navega para `/`).

### 2) Navegacao global (Sidebar + Topbar)

- Menu lateral principal (`Conversas`, `Contatos`, `Chatbots`, `Campanhas`, `Relatorios`, `Configuracoes`) -> Conectado (navegacao local).
- Avatar de perfil (abre card) -> Conectado (toggle local).
- Itens do card de perfil (`Minha conta`, `Preferencias`, `Meu perfil`, `Assinaturas e planos`, `Indique e ganhe`, `Organizacao`, `Sair`) -> Nao implementado (sem acao).
- Icones da topbar (buscar/chat/enviar) -> Nao implementado (sem acao).
- Icones de notificacao/IA na sidebar -> Nao implementado (sem acao).

### 3) Conversas (`/`)

- Busca da fila -> Conectado (filtro local).
- Abas `Entrada/Esperando/Finalizados` -> Conectado (filtro por status).
- `Acoes em massa` -> Conectado (selecionar, distribuir, transferir, finalizar no estado local).
- `Ordenar por` -> Conectado (ordenacao local).
- `+` (abrir contato para nova conversa) -> Conectado (modal abre e inicia conversa mock).
- Botao de trial (`Ja quero assinar`) -> Nao implementado.
- Menu `...` da fila com:
- `Atribuir para mim`, `Adicionar etiqueta`, `Marcar como nao lida`, `Bloquear contato`, `Abrir lado-a-lado`, `Finalizar conversa`, `Marcar como esperando`, `Fixar conversa` -> Conectado (estado local).
- Topbar da conversa:
- `Contato`, `Detalhes da conversa`, `Agendar envio`, `Transferir`, `Finalizar` -> Conectado (drawers/acoes locais).
- Menu da mensagem (`Responder`, `Copiar`, `Encaminhar`, `Apagar`) -> Conectado (inclui clipboard quando permitido).
- Composer:
- Toggle `Assinar`
- Tabs `Mensagem/Notas`
- Envio da mensagem/nota
- `Anexar`, `Emoji`, `Figurinhas`, `Respostas rapidas`, `Executar chatbot`, `Agendar`
-> Conectado (estado local).
- Anexos:
- `Novo arquivo` -> Conectado (abre seletor do Windows via `<input type="file">`).
- `Biblioteca de midia` -> Conectado (lista midias da conversa atual).
- `Todos os arquivos` (filtro no anexo) -> Parcial (botao visual sem acao de filtro).

### 4) Contatos (`/contacts`)

- Busca, filtro por etiqueta, ordenacao -> Conectado (estado local).
- `Adicionar contato` -> Conectado (modal + criacao local).
- Menu de linha (`Ver detalhes`, `Desativar/Ativar`, `Alternar etiqueta`) -> Conectado (estado local).
- Painel de detalhes (observacoes editaveis) -> Conectado (estado local).

### 5) Chatbots (`/chatbots`)

- Cards de chatbot renderizados -> Parcial (visual pronto).
- `Editar fluxo` -> Nao implementado (sem acao).
- Fluxo visual estilo builder dos prints -> Nao implementado.

### 6) Campanhas (`/bulksend`)

- Busca e listagem -> Conectado (estado local).
- `Nova campanha` + modal -> Conectado (criacao local).
- Status da campanha (toggle de estado) -> Conectado (estado local).
- Segmentacao/execucao real de disparo -> Fora do escopo atual.

### 7) Relatorios (`/dashboard`)

- Cards metricos + blocos visuais -> Parcial (mock visual).
- Filtros detalhados, periodos, drill-down e exportacoes dos prints -> Nao implementado.

### 8) Configuracoes indice (`/settings`)

- Lista de modulos com navegacao para rotas implementadas -> Parcial.
- Modulos sem rota dedicada ainda:
- `Financeiro`, `Setores`, `Agentes de IA`, `Bases de conhecimento`, `Chatbots` (config), entre outros do inventario
-> Nao implementado.

### 9) Canais (`/settings/channels`)

- Listagem, status, novo canal -> Conectado (estado local).
- Fluxo de conexao real de canal -> Nao implementado.

### 10) Atendentes (`/settings/agents`)

- Busca, mostrar desativados, permissao, reatribuicao, ativar/desativar -> Conectado (estado local).
- `Convidar atendente` -> Conectado (modal + criacao local).
- Convite real por e-mail/perfil -> Nao implementado.

### 11) Etiquetas (`/settings/tags`)

- Listagem, ativar/inativar, criar etiqueta -> Conectado (estado local).

### 12) Respostas rapidas (`/settings/quick-replies`)

- Listagem, ativar/inativar, criar resposta -> Conectado (estado local).

### 13) Templates WhatsApp (`/settings/templates`)

- Canais de template, troca de canal, criar canal, criar template, status -> Conectado (estado local).
- Integracao oficial de template/qualidade/provedor -> Nao implementado.

## Itens dos prints que ainda faltam fidelidade funcional

- Menus de perfil e organizacao com comportamento completo.
- Fluxos avancados de dashboard (filtros, tabs, paineis laterais, graficos fieis).
- Modulos de configuracao ainda nao transformados em paginas completas (`Financeiro`, `Setores`, `Agentes de IA`, `Bases de conhecimento`, etc.).
- Chatbot builder visual (canvas com nos e conexoes).
- Regras de permissao reais por papel/organizacao.
- Persistencia real (backend, banco, auditoria, historico completo).

## Backlog recomendado para a proxima iteracao

1. Fechar a arquitetura da Fase 0 para sair de mock local.
2. Implementar `Conversas` com banco, realtime, filas, chatbot e anexos reais.
3. Implementar `Contatos` com persistencia e impacto cruzado na inbox.
4. Implementar `Atendentes`, `Departamentos` e controle de acesso `Admin`/`Operador`.
5. Ligar `Dashboards` apenas apos os dados reais dos modulos acima.
