# UTalk Rebuild - Mapa da Tela de Conversas

Politica de privacidade desta documentacao:
- nenhuma pessoa, telefone, e-mail ou identificador real dos prints foi copiado
- os nomes abaixo representam apenas papeis visuais e estados da interface
- qualquer exemplo funcional deve continuar sintetico

## Objetivo desta etapa

Mapear a inbox principal antes de aprofundar regras de negocio. A ideia e:
- posicionar botoes e estados visuais conforme os prints
- garantir que cada menu, drawer e modal tenha um equivalente navegavel
- documentar o que cada controle aparenta fazer

## Estrutura principal

### Coluna esquerda

Elementos fixos identificados:
- titulo `Conversas`
- busca por `nome ou telefone`
- abas `Entrada`, `Esperando`, `Finalizados`
- banner de trial
- lista de cards de conversa com avatar, horario, categoria e quantidade de nao lidas

Controles mapeados:
- botao de filtro/acoes em massa
- botao para abrir selecao de contato
- seletor `Acoes em massa`
- seletor `Ordenar por`

Menus associados:
- menu de acoes em massa
  - selecionar todas as conversas
  - distribuir para um atendente
  - transferir para outro setor
  - finalizar selecionadas
- menu de ordenacao
  - data de criacao
  - ultima mensagem
  - esperando resposta

### Centro da tela

Elementos fixos identificados:
- topbar da conversa com avatar, nome e status sintetico
- area de mensagens com marcacao temporal
- bolhas recebidas, enviadas e de sistema
- composer com alternancia entre `Mensagem` e `Notas`

Controles da topbar:
- `Contato`
- `Detalhes da conversa`
- `Agendar envio`
- transferir conversa
- finalizar conversa

Controles do composer:
- anexos
- emojis
- figurinhas
- respostas rapidas
- executar chatbot
- agendar mensagem
- envio da mensagem ou nota

## Overlays e paines confirmados pelos prints

### Modal `Escolha um contato`

Uso visual:
- iniciar uma nova conversa
- pesquisar um contato sintetico
- abrir um perfil da lista

### Painel lateral `Contato`

Uso visual:
- resumir dados do contato
- mostrar etiquetas e origem
- exibir campos editaveis sinteticos

### Painel lateral `Detalhes da conversa`

Uso visual:
- status atual
- canal
- horario de abertura e ultima atividade
- timeline operacional

### Painel lateral `Respostas rapidas`

Uso visual:
- pesquisar atalhos
- exibir grupos de respostas
- disponibilizar snippets para inserir no composer

### Painel lateral `Executar chatbot`

Uso visual:
- listar fluxos disponiveis
- iniciar uma automacao na conversa atual

### Painel lateral `Agendamento de mensagens`

Uso visual:
- revisar mensagem agendada
- escolher data e hora
- sugerir horarios rapidos

### Popover `Anexar`

Uso visual:
- arquivo do computador
- biblioteca de midias
- documento/modelo

### Popover `Reacoes rapidas`

Uso visual:
- grade compacta de reacoes
- insercao rapida no composer

### Drawer inferior `Figurinhas salvas`

Uso visual:
- grid de itens salvos
- selecao rapida durante o atendimento

## Sequencia recomendada para as proximas iteracoes

1. Refinar o layout da lista de conversas e seus estados hover/selecionado.
2. Ajustar a topbar da conversa com mais fidelidade de espacamento.
3. Especificar a funcionalidade de cada painel lateral.
4. Descrever o fluxo operacional da inbox: receber, assumir, transferir, agendar, finalizar.
5. Ligar os componentes a dados mockados por modulo antes de pensar em backend.
