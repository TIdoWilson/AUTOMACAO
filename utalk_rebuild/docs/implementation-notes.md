# UTalk Rebuild - Notas de Implementacao

## Estrategia

Este projeto sera refinado em duas camadas:

1. Camada visual
- replica layout
- replica espacamento
- replica hierarquia
- replica componentes e estados

2. Camada funcional
- navegacao real
- abertura de menus e modais
- filtros
- formularios
- persistencia e regras de negocio

No momento, o projeto esta na fase de base visual navegavel.

## Regra de Privacidade

- nao copiar dados reais dos prints para o codigo
- nao colocar miniaturas dos prints dentro de `public/`
- usar apenas placeholders sinteticos como `Operador atual`, `Contato prioritario`, `usuario@empresa.local`
- manter os prints reais somente como referencia externa de implementacao

## Regras de Refinamento

- Cada botao deve primeiro existir na posicao correta.
- Depois cada botao recebe nome, contexto e tela de origem.
- So depois implementamos comportamento.
- Sempre que houver print de um estado aberto, esse estado deve virar um componente separado.

## Componentes Que Precisam Ser Quebrados

O arquivo `src/App.jsx` hoje concentra a prova de conceito. Nos proximos refinamentos, dividir em:

- `components/layout/Sidebar`
- `components/layout/Topbar`
- `components/chat/ConversationList`
- `components/chat/ChatStage`
- `components/chat/ConversationToolbar`
- `components/chat/Composer`
- `components/contacts/ContactsTable`
- `components/settings/SettingsIndex`
- `components/settings/AgentTable`
- `components/settings/TemplateEmptyState`
- `components/dashboard/MetricCards`
- `components/dashboard/Charts`

## Estados Visuais a Modelar

### Conversas
- lista padrao
- item selecionado
- menu aberto
- drawer lateral
- popover de acoes
- area de composer minimizada
- area de composer expandida
- agendamento
- detalhes do contato

### Contatos
- lista padrao
- filtros abertos
- formulario de novo contato
- detalhes do contato

### Configuracoes
- lista indice
- tabela com dropdown aberto
- estado vazio
- modal
- pagina de formulario

### Relatorios
- cards
- tabs
- filtros
- grafico de barras
- grafico de linha
- grafico circular

## Ordem de Entrega Recomendada

1. Conversas
2. Contatos
3. Configuracoes indice
4. Atendentes
5. Etiquetas
6. Respostas rapidas
7. Canais
8. Dashboard
9. Organizacao

## Integracao com WhatsApp

### O que e seguro planejar agora
- arquitetura multiatendente
- filas
- setores
- permissao de operadores
- contatos
- etiquetas
- respostas rapidas
- dashboard
- logs

### O que nao deve virar base do produto
- automacao nao oficial do WhatsApp Web
- scraping do cliente web
- dependencia de conexao via navegador controlado
- fluxo de QR code baseado em cliente nao suportado

### Direcao recomendada

Construir o produto primeiro como plataforma de atendimento multicanal e deixar a camada de integracao desacoplada.
Assim, depois podemos plugar um conector suportado sem reescrever o sistema inteiro.
