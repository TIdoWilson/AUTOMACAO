# UTalk Rebuild - Fluxos da Inbox

Documento de apoio para a etapa atual da tela de conversas.

## Objetivo

Descrever o comportamento sintetico dos principais botoes da inbox antes de qualquer backend real.

## Regras desta simulacao

- todos os dados permanecem sinteticos
- nenhuma acao conversa com API externa
- cada interacao apenas atualiza estado local da interface

## Fluxos ja ligados no frontend

### Lista e filtros

- abas `Entrada`, `Esperando` e `Finalizados` filtram a lista local
- busca filtra por nome e resumo da conversa
- `Ordenar por` reorganiza a lista por criacao, ultima mensagem ou prioridade de resposta

### Acoes em massa

- `Selecionar todas as conversas`
- `Distribuir para um atendente`
- `Transferir para outro setor`
- `Finalizar selecionadas`

Essas acoes alteram status, fila e timeline da conversa.

### Fluxo da conversa individual

- `Contato` abre dados gerais e permite salvar observacoes sinteticas
- `Detalhes da conversa` mostra status e timeline operacional
- `Transferir conversa` muda a fila/setor
- `Finalizar conversa` move a conversa para a aba de encerradas
- `Agendar envio` salva data e hora simuladas na conversa
- menu de mensagem com `Responder`, `Copiar`, `Encaminhar` e `Apagar`
- menu de fila nos `3 pontinhos` com atribuicao, etiqueta, nao lida, bloqueio, espera, fixacao e finalizacao

### Composer

- `Mensagem` envia uma nova bolha para a thread local
- `Notas` grava um evento interno na timeline
- `Respostas rapidas` carrega um texto pronto no composer
- `Executar chatbot` registra um fluxo automatizado na timeline
- `Anexar`, `Emojis` e `Figurinhas` inserem marcadores sinteticos no composer
- `Novo arquivo` abre o seletor local de arquivos do Windows
- `Biblioteca de midia` lista os arquivos associados a conversa atual

## Proxima camada recomendada

1. Separar a inbox em componentes menores.
2. Criar mocks por modulo para timeline, contato, fila e agendamento.
3. Definir contrato de dados da conversa antes de backend real.
4. Modelar permissao por atendente, setor e ownership da conversa.
