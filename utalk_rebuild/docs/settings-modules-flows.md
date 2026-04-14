# UTalk Rebuild - Modulos de Configuracoes

## Objetivo

Registrar os modulos de configuracao que deixaram de ser apenas indice visual e passaram a ter comportamento proprio.

## Canais de atendimento

Fluxos ligados no frontend:
- listagem de canais
- alternancia visual de status
- criacao de novo canal sintetico

## Etiquetas

Fluxos ligados no frontend:
- listagem de etiquetas em cards
- alternancia entre ativa e inativa
- criacao de nova etiqueta com nome e cor

## Respostas rapidas

Fluxos ligados no frontend:
- listagem de atalhos
- alternancia entre ativa e inativa
- criacao de nova resposta rapida com atalho, titulo, mensagem e escopo

## Proximo passo recomendado

1. Quebrar `App.jsx` em componentes menores por dominio.
2. Criar uma camada de mock data centralizada para evitar estado espalhado.
3. Evoluir `Dashboard` e `Chatbots` com interacoes reais do frontend.
