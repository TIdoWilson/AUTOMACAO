# UTalk Rebuild - Fluxos de Campanhas e Templates

## Objetivo

Registrar o que ja esta funcional nessas areas para que a evolucao do projeto continue consistente.

## Templates WhatsApp Business API

Fluxos ligados no frontend:
- listagem de canais internos
- alternancia entre canais
- criacao de novo canal
- criacao de novo template
- alternancia de status do template entre rascunho, revisao e aprovado

## Campanhas

Fluxos ligados no frontend:
- busca por campanha
- criacao de nova campanha
- listagem com publico, template e canal
- alternancia de status entre rascunho, programada e concluida

## Proximo passo recomendado

1. Refinar `Configuracoes` para abrir modulos com estado proprio.
2. Comecar a quebrar `App.jsx` em componentes reutilizaveis.
3. Preparar uma camada de dados mockada por modulo para facilitar a futura migracao para backend real.
