# UTalk Rebuild - Fluxos de Contatos e Atendentes

## Objetivo

Registrar o que ja esta funcional nessas duas areas para orientar as proximas iteracoes.

## Contatos

Fluxos ligados no frontend:
- busca por nome, telefone e e-mail sintetico
- filtro por etiqueta
- ordenacao por nome ou ultima interacao
- selecao de linhas
- menu por contato com visualizacao, ativacao/desativacao e alternancia de etiqueta
- painel de detalhes do contato com observacoes editaveis
- modal de criacao de contato sintetico

## Atendentes

Fluxos ligados no frontend:
- busca por nome e e-mail sintetico
- alternancia para mostrar desativados
- alteracao de permissao por dropdown
- alteracao de reatribuicao por dropdown
- ativacao e desativacao por acao de linha
- modal de convite/cadastro de atendente sintetico

## Proximos passos naturais

1. Dar o mesmo nivel de funcionalidade para `Templates` e `Campanhas`.
2. Transformar `Configuracoes` em hub de modulos com estados vazios e formularios.
3. Separar o arquivo `App.jsx` em componentes menores antes que a manutencao fique pesada.
