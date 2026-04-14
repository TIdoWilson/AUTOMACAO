export const navItems = [
  { key: 'conversations', label: 'Conversas', path: '/', icon: 'chat' },
  { key: 'contacts', label: 'Contatos', path: '/contacts', icon: 'user' },
  { key: 'chatbots', label: 'Chatbots', path: '/chatbots', icon: 'bot' },
  { key: 'bulksend', label: 'Campanhas', path: '/bulksend', icon: 'send' },
  { key: 'dashboard', label: 'Relatorios', path: '/dashboard', icon: 'chart' },
  { key: 'settings', label: 'Configuracoes', path: '/settings', icon: 'gear' },
]

export const conversationItems = [
  { id: 1, name: 'Contato prioritario', preview: 'Mensagem de teste recebida.', time: '08:47', unread: 1, category: 'Geral', color: '#ffcb75' },
  { id: 2, name: 'Contato com numero oculto', preview: 'Obrigado pelo retorno.', time: '08:43', unread: 0, category: 'Geral', color: '#3da3ff' },
  { id: 3, name: 'Cliente recorrente', preview: 'Preciso de ajuda com o atendimento.', time: '08:41', unread: 2, category: 'Geral', color: '#8baf5a' },
  { id: 4, name: 'Lead recente', preview: 'Aguardando aprovacao.', time: '08:41', unread: 0, category: 'Geral', color: '#dc8c7c' },
  { id: 5, name: 'Umbler Chatbot', preview: 'Serio, responde ai alguma coisa...', time: '08:38', unread: 0, category: '', color: '#4c78ff' },
]

export const contactItems = [
  { id: 1, name: 'Contato sem identificacao', phone: 'Telefone oculto', note: '', color: '#d48377' },
  { id: 2, name: 'Contato importado', phone: 'Telefone oculto', note: '', color: '#47a6ff' },
  { id: 3, name: 'Cliente principal', phone: 'Telefone oculto', note: 'ha 9 minutos', color: '#ffd577' },
  { id: 4, name: 'Cliente secundario', phone: 'Telefone oculto', note: 'ha 15 minutos', color: '#8bbd68' },
  { id: 5, name: 'Chatbot interno', phone: 'Canal automatizado', note: '', color: '#5b80ff' },
]

export const settingsGroups = [
  ['Canais de atendimento', 'Os canais que sua organizacao usa para se comunicar com os seus contatos'],
  ['Financeiro', 'Veja todas as suas informacoes de pagamento incluindo historico e cartoes registrados.'],
  ['Setores', 'Configuracoes das sub-divisoes da sua organizacao'],
  ['Agentes de IA', 'Configure e treine os seus agentes de IA'],
  ['Bases de conhecimento', 'Importe conteudos que a IA usara como fonte para responder seus clientes'],
  ['Atendentes', 'As pessoas que lhe ajudam com o relacionamento com seus contatos'],
  ['Etiquetas', 'Configuracoes das etiquetas da sua organizacao'],
  ['Chatbots', 'Robos para automatizar atendimentos'],
  ['Respostas rapidas', 'Mensagens pre-configuradas para enviar para seus contatos'],
  ['Templates WhatsApp Business API', 'Mensagens para enviar a contatos que nao estao ativos no WhatsApp Business API'],
]

export const settingRoutes = [
  { title: 'Canais de atendimento', path: '/settings/channels' },
  { title: 'Atendentes', path: '/settings/agents' },
  { title: 'Etiquetas', path: '/settings/tags' },
  { title: 'Respostas rapidas', path: '/settings/quick-replies' },
  { title: 'Templates WhatsApp Business API', path: '/settings/templates' },
]

export const dashboardCards = [
  { title: 'Conversas atendidas', value: '248', delta: '+14%' },
  { title: 'Tempo medio de resposta', value: '02:31', delta: '-9%' },
  { title: 'Contatos ativos', value: '1.284', delta: '+6%' },
  { title: 'Campanhas concluidas', value: '18', delta: '+2' },
]

export const chatbotCards = [
  { title: 'Boas-vindas automatizadas', subtitle: 'Ativo em 3 canais' },
  { title: 'Triagem comercial', subtitle: 'Leads qualificados por setor' },
  { title: 'Pos-venda', subtitle: 'Coleta NPS e reabertura de tickets' },
]

export const campaignRows = [
  ['Reativacao de leads', '2.480 contatos', 'Programada'],
  ['Cobertura de feriado', '612 contatos', 'Rascunho'],
  ['Aviso de manutencao', '1.145 contatos', 'Concluida'],
]

export const conversationMessageThread = [
  { id: 1, side: 'left', body: 'Ola, preciso confirmar uma atualizacao cadastral.', time: '08:41' },
  { id: 2, side: 'left', body: 'Tambem gostaria de validar o historico do ultimo atendimento.', time: '08:42' },
  { id: 3, side: 'right', author: 'Operador atual', body: 'Perfeito. Estou abrindo a conversa e revisando os detalhes agora.', time: '08:44' },
  { id: 4, side: 'system', body: 'Contato movido para a fila geral por uma regra automatica.', time: '08:45' },
  { id: 5, side: 'left', body: 'Obrigado. Fico no aguardo do retorno.', time: '08:46' },
]

export const bulkActionItems = [
  'Selecionar todas as conversas',
  'Distribuir para um atendente',
  'Transferir para outro setor',
  'Finalizar selecionadas',
]

export const sortActionItems = [
  'Data de criacao',
  'Ultima mensagem',
  'Esperando resposta',
]

export const quickReplyGroups = [
  {
    title: 'Saudacao',
    items: [
      { shortcut: '/inicio', title: 'Boas-vindas padrao', body: 'Ola! Recebi sua mensagem e vou seguir com o atendimento.' },
      { shortcut: '/fila', title: 'Fila de espera', body: 'Seu contato foi registrado e um operador respondera em instantes.' },
    ],
  },
  {
    title: 'Operacao',
    items: [
      { shortcut: '/status', title: 'Atualizacao interna', body: 'Estou revisando os dados e retornarei com a confirmacao.' },
      { shortcut: '/encerrar', title: 'Encerramento', body: 'Se precisar de algo mais, estou a disposicao.' },
    ],
  },
]

export const chatbotPlaybooks = [
  { title: 'Triagem inicial', description: 'Encaminha o contato conforme o motivo do atendimento.' },
  { title: 'Pos-venda', description: 'Solicita confirmacao e abre fluxo de acompanhamento.' },
  { title: 'Reengajamento', description: 'Retoma conversas pausadas com uma mensagem automatica.' },
]

export const schedulerBlocks = [
  { title: 'Hoje', options: ['10:30', '14:00', '16:15'] },
  { title: 'Amanha', options: ['09:00', '11:45', '15:30'] },
]

export const stickerItems = ['Atendimento', 'Confirmado', 'Processando', 'Ok', 'Equipe', 'Lembrete', 'Retorno', 'Aprovado']

export const contactInsightItems = [
  ['Canal principal', 'WhatsApp Web conectado'],
  ['Origem', 'Fila geral'],
  ['Ultimo operador', 'Operador atual'],
  ['Setor sugerido', 'Suporte'],
]

export const conversationDetailItems = [
  ['Status', 'Aguardando resposta da equipe'],
  ['Canal', 'WhatsApp'],
  ['Primeira mensagem', 'Hoje, 08:41'],
  ['Ultima atividade', 'Hoje, 08:46'],
  ['Etiqueta ativa', 'Prioridade media'],
]

export const tabDefinitions = [
  ['entrada', 'Entrada'],
  ['esperando', 'Esperando'],
  ['finalizados', 'Finalizados'],
]

export const operatorOptions = ['Operador atual', 'Fila de suporte', 'Equipe comercial']
export const sectorOptions = ['Geral', 'Suporte', 'Comercial', 'Financeiro']
export const contactTagOptions = ['Todos', 'Vip', 'Lead', 'Suporte', 'Financeiro']
export const agentPermissionOptions = ['Membro', 'Operador', 'Admin', 'Proprietario']
export const agentReassignmentOptions = ['Desligada', 'Ligada']
export const queueActionItems = [
  'Atribuir para mim',
  'Adicionar etiqueta',
  'Marcar como nao lida',
  'Bloquear contato',
  'Abrir lado-a-lado',
  'Finalizar conversa',
  'Marcar como esperando',
  'Fixar conversa',
]

export function buildConversationState() {
  return conversationItems.map((item, index) => ({
    ...item,
    status: index === 2 ? 'esperando' : 'entrada',
    assignedTo: index === 4 ? 'Fluxo automatizado' : operatorOptions[0],
    priority: index === 2 ? 'alta' : index === 0 ? 'media' : 'normal',
    selected: false,
    note: 'Perfil utilizado apenas para validacao visual da tela.',
    scheduledFor: null,
    sector: item.category || 'Geral',
    blocked: false,
    pinned: false,
    splitView: false,
  }))
}

export function buildThreadState() {
  return Object.fromEntries(
    conversationItems.map((item, index) => [
      item.id,
      conversationMessageThread.map((message, messageIndex) => ({
        ...message,
        id: `${item.id}-${message.id}`,
        body:
          index === 0 || messageIndex !== 0
            ? message.body
            : `Ola, preciso confirmar uma atualizacao relacionada a ${item.name.toLowerCase()}.`,
      })),
    ]),
  )
}

export function buildTimelineState() {
  return Object.fromEntries(
    conversationItems.map((item) => [
      item.id,
      [
        { time: '08:41', text: `${item.name} entrou na fila geral.` },
        { time: '08:44', text: 'Atendimento assumido por Operador atual.' },
        { time: '08:45', text: 'Regra automatica aplicou etiqueta de prioridade.' },
      ],
    ]),
  )
}

export function buildMediaLibraryState() {
  return {
    1: [
      { id: 1, name: 'boas-vindas.png', type: 'Imagem', origin: 'Conversa', size: '240 KB' },
      { id: 2, name: 'resumo-atendimento.pdf', type: 'Documento', origin: 'Conversa', size: '120 KB' },
    ],
    2: [{ id: 3, name: 'print-suporte.jpg', type: 'Imagem', origin: 'Conversa', size: '310 KB' }],
    3: [],
    4: [],
    5: [{ id: 4, name: 'roteiro-chatbot.txt', type: 'Documento', origin: 'Conversa', size: '18 KB' }],
  }
}

export function nowTimeLabel() {
  const now = new Date()
  return now.toLocaleTimeString('pt-BR', { hour: '2-digit', minute: '2-digit' })
}

export function buildContactState() {
  return contactItems.map((item, index) => ({
    ...item,
    selected: false,
    email: `contato${index + 1}@empresa.local`,
    tag: index === 0 ? 'Vip' : index === 1 ? 'Lead' : index === 4 ? 'Suporte' : 'Financeiro',
    origin: index === 0 ? 'Importacao manual' : index === 1 ? 'Formulario interno' : 'Atendimento',
    company: index === 2 ? 'Conta principal' : 'Operacao interna',
    status: index === 3 ? 'Inativo' : 'Ativo',
    notes: 'Registro sintetico usado apenas para modelagem visual do sistema.',
  }))
}

export function buildAgentState() {
  return [
    { id: 1, name: 'Operador atual', email: 'usuario@empresa.local', permission: 'Proprietario', reassignment: 'Desligada', selected: false, isSelf: true, active: true, color: '#53c86f' },
    { id: 2, name: 'Coordenacao interna', email: 'coordenacao@empresa.local', permission: 'Admin', reassignment: 'Ligada', selected: false, isSelf: false, active: true, color: '#5f7df5' },
    { id: 3, name: 'Fila de apoio', email: 'apoio@empresa.local', permission: 'Operador', reassignment: 'Desligada', selected: false, isSelf: false, active: false, color: '#dc8c7c' },
  ]
}

export function buildTemplateState() {
  return [
    {
      id: 1,
      channelName: 'Canal principal',
      phone: 'Numero verificado',
      provider: 'Conexao interna',
      quality: 'Boa',
      templates: [
        { id: 11, title: 'Boas-vindas operacionais', category: 'Marketing', status: 'Aprovado', language: 'pt-BR' },
        { id: 12, title: 'Retorno de acompanhamento', category: 'Utilitario', status: 'Em revisao', language: 'pt-BR' },
      ],
    },
  ]
}

export function buildCampaignState() {
  return campaignRows.map((row, index) => ({
    id: index + 1,
    title: row[0],
    audience: row[1],
    status: row[2],
    channel: index === 0 ? 'WhatsApp' : index === 1 ? 'E-mail interno' : 'Aviso interno',
    template: index === 0 ? 'Boas-vindas operacionais' : index === 1 ? 'Retorno de acompanhamento' : 'Comunicado geral',
    selected: false,
  }))
}

export function buildChannelSettingsState() {
  return [
    { id: 1, name: 'WhatsApp principal', type: 'WhatsApp', status: 'Conectado', owner: 'Operador atual', selected: false },
    { id: 2, name: 'Canal comercial', type: 'Instagram', status: 'Em configuracao', owner: 'Equipe comercial', selected: false },
  ]
}

export function buildTagState() {
  return [
    { id: 1, name: 'Vip', color: '#5f7df5', usage: 18, active: true },
    { id: 2, name: 'Lead', color: '#53c86f', usage: 42, active: true },
    { id: 3, name: 'Financeiro', color: '#ef9b53', usage: 9, active: false },
  ]
}

export function buildQuickReplyState() {
  return [
    { id: 1, shortcut: '/inicio', title: 'Boas-vindas', body: 'Ola! Recebi sua mensagem e vou seguir com o atendimento.', scope: 'Geral', active: true },
    { id: 2, shortcut: '/status', title: 'Atualizacao interna', body: 'Estou revisando os dados e retorno em instantes.', scope: 'Suporte', active: true },
    { id: 3, shortcut: '/fechar', title: 'Encerramento', body: 'Se precisar de algo mais, sigo a disposicao.', scope: 'Comercial', active: false },
  ]
}
