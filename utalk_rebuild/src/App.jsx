import { useEffect, useRef, useState } from 'react'
import { Sidebar } from './components/layout/Sidebar'
import { Topbar } from './components/layout/Topbar'
import { Glyph } from './components/shared/Glyph'
import { usePathname, navigate } from './lib/navigation'
import './App.css'

const navItems = [
  { key: 'conversations', label: 'Conversas', path: '/', icon: 'chat' },
  { key: 'contacts', label: 'Contatos', path: '/contacts', icon: 'user' },
  { key: 'chatbots', label: 'Chatbots', path: '/chatbots', icon: 'bot' },
  { key: 'bulksend', label: 'Campanhas', path: '/bulksend', icon: 'send' },
  { key: 'dashboard', label: 'Relatorios', path: '/dashboard', icon: 'chart' },
  { key: 'settings', label: 'Configuracoes', path: '/settings', icon: 'gear' },
]

const conversationItems = [
  {
    id: 1,
    name: 'Contato prioritario',
    preview: 'Mensagem de teste recebida.',
    time: '08:47',
    unread: 1,
    category: 'Geral',
    color: '#ffcb75',
  },
  {
    id: 2,
    name: 'Contato com numero oculto',
    preview: 'Obrigado pelo retorno.',
    time: '08:43',
    unread: 0,
    category: 'Geral',
    color: '#3da3ff',
  },
  {
    id: 3,
    name: 'Cliente recorrente',
    preview: 'Preciso de ajuda com o atendimento.',
    time: '08:41',
    unread: 2,
    category: 'Geral',
    color: '#8baf5a',
  },
  {
    id: 4,
    name: 'Lead recente',
    preview: 'Aguardando aprovacao.',
    time: '08:41',
    unread: 0,
    category: 'Geral',
    color: '#dc8c7c',
  },
  {
    id: 5,
    name: 'Umbler Chatbot',
    preview: 'Serio, responde ai alguma coisa...',
    time: '08:38',
    unread: 0,
    category: '',
    color: '#4c78ff',
  },
]

const contactItems = [
  { id: 1, name: 'Contato sem identificacao', phone: 'Telefone oculto', note: '', color: '#d48377' },
  { id: 2, name: 'Contato importado', phone: 'Telefone oculto', note: '', color: '#47a6ff' },
  { id: 3, name: 'Cliente principal', phone: 'Telefone oculto', note: 'ha 9 minutos', color: '#ffd577' },
  { id: 4, name: 'Cliente secundario', phone: 'Telefone oculto', note: 'ha 15 minutos', color: '#8bbd68' },
  { id: 5, name: 'Chatbot interno', phone: 'Canal automatizado', note: '', color: '#5b80ff' },
]

const settingsGroups = [
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

const settingRoutes = [
  { title: 'Canais de atendimento', path: '/settings/channels' },
  { title: 'Atendentes', path: '/settings/agents' },
  { title: 'Etiquetas', path: '/settings/tags' },
  { title: 'Respostas rapidas', path: '/settings/quick-replies' },
  { title: 'Templates WhatsApp Business API', path: '/settings/templates' },
]

const dashboardCards = [
  { title: 'Conversas atendidas', value: '248', delta: '+14%' },
  { title: 'Tempo medio de resposta', value: '02:31', delta: '-9%' },
  { title: 'Contatos ativos', value: '1.284', delta: '+6%' },
  { title: 'Campanhas concluídas', value: '18', delta: '+2' },
]

const chatbotCards = [
  { title: 'Boas-vindas automatizadas', subtitle: 'Ativo em 3 canais' },
  { title: 'Triagem comercial', subtitle: 'Leads qualificados por setor' },
  { title: 'Pos-venda', subtitle: 'Coleta NPS e reabertura de tickets' },
]

const campaignRows = [
  ['Reativacao de leads', '2.480 contatos', 'Programada'],
  ['Cobertura de feriado', '612 contatos', 'Rascunho'],
  ['Aviso de manutencao', '1.145 contatos', 'Concluida'],
]

const conversationMessageThread = [
  { id: 1, side: 'left', body: 'Ola, preciso confirmar uma atualizacao cadastral.', time: '08:41' },
  { id: 2, side: 'left', body: 'Tambem gostaria de validar o historico do ultimo atendimento.', time: '08:42' },
  { id: 3, side: 'right', author: 'Operador atual', body: 'Perfeito. Estou abrindo a conversa e revisando os detalhes agora.', time: '08:44' },
  { id: 4, side: 'system', body: 'Contato movido para a fila geral por uma regra automatica.', time: '08:45' },
  { id: 5, side: 'left', body: 'Obrigado. Fico no aguardo do retorno.', time: '08:46' },
]

const bulkActionItems = [
  'Selecionar todas as conversas',
  'Distribuir para um atendente',
  'Transferir para outro setor',
  'Finalizar selecionadas',
]

const sortActionItems = [
  'Data de criacao',
  'Ultima mensagem',
  'Esperando resposta',
]

const quickReplyGroups = [
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

const chatbotPlaybooks = [
  { title: 'Triagem inicial', description: 'Encaminha o contato conforme o motivo do atendimento.' },
  { title: 'Pos-venda', description: 'Solicita confirmacao e abre fluxo de acompanhamento.' },
  { title: 'Reengajamento', description: 'Retoma conversas pausadas com uma mensagem automatica.' },
]

const schedulerBlocks = [
  { title: 'Hoje', options: ['10:30', '14:00', '16:15'] },
  { title: 'Amanha', options: ['09:00', '11:45', '15:30'] },
]

const stickerItems = [
  'Atendimento',
  'Confirmado',
  'Processando',
  'Ok',
  'Equipe',
  'Lembrete',
  'Retorno',
  'Aprovado',
]

const contactInsightItems = [
  ['Canal principal', 'WhatsApp Web conectado'],
  ['Origem', 'Fila geral'],
  ['Ultimo operador', 'Operador atual'],
  ['Setor sugerido', 'Suporte'],
]

const conversationDetailItems = [
  ['Status', 'Aguardando resposta da equipe'],
  ['Canal', 'WhatsApp'],
  ['Primeira mensagem', 'Hoje, 08:41'],
  ['Ultima atividade', 'Hoje, 08:46'],
  ['Etiqueta ativa', 'Prioridade media'],
]

const tabDefinitions = [
  ['entrada', 'Entrada'],
  ['esperando', 'Esperando'],
  ['finalizados', 'Finalizados'],
]

const operatorOptions = ['Operador atual', 'Fila de suporte', 'Equipe comercial']
const sectorOptions = ['Geral', 'Suporte', 'Comercial', 'Financeiro']
const contactTagOptions = ['Todos', 'Vip', 'Lead', 'Suporte', 'Financeiro']
const agentPermissionOptions = ['Membro', 'Operador', 'Admin', 'Proprietario']
const agentReassignmentOptions = ['Desligada', 'Ligada']
const queueActionItems = [
  'Atribuir para mim',
  'Adicionar etiqueta',
  'Marcar como nao lida',
  'Bloquear contato',
  'Abrir lado-a-lado',
  'Finalizar conversa',
  'Marcar como esperando',
  'Fixar conversa',
]

function buildConversationState() {
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

function buildThreadState() {
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

function buildTimelineState() {
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

function buildMediaLibraryState() {
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

function nowTimeLabel() {
  const now = new Date()
  return now.toLocaleTimeString('pt-BR', { hour: '2-digit', minute: '2-digit' })
}

function buildContactState() {
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

function buildAgentState() {
  return [
    {
      id: 1,
      name: 'Operador atual',
      email: 'usuario@empresa.local',
      permission: 'Proprietario',
      reassignment: 'Desligada',
      selected: false,
      isSelf: true,
      active: true,
      color: '#53c86f',
    },
    {
      id: 2,
      name: 'Coordenacao interna',
      email: 'coordenacao@empresa.local',
      permission: 'Admin',
      reassignment: 'Ligada',
      selected: false,
      isSelf: false,
      active: true,
      color: '#5f7df5',
    },
    {
      id: 3,
      name: 'Fila de apoio',
      email: 'apoio@empresa.local',
      permission: 'Operador',
      reassignment: 'Desligada',
      selected: false,
      isSelf: false,
      active: false,
      color: '#dc8c7c',
    },
  ]
}

function buildTemplateState() {
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

function buildCampaignState() {
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

function buildChannelSettingsState() {
  return [
    { id: 1, name: 'WhatsApp principal', type: 'WhatsApp', status: 'Conectado', owner: 'Operador atual', selected: false },
    { id: 2, name: 'Canal comercial', type: 'Instagram', status: 'Em configuracao', owner: 'Equipe comercial', selected: false },
  ]
}

function buildTagState() {
  return [
    { id: 1, name: 'Vip', color: '#5f7df5', usage: 18, active: true },
    { id: 2, name: 'Lead', color: '#53c86f', usage: 42, active: true },
    { id: 3, name: 'Financeiro', color: '#ef9b53', usage: 9, active: false },
  ]
}

function buildQuickReplyState() {
  return [
    { id: 1, shortcut: '/inicio', title: 'Boas-vindas', body: 'Ola! Recebi sua mensagem e vou seguir com o atendimento.', scope: 'Geral', active: true },
    { id: 2, shortcut: '/status', title: 'Atualizacao interna', body: 'Estou revisando os dados e retorno em instantes.', scope: 'Suporte', active: true },
    { id: 3, shortcut: '/fechar', title: 'Encerramento', body: 'Se precisar de algo mais, sigo a disposicao.', scope: 'Comercial', active: false },
  ]
}

function App() {
  const pathname = usePathname()

  useEffect(() => {
    const titles = {
      '/login': 'Entre no Umbler Talk',
      '/': 'Conversas',
      '/contacts': 'Contatos',
      '/chatbots': 'Chatbots',
      '/bulksend': 'Campanhas',
      '/dashboard': 'Relatorios',
      '/settings': 'Configuracoes',
      '/settings/channels': 'Canais de atendimento',
      '/settings/agents': 'Atendentes',
      '/settings/tags': 'Etiquetas',
      '/settings/quick-replies': 'Respostas rapidas',
      '/settings/templates': 'Templates WhatsApp Business API',
    }
    document.title = `${titles[pathname] ?? 'Umbler Talk'} | Rebuild`
  }, [pathname])

  if (pathname === '/login') {
    return <LoginPage />
  }

  return <WorkspacePage pathname={pathname} />
}

function LoginPage() {
  return (
    <div className="login-shell">
      <section className="login-panel">
        <div className="brand-row">
          <img src="/assets/favicon.svg" alt="" className="brand-favicon" />
          <span className="brand-wordmark">umbler talk</span>
        </div>

        <div className="login-copy">
          <h1>Faca login para fazer parte da organizacao</h1>
        </div>

        <div className="social-row">
          <button className="social-card">
            <span className="mini-avatar">M</span>
            <span className="social-meta">
              <strong>Fazer login com a conta da equipe</strong>
              <small>usuario@empresa.local</small>
            </span>
            <span className="social-badge google">G</span>
          </button>
          <button className="social-facebook">Entrar com o Facebook</button>
        </div>

        <div className="divider">ou</div>

        <label className="form-field">
          <span>E-mail</span>
          <input type="text" placeholder="" />
        </label>

        <label className="form-field">
          <span>Senha</span>
          <input type="password" placeholder="" />
        </label>

        <div className="login-actions">
          <label className="checkbox-line">
            <input type="checkbox" />
            <span>Manter-me conectado</span>
          </label>
          <a href="/">Esqueci minha senha</a>
        </div>

        <button className="primary-button">Entrar</button>
        <button className="ghost-link" onClick={() => navigate('/')}>Ainda nao possui conta? Cadastre-se</button>
      </section>

      <section className="promo-panel">
        <div className="promo-hero">
          <div className="promo-copy">
            <h2>Todo seu time oferecendo suporte em um so WhatsApp</h2>
            <p>Como o Umbler Talk, voce coloca quantos operadores quiser atendendo simultaneamente em varios computadores usando um unico numero.</p>
          </div>
          <div className="promo-mockup">
            <div className="promo-sidebar" />
            <div className="promo-card">
              <div className="promo-search" />
              <div className="promo-tabs">
                <span>Entrada</span>
                <span>Esperando</span>
                <span>Finalizados</span>
              </div>
              <div className="promo-list">
                <div className="promo-list-item large" />
                <div className="promo-list-item" />
                <div className="promo-list-item" />
              </div>
            </div>
            <div className="promo-floating-card">
              <div className="promo-floating-header">
                <span>Etiquetas</span>
                <button>Criar etiqueta</button>
              </div>
              <div className="promo-floating-input" />
              <div className="promo-floating-chip">orcamento</div>
            </div>
          </div>
        </div>
      </section>
    </div>
  )
}

function WorkspacePage({ pathname }) {
  const isConversationRoute = pathname === '/' || pathname.startsWith('/chats')
  const routeLabelMap = {
    '/settings/channels': 'Canais de atendimento',
    '/settings/agents': 'Atendentes',
    '/settings/tags': 'Etiquetas',
    '/settings/quick-replies': 'Respostas rapidas',
    '/settings/templates': 'Templates WhatsApp Business API',
  }
  const currentLabel = routeLabelMap[pathname] ?? navItems.find((item) => item.path === pathname)?.label ?? 'Configuracoes'

  return (
    <div className={`workspace-shell ${isConversationRoute ? 'is-chat' : ''}`}>
      <Sidebar pathname={pathname} />
      <main className="workspace-main">
        <Topbar currentLabel={currentLabel} pathname={pathname} />
        <div className="workspace-body">
          {pathname === '/' && <ConversationsPage />}
          {pathname === '/contacts' && <ContactsPage />}
          {pathname === '/settings' && <SettingsPage />}
          {pathname === '/settings/channels' && <ChannelsPage />}
          {pathname === '/settings/agents' && <AgentsPage />}
          {pathname === '/settings/tags' && <TagsPage />}
          {pathname === '/settings/quick-replies' && <QuickRepliesPage />}
          {pathname === '/settings/templates' && <TemplatesPage />}
          {pathname === '/dashboard' && <DashboardPage />}
          {pathname === '/chatbots' && <ChatbotsPage />}
          {pathname === '/bulksend' && <CampaignsPage />}
          {!['/', '/contacts', '/settings', '/settings/channels', '/settings/agents', '/settings/tags', '/settings/quick-replies', '/settings/templates', '/dashboard', '/chatbots', '/bulksend'].includes(pathname) && (
            <SettingsPage />
          )}
        </div>
      </main>
    </div>
  )
}

function ConversationsPage() {
  const [conversations, setConversations] = useState(buildConversationState)
  const [selectedConversationId, setSelectedConversationId] = useState(conversationItems[0].id)
  const [listPopover, setListPopover] = useState(null)
  const [activePanel, setActivePanel] = useState(null)
  const [showAttachments, setShowAttachments] = useState(false)
  const [showEmojiPicker, setShowEmojiPicker] = useState(false)
  const [showStickerDrawer, setShowStickerDrawer] = useState(false)
  const [showContactPicker, setShowContactPicker] = useState(false)
  const [composerMode, setComposerMode] = useState('message')
  const [activeTab, setActiveTab] = useState('entrada')
  const [sortMode, setSortMode] = useState(sortActionItems[0])
  const [searchTerm, setSearchTerm] = useState('')
  const [composerText, setComposerText] = useState('')
  const [contactNoteDraft, setContactNoteDraft] = useState('Perfil utilizado apenas para validacao visual da tela.')
  const [transferTarget, setTransferTarget] = useState(sectorOptions[1])
  const [scheduleDraft, setScheduleDraft] = useState({ date: '15/04/2026', time: '09:00' })
  const [messageThreads, setMessageThreads] = useState(buildThreadState)
  const [activityTimeline, setActivityTimeline] = useState(buildTimelineState)
  const [mediaLibrary, setMediaLibrary] = useState(buildMediaLibraryState)
  const [notice, setNotice] = useState('Selecione uma acao para simular o fluxo operacional da inbox.')
  const [activeQueueMenuId, setActiveQueueMenuId] = useState(null)
  const [activeMessageMenuId, setActiveMessageMenuId] = useState(null)
  const [attachmentTab, setAttachmentTab] = useState('new')
  const [mediaSearch, setMediaSearch] = useState('')
  const [composerSubscribed, setComposerSubscribed] = useState(true)
  const fileInputRef = useRef(null)

  const selectedConversation = conversations.find((item) => item.id === selectedConversationId) ?? conversations[0]
  const selectedThread = messageThreads[selectedConversationId] ?? []
  const visibleMedia = (mediaLibrary[selectedConversationId] ?? []).filter((item) => item.name.toLowerCase().includes(mediaSearch.toLowerCase()))
  const visibleConversations = [...conversations]
    .filter((item) => item.status === activeTab)
    .filter((item) => `${item.name} ${item.preview}`.toLowerCase().includes(searchTerm.toLowerCase()))
    .sort((left, right) => {
      if (left.pinned !== right.pinned) return Number(right.pinned) - Number(left.pinned)
      if (sortMode === 'Ultima mensagem') return right.time.localeCompare(left.time)
      if (sortMode === 'Esperando resposta') return Number(right.unread > 0) - Number(left.unread > 0)
      return left.id - right.id
    })

  useEffect(() => {
    if (selectedConversation && selectedConversation.status === activeTab) return
    const nextConversation = conversations.find((item) => item.status === activeTab) ?? conversations[0]
    if (nextConversation && nextConversation.id !== selectedConversationId) {
      setSelectedConversationId(nextConversation.id)
    }
  }, [activeTab, conversations, selectedConversation, selectedConversationId])

  useEffect(() => {
    if (selectedConversation) {
      setContactNoteDraft(selectedConversation.note)
    }
  }, [selectedConversationId, selectedConversation])

  function closeFloatingLayers() {
    setListPopover(null)
    setShowAttachments(false)
    setShowEmojiPicker(false)
    setShowStickerDrawer(false)
    setActiveQueueMenuId(null)
    setActiveMessageMenuId(null)
  }

  function openPanel(panelName) {
    closeFloatingLayers()
    setActivePanel((current) => (current === panelName ? null : panelName))
  }

  function patchConversation(targetId, updater) {
    setConversations((current) => current.map((item) => (item.id === targetId ? updater(item) : item)))
  }

  function appendTimelineEntry(targetId, text) {
    const time = nowTimeLabel()
    setActivityTimeline((current) => ({
      ...current,
      [targetId]: [{ time, text }, ...(current[targetId] ?? [])],
    }))
  }

  function handleOpenConversation(nextConversationId) {
    setSelectedConversationId(nextConversationId)
    closeFloatingLayers()
    setActivePanel(null)
    patchConversation(nextConversationId, (item) => ({ ...item, unread: 0 }))
  }

  function handleFileSelection(event) {
    const files = Array.from(event.target.files ?? [])
    if (files.length === 0) return
    setMediaLibrary((current) => ({
      ...current,
      [selectedConversationId]: [
        ...files.map((file, index) => ({
          id: Date.now() + index,
          name: file.name,
          type: file.type?.startsWith('image/') ? 'Imagem' : 'Arquivo',
          origin: 'Upload local',
          size: `${Math.max(1, Math.round(file.size / 1024))} KB`,
        })),
        ...(current[selectedConversationId] ?? []),
      ],
    }))
    setAttachmentTab('library')
    setNotice(`${files.length} arquivo(s) adicionados a biblioteca de midia da conversa.`)
    event.target.value = ''
  }

  function appendOutgoingMessage(body, author = operatorOptions[0]) {
    const trimmed = body.trim()
    if (!trimmed) return
    const time = nowTimeLabel()
    setMessageThreads((current) => ({
      ...current,
      [selectedConversationId]: [
        ...(current[selectedConversationId] ?? []),
        { id: `${selectedConversationId}-${Date.now()}`, side: 'right', author, body: trimmed, time },
      ],
    }))
    patchConversation(selectedConversationId, (item) => ({
      ...item,
      preview: trimmed,
      time,
      unread: 0,
      status: 'entrada',
      assignedTo: operatorOptions[0],
    }))
  }

  function handleQuickReplyInsert(text) {
    setComposerMode('message')
    setComposerText(text)
    setActivePanel(null)
    setNotice('Resposta rapida carregada no composer para revisao antes do envio.')
  }

  function handleChatbotRun(title) {
    patchConversation(selectedConversationId, (item) => ({
      ...item,
      preview: `Fluxo "${title}" em execucao`,
      assignedTo: 'Fluxo automatizado',
    }))
    appendTimelineEntry(selectedConversationId, `Fluxo automatizado "${title}" foi iniciado manualmente.`)
    setActivePanel(null)
    setNotice(`Chatbot "${title}" executado na conversa atual.`)
  }

  function handleSaveSchedule() {
    patchConversation(selectedConversationId, (item) => ({
      ...item,
      scheduledFor: `${scheduleDraft.date} ${scheduleDraft.time}`,
    }))
    appendTimelineEntry(selectedConversationId, `Mensagem agendada para ${scheduleDraft.date} as ${scheduleDraft.time}.`)
    setNotice(`Agendamento salvo para ${scheduleDraft.date} as ${scheduleDraft.time}.`)
    setActivePanel(null)
  }

  function handleTransfer() {
    patchConversation(selectedConversationId, (item) => ({
      ...item,
      sector: transferTarget,
      category: transferTarget,
      assignedTo: `Fila ${transferTarget.toLowerCase()}`,
      status: 'esperando',
    }))
    appendTimelineEntry(selectedConversationId, `Conversa transferida para o setor ${transferTarget}.`)
    setNotice(`Conversa movida para ${transferTarget} e marcada como aguardando.`)
    setActiveTab('esperando')
    setActivePanel(null)
  }

  function handleFinalizeSingle() {
    patchConversation(selectedConversationId, (item) => ({
      ...item,
      status: 'finalizados',
      unread: 0,
      preview: 'Conversa finalizada pela equipe.',
      assignedTo: 'Encerrada',
    }))
    appendTimelineEntry(selectedConversationId, 'Conversa finalizada manualmente.')
    setNotice('Conversa atual marcada como finalizada.')
    setActiveTab('finalizados')
  }

  function handleQueueAction(conversationId, actionLabel) {
    if (actionLabel === 'Atribuir para mim') {
      patchConversation(conversationId, (item) => ({ ...item, assignedTo: operatorOptions[0] }))
      appendTimelineEntry(conversationId, 'Conversa atribuida manualmente ao operador atual.')
    }
    if (actionLabel === 'Adicionar etiqueta') {
      patchConversation(conversationId, (item) => ({ ...item, category: item.category === 'Vip' ? 'Geral' : 'Vip' }))
      appendTimelineEntry(conversationId, 'Etiqueta visual alternada nesta conversa.')
    }
    if (actionLabel === 'Marcar como nao lida') {
      patchConversation(conversationId, (item) => ({ ...item, unread: Math.max(1, item.unread || 1) }))
    }
    if (actionLabel === 'Bloquear contato') {
      patchConversation(conversationId, (item) => ({ ...item, blocked: !item.blocked }))
      appendTimelineEntry(conversationId, 'Contato alternou entre bloqueado e desbloqueado.')
    }
    if (actionLabel === 'Abrir lado-a-lado') {
      patchConversation(conversationId, (item) => ({ ...item, splitView: !item.splitView }))
    }
    if (actionLabel === 'Finalizar conversa') {
      patchConversation(conversationId, (item) => ({ ...item, status: 'finalizados', unread: 0, preview: 'Conversa finalizada pela fila.' }))
      setActiveTab('finalizados')
    }
    if (actionLabel === 'Marcar como esperando') {
      patchConversation(conversationId, (item) => ({ ...item, status: 'esperando' }))
      setActiveTab('esperando')
    }
    if (actionLabel === 'Fixar conversa') {
      patchConversation(conversationId, (item) => ({ ...item, pinned: !item.pinned }))
    }
    setActiveQueueMenuId(null)
    setNotice(`Acao aplicada: ${actionLabel.toLowerCase()}.`)
  }

  async function handleMessageAction(message, actionLabel) {
    if (actionLabel === 'Responder') {
      setComposerText(`@resposta ${message.body}`)
      setNotice('Mensagem carregada no composer para resposta contextual.')
    }
    if (actionLabel === 'Copiar') {
      try {
        await navigator.clipboard.writeText(message.body)
        setNotice('Conteudo da mensagem copiado.')
      } catch {
        setNotice('Nao foi possivel copiar automaticamente nesta simulacao.')
      }
    }
    if (actionLabel === 'Encaminhar') {
      setComposerText(`[Encaminhado] ${message.body}`)
      setNotice('Mensagem preparada para encaminhamento interno.')
    }
    if (actionLabel === 'Apagar') {
      setMessageThreads((current) => ({
        ...current,
        [selectedConversationId]: (current[selectedConversationId] ?? []).filter((item) => item.id !== message.id),
      }))
      setNotice('Mensagem removida da thread local.')
    }
    setActiveMessageMenuId(null)
  }

  function handleBulkAction(actionLabel) {
    const selectedIds = conversations.filter((item) => item.selected && item.status === activeTab).map((item) => item.id)
    if (actionLabel === 'Selecionar todas as conversas') {
      setConversations((current) => current.map((item) => (item.status === activeTab ? { ...item, selected: true } : item)))
      setNotice(`Todas as conversas da aba ${activeTab} foram selecionadas.`)
      setListPopover(null)
      return
    }
    if (selectedIds.length === 0) {
      setNotice('Selecione pelo menos uma conversa antes de usar acoes em massa.')
      setListPopover(null)
      return
    }
    if (actionLabel === 'Distribuir para um atendente') {
      setConversations((current) =>
        current.map((item) =>
          selectedIds.includes(item.id) ? { ...item, assignedTo: operatorOptions[0], selected: false, status: 'entrada' } : item,
        ),
      )
      selectedIds.forEach((id) => appendTimelineEntry(id, 'Conversa distribuida para o operador atual.'))
      setNotice(`${selectedIds.length} conversa(s) distribuidas para o operador atual.`)
    }
    if (actionLabel === 'Transferir para outro setor') {
      setConversations((current) =>
        current.map((item) =>
          selectedIds.includes(item.id) ? { ...item, sector: 'Suporte', category: 'Suporte', selected: false, status: 'esperando' } : item,
        ),
      )
      selectedIds.forEach((id) => appendTimelineEntry(id, 'Conversa enviada para a fila de suporte.'))
      setNotice(`${selectedIds.length} conversa(s) transferidas para a fila de suporte.`)
      setActiveTab('esperando')
    }
    if (actionLabel === 'Finalizar selecionadas') {
      setConversations((current) =>
        current.map((item) =>
          selectedIds.includes(item.id)
            ? { ...item, status: 'finalizados', preview: 'Conversa finalizada pela equipe.', selected: false, unread: 0 }
            : item,
        ),
      )
      selectedIds.forEach((id) => appendTimelineEntry(id, 'Conversa finalizada por acao em massa.'))
      setNotice(`${selectedIds.length} conversa(s) finalizadas.`)
      setActiveTab('finalizados')
    }
    setListPopover(null)
  }

  function handleSendComposer() {
    if (composerMode === 'message') {
      appendOutgoingMessage(composerText)
      appendTimelineEntry(selectedConversationId, 'Mensagem enviada manualmente pelo operador.')
      setNotice('Mensagem enviada na conversa atual.')
    } else {
      const trimmed = composerText.trim()
      if (!trimmed) return
      patchConversation(selectedConversationId, (item) => ({ ...item, note: trimmed }))
      setContactNoteDraft(trimmed)
      appendTimelineEntry(selectedConversationId, `Nota interna adicionada: ${trimmed}`)
      setNotice('Nota interna registrada para a equipe.')
    }
    setComposerText('')
  }

  return (
    <section className="conversation-screen">
      <aside className="conversation-list">
        <div className="conversation-list-header">
          <h1>Conversas</h1>
          <div className="conversation-list-tools">
            <button className="icon-button small" onClick={() => setListPopover((current) => (current === 'bulk' ? null : 'bulk'))}><Glyph name="filter" /></button>
            <button className="icon-button small" onClick={() => setShowContactPicker(true)}><Glyph name="plus" /></button>
          </div>
        </div>

        <div className="search-field compact">
          <Glyph name="search" />
          <input type="text" placeholder="Buscar por nome ou telefone" value={searchTerm} onChange={(event) => setSearchTerm(event.target.value)} />
        </div>

        <div className="tab-strip">
          {tabDefinitions.map(([key, label]) => (
            <button key={key} className={`pill ${activeTab === key ? 'active' : ''}`} onClick={() => { setActiveTab(key); setListPopover(null) }}>
              {label}
              <span>{conversations.filter((item) => item.status === key).length}</span>
            </button>
          ))}
        </div>

        <div className="trial-banner">
          <div className="trial-circle">7</div>
          <div>
            <strong>Aproveite seus 7 dias gratis para testar o Umbler Talk!</strong>
            <button>Ja quero assinar</button>
          </div>
        </div>

        <div className="conversation-subtoolbar">
          <button className={`select-button compact ${listPopover === 'bulk' ? 'active' : ''}`} onClick={() => setListPopover((current) => (current === 'bulk' ? null : 'bulk'))}>
            <span>Acoes em massa</span>
            <Glyph name="menu" />
          </button>
          <button className={`select-button compact ${listPopover === 'sort' ? 'active' : ''}`} onClick={() => setListPopover((current) => (current === 'sort' ? null : 'sort'))}>
            <span>Ordenar por</span>
            <Glyph name="arrowDown" />
          </button>
        </div>

        {listPopover === 'bulk' && (
          <div className="floating-menu list-menu left-column">
            <strong>Acoes em massa</strong>
            {bulkActionItems.map((label) => (
              <button key={label} onClick={() => handleBulkAction(label)}>
                <span>{label}</span>
              </button>
            ))}
          </div>
        )}

        {listPopover === 'sort' && (
          <div className="floating-menu list-menu left-column sort-menu">
            <strong>Ordenar por</strong>
            {sortActionItems.map((label) => (
              <button key={label} className={sortMode === label ? 'selected' : ''} onClick={() => { setSortMode(label); setListPopover(null); setNotice(`Lista ordenada por ${label.toLowerCase()}.`) }}>
                <span>{label}</span>
                {sortMode === label && <Glyph name="check" />}
              </button>
            ))}
          </div>
        )}

        <div className="conversation-cards">
          {visibleConversations.map((item) => (
            <article key={item.id} className={`conversation-card ${selectedConversationId === item.id ? 'selected' : ''}`} onClick={() => handleOpenConversation(item.id)}>
              <label className="card-selector" onClick={(event) => event.stopPropagation()}>
                <input
                  type="checkbox"
                  checked={item.selected}
                  onChange={(event) =>
                    setConversations((current) =>
                      current.map((row) => (row.id === item.id ? { ...row, selected: event.target.checked } : row)),
                    )
                  }
                />
              </label>
              <div className="avatar" style={{ '--avatar': item.color }}>
                {item.name[0]}
              </div>
              <div className="conversation-card-body">
                <div className="conversation-row">
                  <div className="card-title-row">
                    <strong>{item.name}</strong>
                    {item.pinned && <span className="mini-meta">Fixada</span>}
                    {item.blocked && <span className="mini-meta">Bloqueado</span>}
                  </div>
                  <span>{item.time}</span>
                </div>
                <p>{item.preview}</p>
                <div className="conversation-row">
                  <div className="card-meta-row">
                    <span className="category-tag">{item.category || 'Bot'}</span>
                    <span className="mini-meta">{item.assignedTo}</span>
                    {item.splitView && <span className="mini-meta">Lado a lado</span>}
                  </div>
                  <div className="queue-row-actions">
                    {item.unread > 0 && <span className="unread-dot">{item.unread}</span>}
                    <button className="icon-button tiny subtle" onClick={(event) => { event.stopPropagation(); setActiveQueueMenuId((current) => (current === item.id ? null : item.id)) }}>
                      <Glyph name="menu" />
                    </button>
                    {activeQueueMenuId === item.id && (
                      <div className="floating-menu queue-menu">
                        {queueActionItems.map((action) => (
                          <button key={action} onClick={(event) => { event.stopPropagation(); handleQueueAction(item.id, action) }}>
                            <span>{action}</span>
                          </button>
                        ))}
                      </div>
                    )}
                  </div>
                </div>
              </div>
            </article>
          ))}
          {visibleConversations.length === 0 && (
            <div className="empty-inline-state">
              <strong>Nenhuma conversa encontrada</strong>
              <span>Ajuste a busca, troque a aba ou abra um contato sintetico para continuar o refinamento.</span>
            </div>
          )}
        </div>
      </aside>
      <div className="conversation-stage">
        <header className="chat-topbar">
          <div className="chat-person">
            <div className="avatar large" style={{ '--avatar': selectedConversation.color }}>{selectedConversation.name[0]}</div>
            <div>
              <strong>{selectedConversation.name}</strong>
              <div className="chat-person-meta">
                <span className="category-tag">{selectedConversation.category || 'Bot'}</span>
                <span className="presence-text">
                  {selectedConversation.status === 'finalizados'
                    ? 'Conversa encerrada'
                    : selectedConversation.status === 'esperando'
                      ? 'Aguardando retorno da equipe'
                      : 'Atendimento em andamento'}
                </span>
              </div>
            </div>
          </div>
          <div className="chat-toolbar">
            <button className={`toolbar-chip ${activePanel === 'contact' ? 'active' : ''}`} onClick={() => openPanel('contact')}>Contato</button>
            <button className={`toolbar-chip ${activePanel === 'details' ? 'active' : ''}`} onClick={() => openPanel('details')}>Detalhes da conversa</button>
            <button className={`toolbar-chip ${activePanel === 'schedule' ? 'active' : ''}`} onClick={() => openPanel('schedule')}>Agendar envio</button>
            <button className={`icon-button small ${activePanel === 'transfer' ? 'is-active' : ''}`} title="Transferir conversa" onClick={() => openPanel('transfer')}><Glyph name="send" /></button>
            <button className="icon-button small" title="Finalizar conversa" onClick={handleFinalizeSingle}><Glyph name="check" /></button>
          </div>
        </header>

        <div className="chat-canvas">
          <span className="date-pill">Hoje</span>
          <div className="status-banner">
            <strong>Estado atual</strong>
            <span>{notice}</span>
          </div>
          {selectedThread.map((message) => (
            <div key={message.id} className={`bubble ${message.side}`}>
              <div className="bubble-head">
                <div>
                  {message.author && <strong>{message.author}</strong>}
                </div>
                <button className="icon-button tiny subtle" onClick={() => setActiveMessageMenuId((current) => (current === message.id ? null : message.id))}>
                  <Glyph name="menu" />
                </button>
              </div>
              <span>{message.body}</span>
              <small>{message.time}</small>
              {activeMessageMenuId === message.id && (
                <div className="floating-menu bubble-menu">
                  {['Responder', 'Copiar', 'Encaminhar', 'Apagar'].map((action) => (
                    <button key={action} onClick={() => handleMessageAction(message, action)}>
                      <span>{action}</span>
                    </button>
                  ))}
                </div>
              )}
            </div>
          ))}

          {showEmojiPicker && (
            <div className="floating-menu emoji-picker">
              <strong>Reacoes rapidas</strong>
              <div className="emoji-grid">
                {[':-)', ';-)', '<3', 'ok', 'wow', 'oi', '++', 'up'].map((emoji) => (
                  <button key={emoji} className="emoji-cell" onClick={() => { setComposerText((current) => `${current} ${emoji}`.trim()); setNotice('Reacao inserida no composer para edicao.'); }}>
                    {emoji}
                  </button>
                ))}
              </div>
            </div>
          )}

          {showAttachments && (
            <aside className="attachment-side-panel">
              <div className="side-panel-header">
                <div>
                  <h2>Anexar arquivo</h2>
                </div>
                <button className="icon-button small" onClick={() => setShowAttachments(false)}><Glyph name="x" /></button>
              </div>
              <div className="attachment-tabs">
                <button className={`toolbar-chip ${attachmentTab === 'new' ? 'active' : ''}`} onClick={() => setAttachmentTab('new')}>Novo arquivo</button>
                <button className={`toolbar-chip ${attachmentTab === 'library' ? 'active' : ''}`} onClick={() => setAttachmentTab('library')}>Biblioteca de midia</button>
              </div>
              <div className="attachment-toolbar">
                <div className="search-field panel-search">
                  <Glyph name="search" />
                  <input type="text" placeholder="Pesquisar" value={mediaSearch} onChange={(event) => setMediaSearch(event.target.value)} />
                </div>
                <button className="select-button">Todos os arquivos</button>
              </div>
              {attachmentTab === 'new' ? (
                <div className="attachment-empty">
                  <strong>Selecione um arquivo do seu computador</strong>
                  <span>Ao clicar abaixo, a janela do Windows Explorer sera aberta para escolher um ou mais arquivos.</span>
                  <button className="primary-button compact" onClick={() => fileInputRef.current?.click()}>Escolher arquivo</button>
                </div>
              ) : visibleMedia.length > 0 ? (
                <div className="media-library-list">
                  {visibleMedia.map((item) => (
                    <button key={item.id} className="media-row" onClick={() => { setComposerText((current) => `${current}\n[Midia selecionada: ${item.name}]`.trim()); setNotice(`Midia "${item.name}" preparada no composer.`); setShowAttachments(false) }}>
                      <div>
                        <strong>{item.name}</strong>
                        <span>{item.type} · {item.size}</span>
                      </div>
                      <span className="mini-meta">{item.origin}</span>
                    </button>
                  ))}
                </div>
              ) : (
                <div className="attachment-empty">
                  <strong>Nenhum arquivo encontrado</strong>
                  <span>A biblioteca de midia exibe todos os arquivos anexados na conversa atual.</span>
                </div>
              )}
            </aside>
          )}

          {showStickerDrawer && (
            <div className="sticker-drawer">
              <div className="sticker-header">
                <strong>Figurinhas salvas</strong>
                <button className="icon-button small" onClick={() => setShowStickerDrawer(false)}><Glyph name="x" /></button>
              </div>
              <div className="sticker-grid">
                {stickerItems.map((label) => (
                  <button key={label} className="sticker-card" onClick={() => { setComposerText((current) => `${current}\n[Figurinha: ${label}]`.trim()); setNotice(`Figurinha sintetica "${label}" adicionada ao composer.`); setShowStickerDrawer(false) }}>
                    <span className="sticker-preview">{label.slice(0, 2)}</span>
                    <small>{label}</small>
                  </button>
                ))}
              </div>
            </div>
          )}

          {activePanel === 'contact' && (
            <ConversationSidePanel
              title="Contato"
              subtitle="Dados gerais e atalhos de acao"
              onClose={() => setActivePanel(null)}
              footer={<button className="primary-button compact" onClick={() => { patchConversation(selectedConversationId, (item) => ({ ...item, note: contactNoteDraft })); setNotice('Observacoes do contato atualizadas para a simulacao.'); }}>Salvar observacoes</button>}
            >
              <div className="panel-profile">
                <div className="avatar medium" style={{ '--avatar': selectedConversation.color }}>{selectedConversation.name[0]}</div>
                <div>
                  <strong>{selectedConversation.name}</strong>
                  <span>Telefone oculto</span>
                </div>
              </div>

              <div className="panel-chip-row">
                <span className="chip">Cliente ativo</span>
                <span className="chip">WhatsApp</span>
                <span className="chip">{selectedConversation.sector}</span>
              </div>

              <div className="panel-section">
                <h3>Resumo</h3>
                <div className="detail-list">
                  {contactInsightItems.map(([label, value]) => (
                    <div key={label} className="detail-row">
                      <span>{label}</span>
                      <strong>{label === 'Setor sugerido' ? selectedConversation.sector : label === 'Ultimo operador' ? selectedConversation.assignedTo : value}</strong>
                    </div>
                  ))}
                </div>
              </div>

              <div className="panel-section">
                <h3>Campos adicionais</h3>
                <label className="form-field mini">
                  <span>Nome interno</span>
                  <input type="text" value={selectedConversation.name} readOnly />
                </label>
                <label className="form-field mini">
                  <span>Observacoes</span>
                  <textarea rows={4} value={contactNoteDraft} onChange={(event) => setContactNoteDraft(event.target.value)} />
                </label>
              </div>
            </ConversationSidePanel>
          )}

          {activePanel === 'details' && (
            <ConversationSidePanel
              title="Detalhes da conversa"
              subtitle="Historico operacional e status da fila"
              onClose={() => setActivePanel(null)}
            >
              <div className="panel-chip-row">
                <span className="chip">{selectedConversation.status === 'finalizados' ? 'Conversa encerrada' : 'Atendimento ativo'}</span>
                <span className="chip">
                  {selectedConversation.priority === 'alta' ? 'Prioridade alta' : selectedConversation.priority === 'media' ? 'Prioridade media' : 'Operacao padrao'}
                </span>
              </div>
              <div className="panel-section">
                <h3>Informacoes gerais</h3>
                <div className="detail-list">
                  {conversationDetailItems.map(([label, value]) => (
                    <div key={label} className="detail-row">
                      <span>{label}</span>
                      <strong>
                        {label === 'Status'
                          ? selectedConversation.status === 'finalizados'
                            ? 'Conversa finalizada'
                            : selectedConversation.status === 'esperando'
                              ? 'Aguardando resposta da equipe'
                              : 'Aguardando resposta da equipe'
                          : label === 'Etiqueta ativa'
                            ? selectedConversation.priority === 'alta'
                              ? 'Prioridade alta'
                              : selectedConversation.priority === 'media'
                                ? 'Prioridade media'
                                : 'Operacao padrao'
                            : label === 'Ultima atividade'
                              ? selectedConversation.time
                              : value}
                      </strong>
                    </div>
                  ))}
                </div>
              </div>
              <div className="panel-section">
                <h3>Timeline</h3>
                <div className="timeline-list">
                  {(activityTimeline[selectedConversationId] ?? []).map((entry) => (
                    <div key={`${entry.time}-${entry.text}`}>
                      <strong>{entry.time}</strong>
                      <span>{entry.text}</span>
                    </div>
                  ))}
                </div>
              </div>
            </ConversationSidePanel>
          )}

          {activePanel === 'quickReplies' && (
            <ConversationSidePanel
              title="Respostas rapidas"
              subtitle="Atalhos prontos para acelerar o atendimento"
              onClose={() => setActivePanel(null)}
            >
              <div className="search-field panel-search">
                <Glyph name="search" />
                <input type="text" placeholder="Buscar atalho ou palavra-chave" />
              </div>
              {quickReplyGroups.map((group) => (
                <div key={group.title} className="panel-section">
                  <h3>{group.title}</h3>
                  <div className="stack-list">
                    {group.items.map((item) => (
                      <button key={item.shortcut} className="reply-card" onClick={() => handleQuickReplyInsert(item.body)}>
                        <strong>{item.title}</strong>
                        <span>{item.body}</span>
                        <small>{item.shortcut}</small>
                      </button>
                    ))}
                  </div>
                </div>
              ))}
            </ConversationSidePanel>
          )}

          {activePanel === 'chatbot' && (
            <ConversationSidePanel
              title="Executar chatbot"
              subtitle="Fluxos automatizados disponiveis para esta conversa"
              onClose={() => setActivePanel(null)}
            >
              <div className="stack-list">
                {chatbotPlaybooks.map((item) => (
                  <div key={item.title} className="automation-card">
                    <div>
                      <strong>{item.title}</strong>
                      <span>{item.description}</span>
                    </div>
                    <button className="select-button compact" onClick={() => handleChatbotRun(item.title)}>Executar</button>
                  </div>
                ))}
              </div>
            </ConversationSidePanel>
          )}

          {activePanel === 'transfer' && (
            <ConversationSidePanel
              title="Transferir conversa"
              subtitle="Reencaminhe este atendimento para outra fila ou setor"
              onClose={() => setActivePanel(null)}
              footer={<button className="primary-button compact" onClick={handleTransfer}>Confirmar transferencia</button>}
            >
              <div className="panel-section">
                <h3>Destino</h3>
                <div className="stack-list">
                  {sectorOptions.map((option) => (
                    <button key={option} className={`option-card ${transferTarget === option ? 'selected' : ''}`} onClick={() => setTransferTarget(option)}>
                      <strong>{option}</strong>
                      <span>{option === 'Geral' ? 'Mantem o atendimento na fila principal.' : `Direciona a conversa para o setor ${option.toLowerCase()}.`}</span>
                    </button>
                  ))}
                </div>
              </div>
            </ConversationSidePanel>
          )}

          {activePanel === 'schedule' && (
            <ConversationSidePanel
              title="Agendamento de mensagens"
              subtitle="Escolha quando a proxima resposta deve ser enviada"
              onClose={() => setActivePanel(null)}
              footer={<button className="primary-button compact" onClick={handleSaveSchedule}>Salvar agendamento</button>}
            >
              <div className="panel-section">
                <h3>Mensagem preparada</h3>
                <div className="schedule-preview">
                  <span>{composerText || 'Retorno de acompanhamento configurado para envio futuro.'}</span>
                </div>
              </div>
              <div className="panel-section">
                <h3>Data e hora</h3>
                <div className="schedule-grid">
                  <label className="form-field mini">
                    <span>Data</span>
                    <input type="text" value={scheduleDraft.date} onChange={(event) => setScheduleDraft((current) => ({ ...current, date: event.target.value }))} />
                  </label>
                  <label className="form-field mini">
                    <span>Hora</span>
                    <input type="text" value={scheduleDraft.time} onChange={(event) => setScheduleDraft((current) => ({ ...current, time: event.target.value }))} />
                  </label>
                </div>
              </div>
              {schedulerBlocks.map((block) => (
                <div key={block.title} className="panel-section">
                  <h3>{block.title}</h3>
                  <div className="panel-chip-row">
                    {block.options.map((option) => (
                      <button key={option} className="chip interactive" onClick={() => setScheduleDraft((current) => ({ ...current, time: option }))}>{option}</button>
                    ))}
                  </div>
                </div>
              ))}
            </ConversationSidePanel>
          )}
        </div>

        <footer className="chat-composer">
          <div className="composer-header">
            <label className="switch">
              <input type="checkbox" checked={composerSubscribed} onChange={(event) => setComposerSubscribed(event.target.checked)} />
              <span />
            </label>
            <strong>{composerSubscribed ? selectedConversation.assignedTo : 'Assinar'}</strong>
          </div>
          <div className="composer-tabs">
            <button className={`pill ${composerMode === 'message' ? 'active' : 'ghost'}`} onClick={() => setComposerMode('message')}>Mensagem</button>
            <button className={`pill ${composerMode === 'note' ? 'active' : 'ghost'}`} onClick={() => setComposerMode('note')}>Notas</button>
          </div>
          <textarea
            placeholder={composerMode === 'message' ? 'Digite sua mensagem ou arraste um arquivo...' : 'Registre uma observacao interna para a equipe...'}
            rows={3}
            value={composerText}
            onChange={(event) => setComposerText(event.target.value)}
          />
          <div className="composer-tools">
            <div className="tool-row">
              <button className={`icon-button tiny ${showAttachments ? 'is-active' : ''}`} onClick={() => { setShowAttachments((current) => !current); setAttachmentTab('new'); setShowEmojiPicker(false); setShowStickerDrawer(false) }} title="Anexar"><Glyph name="paperclip" /></button>
              <button className={`icon-button tiny ${showEmojiPicker ? 'is-active' : ''}`} onClick={() => { setShowEmojiPicker((current) => !current); setShowAttachments(false); setShowStickerDrawer(false) }} title="Emojis"><Glyph name="smile" /></button>
              <button className={`icon-button tiny ${showStickerDrawer ? 'is-active' : ''}`} onClick={() => { setShowStickerDrawer((current) => !current); setShowAttachments(false); setShowEmojiPicker(false) }} title="Figurinhas"><Glyph name="image" /></button>
              <button className={`icon-button tiny ${activePanel === 'quickReplies' ? 'is-active' : ''}`} onClick={() => openPanel('quickReplies')} title="Respostas rapidas"><Glyph name="spark" /></button>
              <button className={`icon-button tiny ${activePanel === 'chatbot' ? 'is-active' : ''}`} onClick={() => openPanel('chatbot')} title="Executar chatbot"><Glyph name="bot" /></button>
              <button className={`icon-button tiny ${activePanel === 'schedule' ? 'is-active' : ''}`} onClick={() => openPanel('schedule')} title="Agendar"><Glyph name="calendar" /></button>
            </div>
            <button className="voice-button" onClick={handleSendComposer}><Glyph name={composerMode === 'message' ? 'send' : 'note'} /></button>
          </div>
          <input ref={fileInputRef} type="file" multiple className="hidden-file-input" onChange={handleFileSelection} />
        </footer>
      </div>
      {showContactPicker && (
        <ContactPickerModal
          onClose={() => setShowContactPicker(false)}
          onOpenContact={(contactId) => {
            setShowContactPicker(false)
            setActiveTab('entrada')
            patchConversation(contactId, (item) => ({ ...item, status: 'entrada', unread: 0 }))
            handleOpenConversation(contactId)
            setNotice('Contato sintetico aberto na inbox para refinamento visual.')
          }}
        />
      )}
    </section>
  )
}

function ConversationSidePanel({ title, subtitle, children, footer, onClose }) {
  return (
    <aside className="conversation-side-panel">
      <div className="side-panel-header">
        <div>
          <h2>{title}</h2>
          <p>{subtitle}</p>
        </div>
        <button className="icon-button small" onClick={onClose}><Glyph name="x" /></button>
      </div>
      <div className="side-panel-body">{children}</div>
      {footer && <div className="side-panel-footer">{footer}</div>}
    </aside>
  )
}

function ContactPickerModal({ onClose, onOpenContact }) {
  return (
    <div className="modal-scrim" onClick={onClose}>
      <div className="contact-picker-modal" onClick={(event) => event.stopPropagation()}>
        <div className="side-panel-header">
          <div>
            <h2>Escolha um contato</h2>
            <p>Selecione um perfil sintetico para iniciar uma nova conversa visual.</p>
          </div>
          <button className="icon-button small" onClick={onClose}><Glyph name="x" /></button>
        </div>
        <div className="search-field">
          <Glyph name="search" />
          <input type="text" placeholder="Pesquisar contato" />
        </div>
        <div className="contact-picker-list">
          {contactItems.map((item) => (
            <button key={item.id} className="contact-picker-row" onClick={() => onOpenContact(item.id)}>
              <div className="contact-cell">
                <div className="avatar small" style={{ '--avatar': item.color }}>{item.name[0]}</div>
                <div>
                  <strong>{item.name}</strong>
                  <span>{item.phone}</span>
                </div>
              </div>
              <span className="chip interactive">Abrir</span>
            </button>
          ))}
        </div>
      </div>
    </div>
  )
}

function ContactsPage() {
  const [contacts, setContacts] = useState(buildContactState)
  const [searchTerm, setSearchTerm] = useState('')
  const [tagFilter, setTagFilter] = useState('Todos')
  const [sortMode, setSortMode] = useState('Nome')
  const [activeContactId, setActiveContactId] = useState(contactItems[0]?.id ?? 1)
  const [activeContactMenuId, setActiveContactMenuId] = useState(null)
  const [showNewContactModal, setShowNewContactModal] = useState(false)
  const [draftContact, setDraftContact] = useState({
    name: '',
    phone: '',
    email: '',
    tag: 'Lead',
    company: '',
  })

  const filteredContacts = [...contacts]
    .filter((item) => tagFilter === 'Todos' || item.tag === tagFilter)
    .filter((item) => `${item.name} ${item.phone} ${item.email}`.toLowerCase().includes(searchTerm.toLowerCase()))
    .sort((left, right) => {
      if (sortMode === 'Ultima interacao') return (right.note || '').localeCompare(left.note || '')
      return left.name.localeCompare(right.name)
    })

  const activeContact = contacts.find((item) => item.id === activeContactId) ?? filteredContacts[0] ?? contacts[0]

  function patchContact(targetId, updater) {
    setContacts((current) => current.map((item) => (item.id === targetId ? updater(item) : item)))
  }

  function resetDraft() {
    setDraftContact({ name: '', phone: '', email: '', tag: 'Lead', company: '' })
  }

  function createContact() {
    if (!draftContact.name.trim()) return
    const nextId = Math.max(...contacts.map((item) => item.id)) + 1
    const nextContact = {
      id: nextId,
      name: draftContact.name.trim(),
      phone: draftContact.phone.trim() || 'Telefone oculto',
      note: 'ha instantes',
      color: '#5f7df5',
      selected: false,
      email: draftContact.email.trim() || `novo${nextId}@empresa.local`,
      tag: draftContact.tag,
      origin: 'Cadastro manual',
      company: draftContact.company.trim() || 'Operacao interna',
      status: 'Ativo',
      notes: 'Contato criado localmente para validar o fluxo do frontend.',
    }
    setContacts((current) => [nextContact, ...current])
    setActiveContactId(nextId)
    setShowNewContactModal(false)
    resetDraft()
  }

  return (
    <section className="content-page">
      <div className="page-heading">
        <h1>Contatos <span>{contacts.length}</span></h1>
        <p>Aqui voce pode gerenciar as informacoes dos seus contatos e acessar os historicos de mensagens</p>
      </div>

      <div className="card table-card">
        <div className="table-toolbar">
          <div className="search-field">
            <Glyph name="search" />
            <input type="text" placeholder="Pesquisar" value={searchTerm} onChange={(event) => setSearchTerm(event.target.value)} />
          </div>
          <button className="select-button" onClick={() => setTagFilter((current) => (current === 'Todos' ? 'Vip' : 'Todos'))}>{tagFilter}</button>
          <button className="select-button" onClick={() => setTagFilter((current) => (current === 'Lead' ? 'Suporte' : 'Lead'))}>Etiquetas: {tagFilter === 'Todos' ? 'Lead' : tagFilter}</button>
          <button className="select-button" onClick={() => setSortMode((current) => (current === 'Nome' ? 'Ultima interacao' : 'Nome'))}>Ordenar por: {sortMode}</button>
          <button className="icon-button subtle" onClick={() => setTagFilter('Todos')}><Glyph name="filter" /></button>
          <div className="toolbar-spacer" />
          <button className="primary-button compact" onClick={() => setShowNewContactModal(true)}>Adicionar contato</button>
        </div>

        <div className="table-head">
          <span />
          <span>Contato</span>
          <span>Acoes</span>
        </div>

        {filteredContacts.map((item) => (
          <div key={item.id} className={`table-row interactive ${activeContact?.id === item.id ? 'is-selected' : ''}`}>
            <input
              type="checkbox"
              checked={item.selected}
              onChange={(event) => patchContact(item.id, (current) => ({ ...current, selected: event.target.checked }))}
            />
            <div className="contact-cell">
              <div className="avatar medium" style={{ '--avatar': item.color }} onClick={() => setActiveContactId(item.id)}>
                {item.name[0]}
              </div>
              <div onClick={() => setActiveContactId(item.id)}>
                <strong>{item.name}</strong>
                {item.note && <small>{item.note}</small>}
                <span>{item.phone}</span>
                <span className="inline-chip-row">
                  <span className="chip">{item.tag}</span>
                  <span className="mini-meta">{item.status}</span>
                </span>
              </div>
            </div>
            <div className="row-actions">
              <button className="icon-button subtle" onClick={() => setActiveContactId(item.id)}><Glyph name="user" /></button>
              <button className="icon-button subtle" onClick={() => setActiveContactMenuId((current) => (current === item.id ? null : item.id))}><Glyph name="menu" /></button>
              {activeContactMenuId === item.id && (
                <div className="dropdown-menu row-dropdown">
                  <button onClick={() => { setActiveContactId(item.id); setActiveContactMenuId(null) }}>Ver detalhes</button>
                  <button onClick={() => { patchContact(item.id, (current) => ({ ...current, status: current.status === 'Ativo' ? 'Inativo' : 'Ativo' })); setActiveContactMenuId(null) }}>
                    {item.status === 'Ativo' ? 'Desativar' : 'Ativar'}
                  </button>
                  <button onClick={() => { patchContact(item.id, (current) => ({ ...current, tag: current.tag === 'Vip' ? 'Lead' : 'Vip' })); setActiveContactMenuId(null) }}>Alternar etiqueta</button>
                </div>
              )}
            </div>
          </div>
        ))}

        {filteredContacts.length === 0 && (
          <div className="empty-inline-state in-card">
            <strong>Nenhum contato encontrado</strong>
            <span>Ajuste a busca ou crie um novo contato sintetico para continuar os testes.</span>
          </div>
        )}
      </div>

      {activeContact && (
        <section className="card contact-detail-card">
          <div className="detail-header">
            <div className="contact-cell">
              <div className="avatar medium" style={{ '--avatar': activeContact.color }}>{activeContact.name[0]}</div>
              <div>
                <strong>{activeContact.name}</strong>
                <span>{activeContact.phone}</span>
                <span>{activeContact.email}</span>
              </div>
            </div>
            <div className="panel-chip-row">
              <span className="chip">{activeContact.tag}</span>
              <span className="chip">{activeContact.status}</span>
            </div>
          </div>

          <div className="detail-grid">
            <div className="detail-list">
              <div className="detail-row"><span>Empresa</span><strong>{activeContact.company}</strong></div>
              <div className="detail-row"><span>Origem</span><strong>{activeContact.origin}</strong></div>
              <div className="detail-row"><span>Ultima interacao</span><strong>{activeContact.note || 'Sem registro'}</strong></div>
            </div>
            <label className="form-field mini">
              <span>Observacoes</span>
              <textarea rows={4} value={activeContact.notes} onChange={(event) => patchContact(activeContact.id, (current) => ({ ...current, notes: event.target.value }))} />
            </label>
          </div>
        </section>
      )}

      {showNewContactModal && (
        <div className="modal-scrim" onClick={() => { setShowNewContactModal(false); resetDraft() }}>
          <div className="contact-picker-modal form-modal" onClick={(event) => event.stopPropagation()}>
            <div className="side-panel-header">
              <div>
                <h2>Novo contato</h2>
                <p>Cadastro sintetico para tornar o frontend utilizavel durante o refinamento.</p>
              </div>
              <button className="icon-button small" onClick={() => { setShowNewContactModal(false); resetDraft() }}><Glyph name="x" /></button>
            </div>
            <div className="modal-form-grid">
              <label className="form-field mini">
                <span>Nome</span>
                <input type="text" value={draftContact.name} onChange={(event) => setDraftContact((current) => ({ ...current, name: event.target.value }))} />
              </label>
              <label className="form-field mini">
                <span>Telefone</span>
                <input type="text" value={draftContact.phone} onChange={(event) => setDraftContact((current) => ({ ...current, phone: event.target.value }))} />
              </label>
              <label className="form-field mini">
                <span>E-mail</span>
                <input type="text" value={draftContact.email} onChange={(event) => setDraftContact((current) => ({ ...current, email: event.target.value }))} />
              </label>
              <label className="form-field mini">
                <span>Empresa</span>
                <input type="text" value={draftContact.company} onChange={(event) => setDraftContact((current) => ({ ...current, company: event.target.value }))} />
              </label>
              <div className="panel-section">
                <h3>Etiqueta</h3>
                <div className="panel-chip-row">
                  {contactTagOptions.filter((item) => item !== 'Todos').map((tag) => (
                    <button key={tag} className={`chip interactive ${draftContact.tag === tag ? 'active-chip' : ''}`} onClick={() => setDraftContact((current) => ({ ...current, tag }))}>{tag}</button>
                  ))}
                </div>
              </div>
            </div>
            <div className="side-panel-footer">
              <button className="primary-button compact" onClick={createContact}>Salvar contato</button>
            </div>
          </div>
        </div>
      )}
    </section>
  )
}

function SettingsPage() {
  return (
    <section className="content-page">
      <div className="card settings-card">
        {settingsGroups.map(([title, description]) => {
          const match = settingRoutes.find((item) => item.title === title)
          return (
            <button
              key={title}
              className="settings-row"
              onClick={() => match ? navigate(match.path) : undefined}
            >
              <span className="settings-icon"><Glyph name={title.includes('Financeiro') ? 'money' : title.includes('Etiquetas') ? 'tag' : title.includes('Atendentes') ? 'user' : title.includes('Chatbots') ? 'bot' : title.includes('Canais') ? 'chat' : 'spark'} /></span>
              <span className="settings-copy">
                <strong>{title}</strong>
                <small>{description}</small>
              </span>
              <span className="settings-arrow">›</span>
            </button>
          )
        })}
      </div>
    </section>
  )
}

function ChannelsPage() {
  const [channels, setChannels] = useState(buildChannelSettingsState)
  const [showModal, setShowModal] = useState(false)
  const [draftChannel, setDraftChannel] = useState({ name: '', type: 'WhatsApp', owner: 'Operador atual' })

  function createChannel() {
    if (!draftChannel.name.trim()) return
    const nextId = Math.max(...channels.map((item) => item.id)) + 1
    setChannels((current) => [
      {
        id: nextId,
        name: draftChannel.name.trim(),
        type: draftChannel.type,
        status: 'Em configuracao',
        owner: draftChannel.owner,
        selected: false,
      },
      ...current,
    ])
    setShowModal(false)
    setDraftChannel({ name: '', type: 'WhatsApp', owner: 'Operador atual' })
  }

  return (
    <section className="content-page">
      <div className="page-heading compact">
        <button className="icon-button subtle back-button" onClick={() => navigate('/settings')}>{'<'}</button>
        <div>
          <h1>Canais de atendimento</h1>
          <p>Gerencie os canais conectados ao ambiente de atendimento e o responsavel operacional de cada um.</p>
        </div>
      </div>

      <div className="card table-card">
        <div className="table-toolbar">
          <div className="search-field">
            <Glyph name="chat" />
            <input type="text" value="Canais ativos" readOnly />
          </div>
          <div className="toolbar-spacer" />
          <button className="primary-button compact" onClick={() => setShowModal(true)}>Novo canal</button>
        </div>

        <div className="table-head campaigns templates-grid">
          <span>Canal</span>
          <span>Tipo</span>
          <span>Status</span>
          <span>Responsavel</span>
        </div>

        {channels.map((channel) => (
          <div className="table-row campaigns templates-grid" key={channel.id}>
            <strong>{channel.name}</strong>
            <span>{channel.type}</span>
            <button className="status-tag interactive-tag" onClick={() => setChannels((current) => current.map((item) => item.id === channel.id ? { ...item, status: item.status === 'Conectado' ? 'Em configuracao' : 'Conectado' } : item))}>
              {channel.status}
            </button>
            <span>{channel.owner}</span>
          </div>
        ))}
      </div>

      {showModal && (
        <div className="modal-scrim" onClick={() => setShowModal(false)}>
          <div className="contact-picker-modal form-modal" onClick={(event) => event.stopPropagation()}>
            <div className="side-panel-header">
              <div>
                <h2>Novo canal</h2>
                <p>Configuracao sintetica para evoluir o modulo de canais.</p>
              </div>
              <button className="icon-button small" onClick={() => setShowModal(false)}><Glyph name="x" /></button>
            </div>
            <div className="modal-form-grid">
              <label className="form-field mini">
                <span>Nome</span>
                <input type="text" value={draftChannel.name} onChange={(event) => setDraftChannel((current) => ({ ...current, name: event.target.value }))} />
              </label>
              <label className="form-field mini">
                <span>Tipo</span>
                <input type="text" value={draftChannel.type} onChange={(event) => setDraftChannel((current) => ({ ...current, type: event.target.value }))} />
              </label>
              <label className="form-field mini">
                <span>Responsavel</span>
                <input type="text" value={draftChannel.owner} onChange={(event) => setDraftChannel((current) => ({ ...current, owner: event.target.value }))} />
              </label>
            </div>
            <div className="side-panel-footer">
              <button className="primary-button compact" onClick={createChannel}>Salvar canal</button>
            </div>
          </div>
        </div>
      )}
    </section>
  )
}

function TagsPage() {
  const [tags, setTags] = useState(buildTagState)
  const [showModal, setShowModal] = useState(false)
  const [draftTag, setDraftTag] = useState({ name: '', color: '#5f7df5' })

  function createTag() {
    if (!draftTag.name.trim()) return
    const nextId = Math.max(...tags.map((item) => item.id)) + 1
    setTags((current) => [
      { id: nextId, name: draftTag.name.trim(), color: draftTag.color, usage: 0, active: true },
      ...current,
    ])
    setShowModal(false)
    setDraftTag({ name: '', color: '#5f7df5' })
  }

  return (
    <section className="content-page">
      <div className="page-heading compact">
        <button className="icon-button subtle back-button" onClick={() => navigate('/settings')}>{'<'}</button>
        <div>
          <h1>Etiquetas</h1>
          <p>Configure etiquetas sinteticas para segmentacao visual das conversas e contatos.</p>
        </div>
      </div>

      <div className="card table-card">
        <div className="table-toolbar">
          <div className="search-field">
            <Glyph name="tag" />
            <input type="text" value="Etiquetas da organizacao" readOnly />
          </div>
          <div className="toolbar-spacer" />
          <button className="primary-button compact" onClick={() => setShowModal(true)}>Nova etiqueta</button>
        </div>

        <div className="tags-grid">
          {tags.map((tag) => (
            <article key={tag.id} className="tag-card">
              <div className="tag-preview" style={{ background: tag.color }} />
              <strong>{tag.name}</strong>
              <span>{tag.usage} uso(s)</span>
              <button className="status-tag interactive-tag" onClick={() => setTags((current) => current.map((item) => item.id === tag.id ? { ...item, active: !item.active } : item))}>
                {tag.active ? 'Ativa' : 'Inativa'}
              </button>
            </article>
          ))}
        </div>
      </div>

      {showModal && (
        <div className="modal-scrim" onClick={() => setShowModal(false)}>
          <div className="contact-picker-modal form-modal" onClick={(event) => event.stopPropagation()}>
            <div className="side-panel-header">
              <div>
                <h2>Nova etiqueta</h2>
                <p>Defina nome e cor para uma etiqueta sintetica.</p>
              </div>
              <button className="icon-button small" onClick={() => setShowModal(false)}><Glyph name="x" /></button>
            </div>
            <div className="modal-form-grid">
              <label className="form-field mini">
                <span>Nome</span>
                <input type="text" value={draftTag.name} onChange={(event) => setDraftTag((current) => ({ ...current, name: event.target.value }))} />
              </label>
              <label className="form-field mini">
                <span>Cor</span>
                <input type="text" value={draftTag.color} onChange={(event) => setDraftTag((current) => ({ ...current, color: event.target.value }))} />
              </label>
            </div>
            <div className="side-panel-footer">
              <button className="primary-button compact" onClick={createTag}>Salvar etiqueta</button>
            </div>
          </div>
        </div>
      )}
    </section>
  )
}

function QuickRepliesPage() {
  const [replies, setReplies] = useState(buildQuickReplyState)
  const [showModal, setShowModal] = useState(false)
  const [draftReply, setDraftReply] = useState({ shortcut: '', title: '', body: '', scope: 'Geral' })

  function createReply() {
    if (!draftReply.shortcut.trim() || !draftReply.title.trim()) return
    const nextId = Math.max(...replies.map((item) => item.id)) + 1
    setReplies((current) => [
      { id: nextId, ...draftReply, active: true },
      ...current,
    ])
    setShowModal(false)
    setDraftReply({ shortcut: '', title: '', body: '', scope: 'Geral' })
  }

  return (
    <section className="content-page">
      <div className="page-heading compact">
        <button className="icon-button subtle back-button" onClick={() => navigate('/settings')}>{'<'}</button>
        <div>
          <h1>Respostas rapidas</h1>
          <p>Gerencie atalhos de texto reutilizaveis para os atendimentos da equipe.</p>
        </div>
      </div>

      <div className="card table-card">
        <div className="table-toolbar">
          <div className="search-field">
            <Glyph name="spark" />
            <input type="text" value="Atalhos cadastrados" readOnly />
          </div>
          <div className="toolbar-spacer" />
          <button className="primary-button compact" onClick={() => setShowModal(true)}>Nova resposta</button>
        </div>

        <div className="stack-list">
          {replies.map((reply) => (
            <div key={reply.id} className="reply-card static-card">
              <div className="reply-card-head">
                <strong>{reply.title}</strong>
                <button className="status-tag interactive-tag" onClick={() => setReplies((current) => current.map((item) => item.id === reply.id ? { ...item, active: !item.active } : item))}>
                  {reply.active ? 'Ativa' : 'Inativa'}
                </button>
              </div>
              <span>{reply.body}</span>
              <small>{reply.shortcut} · {reply.scope}</small>
            </div>
          ))}
        </div>
      </div>

      {showModal && (
        <div className="modal-scrim" onClick={() => setShowModal(false)}>
          <div className="contact-picker-modal form-modal" onClick={(event) => event.stopPropagation()}>
            <div className="side-panel-header">
              <div>
                <h2>Nova resposta rapida</h2>
                <p>Cadastre atalhos sinteticos para reaproveitamento no composer.</p>
              </div>
              <button className="icon-button small" onClick={() => setShowModal(false)}><Glyph name="x" /></button>
            </div>
            <div className="modal-form-grid">
              <label className="form-field mini">
                <span>Atalho</span>
                <input type="text" value={draftReply.shortcut} onChange={(event) => setDraftReply((current) => ({ ...current, shortcut: event.target.value }))} />
              </label>
              <label className="form-field mini">
                <span>Titulo</span>
                <input type="text" value={draftReply.title} onChange={(event) => setDraftReply((current) => ({ ...current, title: event.target.value }))} />
              </label>
              <label className="form-field mini">
                <span>Mensagem</span>
                <textarea rows={4} value={draftReply.body} onChange={(event) => setDraftReply((current) => ({ ...current, body: event.target.value }))} />
              </label>
              <label className="form-field mini">
                <span>Escopo</span>
                <input type="text" value={draftReply.scope} onChange={(event) => setDraftReply((current) => ({ ...current, scope: event.target.value }))} />
              </label>
            </div>
            <div className="side-panel-footer">
              <button className="primary-button compact" onClick={createReply}>Salvar resposta</button>
            </div>
          </div>
        </div>
      )}
    </section>
  )
}

function AgentsPage() {
  const [agents, setAgents] = useState(buildAgentState)
  const [searchTerm, setSearchTerm] = useState('')
  const [showDisabled, setShowDisabled] = useState(false)
  const [openPermissionId, setOpenPermissionId] = useState(1)
  const [openReassignmentId, setOpenReassignmentId] = useState(null)
  const [showInviteModal, setShowInviteModal] = useState(false)
  const [draftAgent, setDraftAgent] = useState({ name: '', email: '', permission: 'Operador' })

  const filteredAgents = agents
    .filter((item) => (showDisabled ? true : item.active))
    .filter((item) => `${item.name} ${item.email}`.toLowerCase().includes(searchTerm.toLowerCase()))

  function patchAgent(targetId, updater) {
    setAgents((current) => current.map((item) => (item.id === targetId ? updater(item) : item)))
  }

  function createAgent() {
    if (!draftAgent.name.trim()) return
    const nextId = Math.max(...agents.map((item) => item.id)) + 1
    setAgents((current) => [
      {
        id: nextId,
        name: draftAgent.name.trim(),
        email: draftAgent.email.trim() || `agente${nextId}@empresa.local`,
        permission: draftAgent.permission,
        reassignment: 'Desligada',
        selected: false,
        isSelf: false,
        active: true,
        color: '#8baf5a',
      },
      ...current,
    ])
    setShowInviteModal(false)
    setDraftAgent({ name: '', email: '', permission: 'Operador' })
  }

  return (
    <section className="content-page">
      <div className="page-heading compact">
        <button className="icon-button subtle back-button" onClick={() => navigate('/settings')}>{'<'}</button>
        <div>
          <h1>Atendentes</h1>
          <p>Aqui voce consegue criar ou gerenciar as pessoas que lhe ajudam com o relacionamento com seus contatos.</p>
        </div>
      </div>

      <div className="card table-card">
        <div className="table-toolbar">
          <div className="search-field">
            <Glyph name="search" />
            <input type="text" placeholder="Pesquisar" value={searchTerm} onChange={(event) => setSearchTerm(event.target.value)} />
          </div>
          <label className="checkbox-line">
            <input type="checkbox" checked={showDisabled} onChange={(event) => setShowDisabled(event.target.checked)} />
            <span>Mostrar desativados</span>
          </label>
          <div className="toolbar-spacer" />
          <button className="primary-button compact" onClick={() => setShowInviteModal(true)}>Convidar atendente</button>
        </div>

        <div className="table-head agents">
          <span />
          <span>Atendente</span>
          <span>Permissao</span>
          <span>Reatribuicao</span>
          <span>Acoes</span>
        </div>

        {filteredAgents.map((agent) => (
          <div key={agent.id} className={`table-row agents expanded ${agent.active ? '' : 'is-muted'}`}>
            <input
              type="checkbox"
              checked={agent.selected}
              onChange={(event) => patchAgent(agent.id, (current) => ({ ...current, selected: event.target.checked }))}
            />
            <div className="contact-cell">
              <div className="avatar small" style={{ '--avatar': agent.color }}>{agent.name[0]}</div>
              <div>
                <strong>{agent.name}</strong>
                <span>{agent.email}</span>
              </div>
              {agent.isSelf && <span className="self-badge">Voce</span>}
              {!agent.active && <span className="status-tag">Desativado</span>}
            </div>
            <div className={`select-surface ${openPermissionId === agent.id ? 'open' : ''}`} onClick={() => setOpenPermissionId((current) => (current === agent.id ? null : agent.id))}>
              <span>{agent.permission}</span>
              <span>v</span>
              {openPermissionId === agent.id && (
                <div className="dropdown-menu">
                  {agentPermissionOptions.map((option) => (
                    <button key={option} className={option === agent.permission ? 'selected' : ''} onClick={() => patchAgent(agent.id, (current) => ({ ...current, permission: option }))}>
                      {option}
                    </button>
                  ))}
                </div>
              )}
            </div>
            <div className={`select-surface ${openReassignmentId === agent.id ? 'open' : ''}`} onClick={() => setOpenReassignmentId((current) => (current === agent.id ? null : agent.id))}>
              <span>{agent.reassignment}</span>
              <span>v</span>
              {openReassignmentId === agent.id && (
                <div className="dropdown-menu">
                  {agentReassignmentOptions.map((option) => (
                    <button key={option} className={option === agent.reassignment ? 'selected' : ''} onClick={() => patchAgent(agent.id, (current) => ({ ...current, reassignment: option }))}>
                      {option}
                    </button>
                  ))}
                </div>
              )}
            </div>
            <button className="icon-button subtle" onClick={() => patchAgent(agent.id, (current) => ({ ...current, active: !current.active }))}><Glyph name="gear" /></button>
          </div>
        ))}
      </div>

      {showInviteModal && (
        <div className="modal-scrim" onClick={() => setShowInviteModal(false)}>
          <div className="contact-picker-modal form-modal" onClick={(event) => event.stopPropagation()}>
            <div className="side-panel-header">
              <div>
                <h2>Convidar atendente</h2>
                <p>Fluxo sintetico para cadastro de operadores internos.</p>
              </div>
              <button className="icon-button small" onClick={() => setShowInviteModal(false)}><Glyph name="x" /></button>
            </div>
            <div className="modal-form-grid">
              <label className="form-field mini">
                <span>Nome</span>
                <input type="text" value={draftAgent.name} onChange={(event) => setDraftAgent((current) => ({ ...current, name: event.target.value }))} />
              </label>
              <label className="form-field mini">
                <span>E-mail</span>
                <input type="text" value={draftAgent.email} onChange={(event) => setDraftAgent((current) => ({ ...current, email: event.target.value }))} />
              </label>
              <div className="panel-section">
                <h3>Permissao inicial</h3>
                <div className="panel-chip-row">
                  {agentPermissionOptions.map((option) => (
                    <button key={option} className={`chip interactive ${draftAgent.permission === option ? 'active-chip' : ''}`} onClick={() => setDraftAgent((current) => ({ ...current, permission: option }))}>
                      {option}
                    </button>
                  ))}
                </div>
              </div>
            </div>
            <div className="side-panel-footer">
              <button className="primary-button compact" onClick={createAgent}>Salvar atendente</button>
            </div>
          </div>
        </div>
      )}
    </section>
  )
}

function TemplatesPage() {
  const [channels, setChannels] = useState(buildTemplateState)
  const [activeChannelId, setActiveChannelId] = useState(1)
  const [showChannelModal, setShowChannelModal] = useState(false)
  const [showTemplateModal, setShowTemplateModal] = useState(false)
  const [channelDraft, setChannelDraft] = useState({ channelName: '', phone: '', provider: 'Conexao interna' })
  const [templateDraft, setTemplateDraft] = useState({ title: '', category: 'Marketing', language: 'pt-BR' })

  const activeChannel = channels.find((item) => item.id === activeChannelId) ?? channels[0]

  function createChannel() {
    if (!channelDraft.channelName.trim()) return
    const nextId = Math.max(0, ...channels.map((item) => item.id)) + 1
    setChannels((current) => [
      ...current,
      {
        id: nextId,
        channelName: channelDraft.channelName.trim(),
        phone: channelDraft.phone.trim() || 'Numero sintetico',
        provider: channelDraft.provider,
        quality: 'Boa',
        templates: [],
      },
    ])
    setActiveChannelId(nextId)
    setShowChannelModal(false)
    setChannelDraft({ channelName: '', phone: '', provider: 'Conexao interna' })
  }

  function createTemplate() {
    if (!templateDraft.title.trim() || !activeChannel) return
    const nextId = Math.max(0, ...(activeChannel.templates.map((item) => item.id))) + 1
    setChannels((current) =>
      current.map((channel) =>
        channel.id === activeChannel.id
          ? {
              ...channel,
              templates: [
                {
                  id: nextId,
                  title: templateDraft.title.trim(),
                  category: templateDraft.category,
                  status: 'Rascunho',
                  language: templateDraft.language,
                },
                ...channel.templates,
              ],
            }
          : channel,
      ),
    )
    setShowTemplateModal(false)
    setTemplateDraft({ title: '', category: 'Marketing', language: 'pt-BR' })
  }

  return (
    <section className="content-page">
      <div className="page-heading compact">
        <button className="icon-button subtle back-button" onClick={() => navigate('/settings')}>{'<'}</button>
        <div>
          <h1>Templates WhatsApp Business API</h1>
          <p>E preciso ter um modelo de mensagem para comecar uma conversa iniciada pela empresa.</p>
        </div>
      </div>

      {channels.length === 0 ? (
        <div className="card empty-state">
          <button className="primary-button compact top-right" onClick={() => setShowChannelModal(true)}>Criar canal de atendimento</button>
          <div className="empty-illustration">
            <div className="paper-stack" />
            <div className="plus-disc">+</div>
          </div>
          <strong>Nenhum canal de atendimento WhatsApp API encontrado!</strong>
        </div>
      ) : (
        <div className="card table-card">
          <div className="table-toolbar">
            <div className="search-field">
              <Glyph name="chat" />
              <input type="text" value={activeChannel?.channelName ?? ''} readOnly />
            </div>
            <button className="select-button">Qualidade: {activeChannel?.quality ?? 'Boa'}</button>
            <button className="select-button">Fornecedor: {activeChannel?.provider ?? 'Conexao interna'}</button>
            <div className="toolbar-spacer" />
            <button className="select-button" onClick={() => setShowChannelModal(true)}>Novo canal</button>
            <button className="primary-button compact" onClick={() => setShowTemplateModal(true)}>Novo template</button>
          </div>

          <div className="channel-switcher">
            {channels.map((channel) => (
              <button key={channel.id} className={`channel-pill ${activeChannelId === channel.id ? 'active' : ''}`} onClick={() => setActiveChannelId(channel.id)}>
                <strong>{channel.channelName}</strong>
                <span>{channel.phone}</span>
              </button>
            ))}
          </div>

          <div className="table-head campaigns templates-grid">
            <span>Template</span>
            <span>Categoria</span>
            <span>Status</span>
          </div>

          {activeChannel?.templates.map((template) => (
            <div className="table-row campaigns templates-grid" key={template.id}>
              <div>
                <strong>{template.title}</strong>
                <span>{template.language}</span>
              </div>
              <span>{template.category}</span>
              <button className="status-tag interactive-tag" onClick={() =>
                setChannels((current) =>
                  current.map((channel) =>
                    channel.id === activeChannel.id
                      ? {
                          ...channel,
                          templates: channel.templates.map((item) =>
                            item.id === template.id
                              ? { ...item, status: item.status === 'Aprovado' ? 'Em revisao' : item.status === 'Em revisao' ? 'Rascunho' : 'Aprovado' }
                              : item,
                          ),
                        }
                      : channel,
                  ),
                )
              }>
                {template.status}
              </button>
            </div>
          ))}
        </div>
      )}

      {showChannelModal && (
        <div className="modal-scrim" onClick={() => setShowChannelModal(false)}>
          <div className="contact-picker-modal form-modal" onClick={(event) => event.stopPropagation()}>
            <div className="side-panel-header">
              <div>
                <h2>Novo canal de atendimento</h2>
                <p>Cadastro sintetico para validar a experiencia de templates.</p>
              </div>
              <button className="icon-button small" onClick={() => setShowChannelModal(false)}><Glyph name="x" /></button>
            </div>
            <div className="modal-form-grid">
              <label className="form-field mini">
                <span>Nome do canal</span>
                <input type="text" value={channelDraft.channelName} onChange={(event) => setChannelDraft((current) => ({ ...current, channelName: event.target.value }))} />
              </label>
              <label className="form-field mini">
                <span>Numero</span>
                <input type="text" value={channelDraft.phone} onChange={(event) => setChannelDraft((current) => ({ ...current, phone: event.target.value }))} />
              </label>
              <label className="form-field mini">
                <span>Fornecedor</span>
                <input type="text" value={channelDraft.provider} onChange={(event) => setChannelDraft((current) => ({ ...current, provider: event.target.value }))} />
              </label>
            </div>
            <div className="side-panel-footer">
              <button className="primary-button compact" onClick={createChannel}>Salvar canal</button>
            </div>
          </div>
        </div>
      )}

      {showTemplateModal && (
        <div className="modal-scrim" onClick={() => setShowTemplateModal(false)}>
          <div className="contact-picker-modal form-modal" onClick={(event) => event.stopPropagation()}>
            <div className="side-panel-header">
              <div>
                <h2>Novo template</h2>
                <p>Estrutura basica para montar os rascunhos de mensagem do sistema.</p>
              </div>
              <button className="icon-button small" onClick={() => setShowTemplateModal(false)}><Glyph name="x" /></button>
            </div>
            <div className="modal-form-grid">
              <label className="form-field mini">
                <span>Titulo</span>
                <input type="text" value={templateDraft.title} onChange={(event) => setTemplateDraft((current) => ({ ...current, title: event.target.value }))} />
              </label>
              <label className="form-field mini">
                <span>Categoria</span>
                <input type="text" value={templateDraft.category} onChange={(event) => setTemplateDraft((current) => ({ ...current, category: event.target.value }))} />
              </label>
              <label className="form-field mini">
                <span>Idioma</span>
                <input type="text" value={templateDraft.language} onChange={(event) => setTemplateDraft((current) => ({ ...current, language: event.target.value }))} />
              </label>
            </div>
            <div className="side-panel-footer">
              <button className="primary-button compact" onClick={createTemplate}>Salvar template</button>
            </div>
          </div>
        </div>
      )}
    </section>
  )
}

function DashboardPage() {
  return (
    <section className="content-page">
      <div className="page-heading">
        <h1>Relatorios</h1>
        <p>Uma visao consolidada dos principais indicadores operacionais do atendimento.</p>
      </div>
      <div className="dashboard-grid">
        {dashboardCards.map((card) => (
          <article key={card.title} className="card metric-card">
            <span>{card.title}</span>
            <strong>{card.value}</strong>
            <small>{card.delta}</small>
          </article>
        ))}
      </div>
      <div className="dashboard-panels">
        <div className="card chart-card">
          <h2>Volume por canal</h2>
          <div className="chart-bars">
            <span style={{ height: '46%' }} />
            <span style={{ height: '70%' }} />
            <span style={{ height: '58%' }} />
            <span style={{ height: '82%' }} />
            <span style={{ height: '65%' }} />
            <span style={{ height: '91%' }} />
          </div>
        </div>
        <div className="card summary-card">
          <h2>Destaques</h2>
          <ul>
            <li>Equipe comercial com maior crescimento de leads no periodo.</li>
            <li>Tempo medio de resposta melhorou em relacao a semana passada.</li>
            <li>Chatbots responderam 31% das abordagens iniciais.</li>
          </ul>
        </div>
      </div>
    </section>
  )
}

function ChatbotsPage() {
  return (
    <section className="content-page">
      <div className="page-heading">
        <h1>Chatbots</h1>
        <p>Automatizacoes prontas para orientar, qualificar e encaminhar conversas.</p>
      </div>
      <div className="dashboard-grid chatbot-grid">
        {chatbotCards.map((item) => (
          <article key={item.title} className="card chatbot-card">
            <span className="chip">Ativo</span>
            <h2>{item.title}</h2>
            <p>{item.subtitle}</p>
            <button className="select-button">Editar fluxo</button>
          </article>
        ))}
      </div>
    </section>
  )
}

function CampaignsPage() {
  const [campaigns, setCampaigns] = useState(buildCampaignState)
  const [searchTerm, setSearchTerm] = useState('')
  const [showCampaignModal, setShowCampaignModal] = useState(false)
  const [draftCampaign, setDraftCampaign] = useState({
    title: '',
    audience: '',
    channel: 'WhatsApp',
    template: 'Boas-vindas operacionais',
  })

  const filteredCampaigns = campaigns.filter((item) => `${item.title} ${item.audience} ${item.channel}`.toLowerCase().includes(searchTerm.toLowerCase()))

  function createCampaign() {
    if (!draftCampaign.title.trim()) return
    const nextId = Math.max(0, ...campaigns.map((item) => item.id)) + 1
    setCampaigns((current) => [
      {
        id: nextId,
        title: draftCampaign.title.trim(),
        audience: draftCampaign.audience.trim() || 'Base interna',
        status: 'Rascunho',
        channel: draftCampaign.channel,
        template: draftCampaign.template,
        selected: false,
      },
      ...current,
    ])
    setShowCampaignModal(false)
    setDraftCampaign({ title: '', audience: '', channel: 'WhatsApp', template: 'Boas-vindas operacionais' })
  }

  return (
    <section className="content-page">
      <div className="page-heading">
        <h1>Envio de campanhas de marketing</h1>
        <p>Organize campanhas por publico, acompanhe status e dispare comunicacoes em lote.</p>
      </div>

      <div className="card table-card">
        <div className="table-toolbar">
          <div className="search-field">
            <Glyph name="search" />
            <input type="text" placeholder="Pesquisar campanhas" value={searchTerm} onChange={(event) => setSearchTerm(event.target.value)} />
          </div>
          <button className="select-button">Campanhas: {campaigns.length}</button>
          <div className="toolbar-spacer" />
          <button className="primary-button compact" onClick={() => setShowCampaignModal(true)}>Nova campanha</button>
        </div>
        <div className="table-head campaigns campaigns-grid">
          <span>Campanha</span>
          <span>Publico</span>
          <span>Status</span>
          <span>Canal</span>
        </div>
        {filteredCampaigns.map((row) => (
          <div className="table-row campaigns campaigns-grid" key={row.id}>
            <div>
              <strong>{row.title}</strong>
              <span>{row.template}</span>
            </div>
            <span>{row.audience}</span>
            <button className="status-tag interactive-tag" onClick={() => setCampaigns((current) => current.map((item) => item.id === row.id ? { ...item, status: item.status === 'Programada' ? 'Concluida' : item.status === 'Concluida' ? 'Rascunho' : 'Programada' } : item))}>
              {row.status}
            </button>
            <span>{row.channel}</span>
          </div>
        ))}

        {filteredCampaigns.length === 0 && (
          <div className="empty-inline-state in-card">
            <strong>Nenhuma campanha encontrada</strong>
            <span>Crie uma nova campanha sintetica para continuar os testes operacionais.</span>
          </div>
        )}
      </div>

      {showCampaignModal && (
        <div className="modal-scrim" onClick={() => setShowCampaignModal(false)}>
          <div className="contact-picker-modal form-modal" onClick={(event) => event.stopPropagation()}>
            <div className="side-panel-header">
              <div>
                <h2>Nova campanha</h2>
                <p>Fluxo sintetico para montagem das campanhas internas.</p>
              </div>
              <button className="icon-button small" onClick={() => setShowCampaignModal(false)}><Glyph name="x" /></button>
            </div>
            <div className="modal-form-grid">
              <label className="form-field mini">
                <span>Titulo</span>
                <input type="text" value={draftCampaign.title} onChange={(event) => setDraftCampaign((current) => ({ ...current, title: event.target.value }))} />
              </label>
              <label className="form-field mini">
                <span>Publico</span>
                <input type="text" value={draftCampaign.audience} onChange={(event) => setDraftCampaign((current) => ({ ...current, audience: event.target.value }))} />
              </label>
              <label className="form-field mini">
                <span>Canal</span>
                <input type="text" value={draftCampaign.channel} onChange={(event) => setDraftCampaign((current) => ({ ...current, channel: event.target.value }))} />
              </label>
              <label className="form-field mini">
                <span>Template base</span>
                <input type="text" value={draftCampaign.template} onChange={(event) => setDraftCampaign((current) => ({ ...current, template: event.target.value }))} />
              </label>
            </div>
            <div className="side-panel-footer">
              <button className="primary-button compact" onClick={createCampaign}>Salvar campanha</button>
            </div>
          </div>
        </div>
      )}
    </section>
  )
}

export default App
