import { useEffect, useRef, useState } from 'react'
import {
  assumeConversationRequest,
  finalizeConversationRequest,
  incomingMessageRequest,
  listContactsRequest,
  listConversationMessagesRequest,
  listConversationsRequest,
  listDepartmentsRequest,
  sendConversationMessageRequest,
  transferConversationRequest,
  waitCustomerConversationRequest,
} from '../../lib/api'
import {
  buildMediaLibraryState,
  operatorOptions,
  sortActionItems,
} from '../../lib/appData'
import { ContactPickerModal } from './conversations/ContactPickerModal'
import { ConversationListPane } from './conversations/ConversationListPane'
import { ConversationProvider } from './conversations/ConversationContext'
import { ConversationStagePane } from './conversations/ConversationStagePane'

const EMPTY_CONVERSATION = {
  id: null,
  name: 'Nenhuma conversa',
  preview: 'Selecione uma conversa para continuar.',
  time: '--:--',
  unread: 0,
  category: 'Geral',
  color: '#6f87d9',
  status: 'entrada',
  assignedTo: 'Fila',
  priority: 'normal',
  selected: false,
  note: 'Sem observacoes.',
  scheduledFor: null,
  sector: 'Geral',
  blocked: false,
  pinned: false,
  splitView: false,
}

const DEFAULT_NOTICE = 'Selecione uma acao para operar o atendimento em tempo real.'
const DEFAULT_NOTE = 'Registro interno da conversa para orientacao da equipe.'

const UI_STATUS_BY_BACKEND = {
  bot_triage: 'entrada',
  queued: 'entrada',
  in_progress: 'entrada',
  waiting_customer: 'esperando',
  finalized: 'finalizados',
}

const COLOR_PALETTE = ['#ffcb75', '#3da3ff', '#8baf5a', '#dc8c7c', '#4c78ff', '#a768da', '#3ab7a1']

function nowTimeLabel() {
  return new Date().toLocaleTimeString('pt-BR', { hour: '2-digit', minute: '2-digit' })
}

function toTimeLabel(value) {
  if (!value) return '--:--'
  const date = new Date(value)
  if (Number.isNaN(date.getTime())) return '--:--'
  return date.toLocaleTimeString('pt-BR', { hour: '2-digit', minute: '2-digit' })
}

function toTimestamp(value) {
  const timestamp = new Date(value).getTime()
  return Number.isNaN(timestamp) ? 0 : timestamp
}

function colorFromText(text) {
  const input = String(text || '')
  let hash = 0
  for (let index = 0; index < input.length; index += 1) {
    hash = (hash << 5) - hash + input.charCodeAt(index)
    hash |= 0
  }
  return COLOR_PALETTE[Math.abs(hash) % COLOR_PALETTE.length]
}

function resolveAssignedTo(row, currentUserName) {
  if (row.status === 'finalized') return 'Encerrada'
  if (row.isMine) return currentUserName || operatorOptions[0]
  if (row.activeParticipants > 0) return `${row.activeParticipants} atendente(s)`
  if (row.status === 'waiting_customer') return 'Esperando cliente'
  if (row.status === 'bot_triage') return 'Fluxo automatizado'
  return row.departmentName ? `Fila ${row.departmentName}` : 'Fila geral'
}

function mapConversationRowToUi(row, previous, currentUserName) {
  const uiStatus = UI_STATUS_BY_BACKEND[row.status] ?? 'entrada'
  const customCategory = previous?.customCategory ?? null
  const category = customCategory || row.departmentName || (row.status === 'bot_triage' ? 'Bot' : 'Geral')

  return {
    id: row.id,
    backendStatus: row.status,
    name: row.contactName || 'Contato',
    preview: row.lastMessage || 'Sem mensagens nesta conversa.',
    time: toTimeLabel(row.updatedAt),
    unread: previous?.unread ?? 0,
    category,
    customCategory,
    color: previous?.color || colorFromText(row.channelIdentifier || row.id),
    status: uiStatus,
    assignedTo: resolveAssignedTo(row, currentUserName),
    priority: row.status === 'queued' ? 'media' : 'normal',
    selected: previous?.selected ?? false,
    note: previous?.note ?? DEFAULT_NOTE,
    scheduledFor: previous?.scheduledFor ?? null,
    sector: row.departmentName || 'Geral',
    blocked: previous?.blocked ?? false,
    pinned: previous?.pinned ?? false,
    splitView: previous?.splitView ?? false,
    updatedAtTs: toTimestamp(row.updatedAt),
    openedAtTs: toTimestamp(row.openedAt),
    channelIdentifier: row.channelIdentifier,
    departmentId: row.departmentId,
    departmentName: row.departmentName,
  }
}

function mapMessagesToThread(messages, currentUserName) {
  return messages.map((message) => {
    const side =
      message.senderType === 'operator'
        ? 'right'
        : message.senderType === 'system'
          ? 'system'
          : 'left'
    const author =
      message.senderType === 'operator'
        ? currentUserName || 'Atendente'
        : message.senderType === 'bot'
          ? 'Chatbot'
          : null

    return {
      id: message.id,
      side,
      author,
      body: message.body,
      time: toTimeLabel(message.createdAt),
      senderType: message.senderType,
    }
  })
}

function mapMessagesToTimeline(messages) {
  const timeline = messages
    .filter((item) => item.senderType === 'system' || item.senderType === 'bot')
    .map((item) => ({
      time: toTimeLabel(item.createdAt),
      text: item.body,
    }))
    .reverse()

  return timeline
}

export function ConversationsPage({ session }) {
  const token = session?.accessToken
  const currentUserName = session?.user?.name || operatorOptions[0]
  const organizationId = session?.user?.organizationId

  const [conversations, setConversations] = useState([])
  const [selectedConversationId, setSelectedConversationId] = useState(null)
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
  const [contactNoteDraft, setContactNoteDraft] = useState(DEFAULT_NOTE)
  const [transferTarget, setTransferTarget] = useState('')
  const [scheduleDraft, setScheduleDraft] = useState({ date: '15/04/2026', time: '09:00' })
  const [messageThreads, setMessageThreads] = useState({})
  const [activityTimeline, setActivityTimeline] = useState({})
  const [mediaLibrary, setMediaLibrary] = useState(buildMediaLibraryState)
  const [notice, setNotice] = useState(DEFAULT_NOTICE)
  const [activeQueueMenuId, setActiveQueueMenuId] = useState(null)
  const [activeMessageMenuId, setActiveMessageMenuId] = useState(null)
  const [attachmentTab, setAttachmentTab] = useState('new')
  const [mediaSearch, setMediaSearch] = useState('')
  const [composerSubscribed, setComposerSubscribed] = useState(true)
  const [contacts, setContacts] = useState([])
  const [departments, setDepartments] = useState([])
  const [isLoadingConversations, setIsLoadingConversations] = useState(true)
  const [isLoadingMessages, setIsLoadingMessages] = useState(false)
  const [isSending, setIsSending] = useState(false)
  const fileInputRef = useRef(null)

  const selectedConversation =
    conversations.find((item) => item.id === selectedConversationId) ?? EMPTY_CONVERSATION
  const selectedThread = selectedConversationId ? messageThreads[selectedConversationId] ?? [] : []
  const visibleMedia = (mediaLibrary[selectedConversationId] ?? []).filter((item) =>
    item.name.toLowerCase().includes(mediaSearch.toLowerCase()),
  )
  const transferOptions = departments.map((department) => ({
    id: department.id,
    name: department.name,
  }))

  const visibleConversations = [...conversations]
    .filter((item) => item.status === activeTab)
    .filter((item) => `${item.name} ${item.preview}`.toLowerCase().includes(searchTerm.toLowerCase()))
    .sort((left, right) => {
      if (left.pinned !== right.pinned) return Number(right.pinned) - Number(left.pinned)
      if (sortMode === 'Ultima mensagem') return right.updatedAtTs - left.updatedAtTs
      if (sortMode === 'Esperando resposta') return Number(right.unread > 0) - Number(left.unread > 0)
      return right.openedAtTs - left.openedAtTs
    })

  useEffect(() => {
    if (!token) return
    let cancelled = false

    async function loadBaseData({ silent = false } = {}) {
      if (!silent) {
        setIsLoadingConversations(true)
      }

      try {
        const [conversationResult, departmentResult, contactResult] = await Promise.all([
          listConversationsRequest(token),
          listDepartmentsRequest(token),
          listContactsRequest(token),
        ])

        if (cancelled) return

        setDepartments(departmentResult?.departments ?? [])
        setContacts(contactResult?.contacts ?? [])

        const rows = conversationResult?.conversations ?? []
        setConversations((current) => {
          const currentMap = new Map(current.map((item) => [item.id, item]))
          return rows.map((row) =>
            mapConversationRowToUi(row, currentMap.get(row.id), currentUserName),
          )
        })

        if (!transferTarget && (departmentResult?.departments?.length ?? 0) > 0) {
          setTransferTarget(departmentResult.departments[0].id)
        }
      } catch (error) {
        if (!cancelled) {
          setNotice(`Nao foi possivel carregar conversas: ${error.message}`)
        }
      } finally {
        if (!silent && !cancelled) {
          setIsLoadingConversations(false)
        }
      }
    }

    loadBaseData()
    const polling = window.setInterval(() => {
      loadBaseData({ silent: true })
    }, 12000)

    return () => {
      cancelled = true
      window.clearInterval(polling)
    }
  }, [token, currentUserName])

  useEffect(() => {
    if (conversations.length === 0) {
      setSelectedConversationId(null)
      return
    }
    const selectedStillExists = conversations.some((item) => item.id === selectedConversationId)
    if (!selectedStillExists) {
      setSelectedConversationId(conversations[0].id)
    }
  }, [conversations, selectedConversationId])

  useEffect(() => {
    if (!selectedConversationId || !token) return
    let cancelled = false

    async function loadMessages() {
      setIsLoadingMessages(true)
      try {
        const result = await listConversationMessagesRequest(token, selectedConversationId)
        if (cancelled) return
        const messages = result?.messages ?? []
        const thread = mapMessagesToThread(messages, currentUserName)
        const timeline = mapMessagesToTimeline(messages)
        setMessageThreads((current) => ({ ...current, [selectedConversationId]: thread }))
        setActivityTimeline((current) => ({ ...current, [selectedConversationId]: timeline }))
      } catch (error) {
        if (!cancelled) {
          setNotice(`Nao foi possivel carregar mensagens: ${error.message}`)
        }
      } finally {
        if (!cancelled) {
          setIsLoadingMessages(false)
        }
      }
    }

    loadMessages()
    return () => {
      cancelled = true
    }
  }, [selectedConversationId, token, currentUserName])

  useEffect(() => {
    if (selectedConversation && selectedConversation.id) {
      setContactNoteDraft(selectedConversation.note)
    }
  }, [selectedConversationId, selectedConversation])

  useEffect(() => {
    if (selectedConversation && selectedConversation.status === activeTab) return
    const nextConversation = conversations.find((item) => item.status === activeTab) ?? conversations[0]
    if (nextConversation && nextConversation.id !== selectedConversationId) {
      setSelectedConversationId(nextConversation.id)
    }
  }, [activeTab, conversations, selectedConversation, selectedConversationId])

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
    setConversations((current) =>
      current.map((item) => (item.id === targetId ? updater(item) : item)),
    )
  }

  function appendTimelineEntry(targetId, text) {
    const time = nowTimeLabel()
    setActivityTimeline((current) => ({
      ...current,
      [targetId]: [{ time, text }, ...(current[targetId] ?? [])],
    }))
  }

  async function reloadConversationsAndCurrentMessages(targetConversationId = selectedConversationId) {
    if (!token) return
    try {
      const [conversationResult, contactResult, departmentResult] = await Promise.all([
        listConversationsRequest(token),
        listContactsRequest(token),
        listDepartmentsRequest(token),
      ])
      setContacts(contactResult?.contacts ?? [])
      setDepartments(departmentResult?.departments ?? [])
      const rows = conversationResult?.conversations ?? []
      setConversations((current) => {
        const currentMap = new Map(current.map((item) => [item.id, item]))
        return rows.map((row) => mapConversationRowToUi(row, currentMap.get(row.id), currentUserName))
      })
      if (targetConversationId) {
        const result = await listConversationMessagesRequest(token, targetConversationId)
        const messages = result?.messages ?? []
        setMessageThreads((current) => ({
          ...current,
          [targetConversationId]: mapMessagesToThread(messages, currentUserName),
        }))
        setActivityTimeline((current) => ({
          ...current,
          [targetConversationId]: mapMessagesToTimeline(messages),
        }))
      }
    } catch (error) {
      setNotice(`Nao foi possivel atualizar os dados: ${error.message}`)
    }
  }

  function handleOpenConversation(nextConversationId) {
    setSelectedConversationId(nextConversationId)
    closeFloatingLayers()
    setActivePanel(null)
    patchConversation(nextConversationId, (item) => ({ ...item, unread: 0 }))
  }

  function handleFileSelection(event) {
    const files = Array.from(event.target.files ?? [])
    if (files.length === 0 || !selectedConversationId) return
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

  function handleQuickReplyInsert(text) {
    setComposerMode('message')
    setComposerText(text)
    setActivePanel(null)
    setNotice('Resposta rapida carregada no composer para revisao.')
  }

  function handleChatbotRun(title) {
    if (!selectedConversationId) return
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
    if (!selectedConversationId) return
    patchConversation(selectedConversationId, (item) => ({
      ...item,
      scheduledFor: `${scheduleDraft.date} ${scheduleDraft.time}`,
    }))
    appendTimelineEntry(
      selectedConversationId,
      `Mensagem agendada para ${scheduleDraft.date} as ${scheduleDraft.time}.`,
    )
    setNotice(`Agendamento salvo para ${scheduleDraft.date} as ${scheduleDraft.time}.`)
    setActivePanel(null)
  }

  async function handleTransfer() {
    if (!selectedConversationId || !transferTarget) {
      setNotice('Selecione um departamento de destino para transferir.')
      return
    }
    try {
      await transferConversationRequest(token, selectedConversationId, transferTarget)
      await reloadConversationsAndCurrentMessages(selectedConversationId)
      appendTimelineEntry(selectedConversationId, 'Conversa transferida para outra fila.')
      setNotice('Conversa transferida com sucesso.')
      setActiveTab('esperando')
      setActivePanel(null)
    } catch (error) {
      setNotice(`Falha ao transferir conversa: ${error.message}`)
    }
  }

  async function handleFinalizeSingle() {
    if (!selectedConversationId) return
    try {
      await finalizeConversationRequest(token, selectedConversationId)
      await reloadConversationsAndCurrentMessages(selectedConversationId)
      setNotice('Conversa atual marcada como finalizada.')
      setActiveTab('finalizados')
    } catch (error) {
      setNotice(`Falha ao finalizar conversa: ${error.message}`)
    }
  }

  async function handleQueueAction(conversationId, actionLabel) {
    try {
      if (actionLabel === 'Atribuir para mim') {
        await assumeConversationRequest(token, conversationId)
        await reloadConversationsAndCurrentMessages(conversationId)
        appendTimelineEntry(conversationId, 'Conversa atribuida ao operador atual.')
      } else if (actionLabel === 'Adicionar etiqueta') {
        patchConversation(conversationId, (item) => ({
          ...item,
          customCategory: item.customCategory === 'Vip' ? null : 'Vip',
          category: item.customCategory === 'Vip' ? item.sector : 'Vip',
        }))
      } else if (actionLabel === 'Marcar como nao lida') {
        patchConversation(conversationId, (item) => ({ ...item, unread: Math.max(1, item.unread || 1) }))
      } else if (actionLabel === 'Bloquear contato') {
        patchConversation(conversationId, (item) => ({ ...item, blocked: !item.blocked }))
        appendTimelineEntry(conversationId, 'Contato alternou entre bloqueado e desbloqueado.')
      } else if (actionLabel === 'Abrir lado-a-lado') {
        patchConversation(conversationId, (item) => ({ ...item, splitView: !item.splitView }))
      } else if (actionLabel === 'Finalizar conversa') {
        await finalizeConversationRequest(token, conversationId)
        await reloadConversationsAndCurrentMessages(conversationId)
        setActiveTab('finalizados')
      } else if (actionLabel === 'Marcar como esperando') {
        await waitCustomerConversationRequest(token, conversationId)
        await reloadConversationsAndCurrentMessages(conversationId)
        setActiveTab('esperando')
      } else if (actionLabel === 'Fixar conversa') {
        patchConversation(conversationId, (item) => ({ ...item, pinned: !item.pinned }))
      }
      setNotice(`Acao aplicada: ${actionLabel.toLowerCase()}.`)
    } catch (error) {
      setNotice(`Falha na acao "${actionLabel}": ${error.message}`)
    } finally {
      setActiveQueueMenuId(null)
    }
  }

  async function handleMessageAction(message, actionLabel) {
    if (!selectedConversationId) return
    if (actionLabel === 'Responder') {
      setComposerText(`@resposta ${message.body}`)
      setNotice('Mensagem carregada no composer para resposta contextual.')
    }
    if (actionLabel === 'Copiar') {
      try {
        await navigator.clipboard.writeText(message.body)
        setNotice('Conteudo da mensagem copiado.')
      } catch {
        setNotice('Nao foi possivel copiar automaticamente.')
      }
    }
    if (actionLabel === 'Encaminhar') {
      setComposerText(`[Encaminhado] ${message.body}`)
      setNotice('Mensagem preparada para encaminhamento.')
    }
    if (actionLabel === 'Apagar') {
      setMessageThreads((current) => ({
        ...current,
        [selectedConversationId]: (current[selectedConversationId] ?? []).filter((item) => item.id !== message.id),
      }))
      setNotice('Mensagem removida da visualizacao local.')
    }
    setActiveMessageMenuId(null)
  }

  async function handleBulkAction(actionLabel) {
    const selectedIds = conversations.filter((item) => item.selected && item.status === activeTab).map((item) => item.id)

    if (actionLabel === 'Selecionar todas as conversas') {
      setConversations((current) =>
        current.map((item) => (item.status === activeTab ? { ...item, selected: true } : item)),
      )
      setNotice(`Todas as conversas da aba ${activeTab} foram selecionadas.`)
      setListPopover(null)
      return
    }

    if (selectedIds.length === 0) {
      setNotice('Selecione pelo menos uma conversa antes de usar acoes em massa.')
      setListPopover(null)
      return
    }

    try {
      if (actionLabel === 'Distribuir para um atendente') {
        await Promise.all(selectedIds.map((id) => assumeConversationRequest(token, id)))
        setNotice(`${selectedIds.length} conversa(s) distribuidas para o operador atual.`)
      }

      if (actionLabel === 'Transferir para outro setor') {
        const supportDepartment =
          departments.find((item) => item.name.toLowerCase().includes('suporte')) ?? departments[0]
        if (!supportDepartment) {
          throw new Error('Nenhum departamento ativo disponivel para transferencia.')
        }
        await Promise.all(
          selectedIds.map((id) => transferConversationRequest(token, id, supportDepartment.id)),
        )
        setNotice(`${selectedIds.length} conversa(s) transferidas para a fila ${supportDepartment.name}.`)
        setActiveTab('esperando')
      }

      if (actionLabel === 'Finalizar selecionadas') {
        await Promise.all(selectedIds.map((id) => finalizeConversationRequest(token, id)))
        setNotice(`${selectedIds.length} conversa(s) finalizadas.`)
        setActiveTab('finalizados')
      }

      await reloadConversationsAndCurrentMessages(selectedConversationId)
      setConversations((current) =>
        current.map((item) => (selectedIds.includes(item.id) ? { ...item, selected: false } : item)),
      )
    } catch (error) {
      setNotice(`Falha em acoes em massa: ${error.message}`)
    }

    setListPopover(null)
  }

  async function handleSendComposer() {
    if (!selectedConversationId) return

    if (composerMode === 'message') {
      const trimmed = composerText.trim()
      if (!trimmed) return
      setIsSending(true)
      try {
        await sendConversationMessageRequest(token, selectedConversationId, trimmed)
        setComposerText('')
        setNotice('Mensagem enviada na conversa atual.')
        appendTimelineEntry(selectedConversationId, 'Mensagem enviada manualmente pelo operador.')
        await reloadConversationsAndCurrentMessages(selectedConversationId)
      } catch (error) {
        setNotice(`Falha ao enviar mensagem: ${error.message}`)
      } finally {
        setIsSending(false)
      }
      return
    }

    const trimmed = composerText.trim()
    if (!trimmed) return
    patchConversation(selectedConversationId, (item) => ({ ...item, note: trimmed }))
    setContactNoteDraft(trimmed)
    appendTimelineEntry(selectedConversationId, `Nota interna adicionada: ${trimmed}`)
    setNotice('Nota interna registrada para a equipe.')
    setComposerText('')
  }

  function toggleConversationSelected(conversationId, selected) {
    setConversations((current) =>
      current.map((item) => (item.id === conversationId ? { ...item, selected } : item)),
    )
  }

  async function handleOpenContact(contact) {
    if (!contact) return
    setShowContactPicker(false)

    const existing = conversations.find((item) => item.channelIdentifier === contact.channelIdentifier)
    if (existing) {
      setActiveTab(existing.status)
      handleOpenConversation(existing.id)
      setNotice('Contato aberto na inbox.')
      return
    }

    try {
      if (!organizationId) {
        throw new Error('Organizacao invalida para iniciar uma conversa.')
      }

      const incoming = await incomingMessageRequest({
        organizationId,
        channelIdentifier: contact.channelIdentifier,
        contactName: contact.displayName,
        phoneMasked: contact.phoneMasked || null,
        text: '1',
      })
      await reloadConversationsAndCurrentMessages(incoming.conversationId)
      setSelectedConversationId(incoming.conversationId)
      setActiveTab('entrada')
      setNotice('Nova conversa iniciada a partir do contato selecionado.')
    } catch (error) {
      setNotice(`Nao foi possivel abrir contato sem conversa existente: ${error.message}`)
    }
  }

  const contextValue = {
    activeMessageMenuId,
    activePanel,
    activeQueueMenuId,
    activeTab,
    activityTimeline,
    attachmentTab,
    composerMode,
    composerSubscribed,
    composerText,
    contactNoteDraft,
    conversations,
    fileInputRef,
    handleBulkAction,
    handleChatbotRun,
    handleFileSelection,
    handleFinalizeSingle,
    handleMessageAction,
    handleOpenConversation,
    handleQueueAction,
    handleQuickReplyInsert,
    handleSaveSchedule,
    handleSendComposer,
    handleTransfer,
    isLoadingConversations,
    isLoadingMessages,
    isSending,
    listPopover,
    mediaSearch,
    notice,
    openPanel,
    patchConversation,
    scheduleDraft,
    searchTerm,
    selectedConversation,
    selectedConversationId,
    selectedThread,
    setActiveMessageMenuId,
    setActivePanel,
    setActiveQueueMenuId,
    setActiveTab,
    setAttachmentTab,
    setComposerMode,
    setComposerSubscribed,
    setComposerText,
    setContactNoteDraft,
    setListPopover,
    setMediaSearch,
    setNotice,
    setScheduleDraft,
    setSearchTerm,
    setShowAttachments,
    setShowEmojiPicker,
    setShowStickerDrawer,
    setSortMode,
    setTransferTarget,
    showAttachments,
    showEmojiPicker,
    showStickerDrawer,
    sortMode,
    toggleConversationSelected,
    transferOptions,
    transferTarget,
    visibleConversations,
    visibleMedia,
  }

  return (
    <ConversationProvider value={contextValue}>
      <section className="conversation-screen">
        <ConversationListPane />
        <ConversationStagePane />
      </section>
      {showContactPicker && (
        <ContactPickerModal
          contacts={contacts}
          onClose={() => setShowContactPicker(false)}
          onOpenContact={handleOpenContact}
        />
      )}
    </ConversationProvider>
  )
}
