import {
  chatbotPlaybooks,
  contactInsightItems,
  conversationDetailItems,
  quickReplyGroups,
  schedulerBlocks,
  stickerItems,
} from '../../../lib/appData'
import { Glyph } from '../../shared/Glyph'
import { useConversationContext } from './ConversationContext'
import { ConversationSidePanel } from './ConversationSidePanel'

export function ConversationStagePane() {
  const {
    activeMessageMenuId,
    activePanel,
    activityTimeline,
    attachmentTab,
    composerMode,
    composerSubscribed,
    composerText,
    contactNoteDraft,
    fileInputRef,
    handleChatbotRun,
    handleFileSelection,
    handleFinalizeSingle,
    handleMessageAction,
    handleQuickReplyInsert,
    handleSaveSchedule,
    handleSendComposer,
    handleTransfer,
    isLoadingMessages,
    isSending,
    mediaSearch,
    notice,
    openPanel,
    patchConversation,
    scheduleDraft,
    selectedConversation,
    selectedConversationId,
    selectedThread,
    setActiveMessageMenuId,
    setActivePanel,
    setAttachmentTab,
    setComposerMode,
    setComposerSubscribed,
    setComposerText,
    setContactNoteDraft,
    setMediaSearch,
    setNotice,
    setScheduleDraft,
    setShowAttachments,
    setShowEmojiPicker,
    setShowStickerDrawer,
    setTransferTarget,
    showAttachments,
    showEmojiPicker,
    showStickerDrawer,
    transferOptions,
    transferTarget,
    visibleMedia,
  } = useConversationContext()

  const hasConversation = Boolean(selectedConversation?.id)

  return (
    <div className="conversation-stage">
      <header className="chat-topbar">
        <div className="chat-person">
          <div className="avatar large" style={{ '--avatar': selectedConversation.color }}>
            {selectedConversation.name[0]}
          </div>
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
          <button
            className={`toolbar-chip ${activePanel === 'contact' ? 'active' : ''}`}
            onClick={() => openPanel('contact')}
          >
            Contato
          </button>
          <button
            className={`toolbar-chip ${activePanel === 'details' ? 'active' : ''}`}
            onClick={() => openPanel('details')}
          >
            Detalhes da conversa
          </button>
          <button
            className={`toolbar-chip ${activePanel === 'schedule' ? 'active' : ''}`}
            onClick={() => openPanel('schedule')}
          >
            Agendar envio
          </button>
          <button
            className={`icon-button small ${activePanel === 'transfer' ? 'is-active' : ''}`}
            title="Transferir conversa"
            onClick={() => openPanel('transfer')}
            disabled={!hasConversation}
          >
            <Glyph name="send" />
          </button>
          <button
            className="icon-button small"
            title="Finalizar conversa"
            onClick={handleFinalizeSingle}
            disabled={!hasConversation}
          >
            <Glyph name="check" />
          </button>
        </div>
      </header>

      <div className="chat-canvas">
        <span className="date-pill">Hoje</span>
        <div className="status-banner">
          <strong>Estado atual</strong>
          <span>{notice}</span>
        </div>

        {isLoadingMessages && (
          <div className="status-banner">
            <strong>Carregando mensagens</strong>
            <span>Aguarde enquanto sincronizamos a conversa selecionada.</span>
          </div>
        )}

        {selectedThread.map((message) => (
          <div key={message.id} className={`bubble ${message.side}`}>
            <div className="bubble-head">
              <div>{message.author && <strong>{message.author}</strong>}</div>
              <button
                className="icon-button tiny subtle"
                onClick={() =>
                  setActiveMessageMenuId((current) => (current === message.id ? null : message.id))
                }
              >
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
                <button
                  key={emoji}
                  className="emoji-cell"
                  onClick={() => {
                    setComposerText((current) => `${current} ${emoji}`.trim())
                    setNotice('Reacao inserida no composer para edicao.')
                  }}
                >
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
              <button className="icon-button small" onClick={() => setShowAttachments(false)}>
                <Glyph name="x" />
              </button>
            </div>
            <div className="attachment-tabs">
              <button
                className={`toolbar-chip ${attachmentTab === 'new' ? 'active' : ''}`}
                onClick={() => setAttachmentTab('new')}
              >
                Novo arquivo
              </button>
              <button
                className={`toolbar-chip ${attachmentTab === 'library' ? 'active' : ''}`}
                onClick={() => setAttachmentTab('library')}
              >
                Biblioteca de midia
              </button>
            </div>
            <div className="attachment-toolbar">
              <div className="search-field panel-search">
                <Glyph name="search" />
                <input
                  type="text"
                  placeholder="Pesquisar"
                  value={mediaSearch}
                  onChange={(event) => setMediaSearch(event.target.value)}
                />
              </div>
              <button className="select-button">Todos os arquivos</button>
            </div>
            {attachmentTab === 'new' ? (
              <div className="attachment-empty">
                <strong>Selecione um arquivo do seu computador</strong>
                <span>
                  Ao clicar abaixo, a janela do Windows Explorer sera aberta para escolher um ou mais
                  arquivos.
                </span>
                <button
                  className="primary-button compact"
                  onClick={() => fileInputRef.current?.click()}
                >
                  Escolher arquivo
                </button>
              </div>
            ) : visibleMedia.length > 0 ? (
              <div className="media-library-list">
                {visibleMedia.map((item) => (
                  <button
                    key={item.id}
                    className="media-row"
                    onClick={() => {
                      setComposerText((current) =>
                        `${current}\n[Midia selecionada: ${item.name}]`.trim(),
                      )
                      setNotice(`Midia "${item.name}" preparada no composer.`)
                      setShowAttachments(false)
                    }}
                  >
                    <div>
                      <strong>{item.name}</strong>
                      <span>
                        {item.type} - {item.size}
                      </span>
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
              <button className="icon-button small" onClick={() => setShowStickerDrawer(false)}>
                <Glyph name="x" />
              </button>
            </div>
            <div className="sticker-grid">
              {stickerItems.map((label) => (
                <button
                  key={label}
                  className="sticker-card"
                  onClick={() => {
                    setComposerText((current) => `${current}\n[Figurinha: ${label}]`.trim())
                    setNotice(`Figurinha sintetica "${label}" adicionada ao composer.`)
                    setShowStickerDrawer(false)
                  }}
                >
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
            footer={
              <button
                className="primary-button compact"
                onClick={() => {
                  patchConversation(selectedConversationId, (item) => ({
                    ...item,
                    note: contactNoteDraft,
                  }))
                  setNotice('Observacoes do contato atualizadas.')
                }}
              >
                Salvar observacoes
              </button>
            }
          >
            <div className="panel-profile">
              <div className="avatar medium" style={{ '--avatar': selectedConversation.color }}>
                {selectedConversation.name[0]}
              </div>
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
                    <strong>
                      {label === 'Setor sugerido'
                        ? selectedConversation.sector
                        : label === 'Ultimo operador'
                          ? selectedConversation.assignedTo
                          : value}
                    </strong>
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
                <textarea
                  rows={4}
                  value={contactNoteDraft}
                  onChange={(event) => setContactNoteDraft(event.target.value)}
                />
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
              <span className="chip">
                {selectedConversation.status === 'finalizados' ? 'Conversa encerrada' : 'Atendimento ativo'}
              </span>
              <span className="chip">
                {selectedConversation.priority === 'alta'
                  ? 'Prioridade alta'
                  : selectedConversation.priority === 'media'
                    ? 'Prioridade media'
                    : 'Operacao padrao'}
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
                            ? 'Aguardando resposta do cliente'
                            : 'Atendimento em andamento'
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
                    <button
                      key={item.shortcut}
                      className="reply-card"
                      onClick={() => handleQuickReplyInsert(item.body)}
                    >
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
                  <button className="select-button compact" onClick={() => handleChatbotRun(item.title)}>
                    Executar
                  </button>
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
            footer={
              <button
                className="primary-button compact"
                onClick={handleTransfer}
                disabled={!hasConversation || !transferTarget}
              >
                Confirmar transferencia
              </button>
            }
          >
            <div className="panel-section">
              <h3>Destino</h3>
              <div className="stack-list">
                {transferOptions.map((option) => (
                  <button
                    key={option.id}
                    className={`option-card ${transferTarget === option.id ? 'selected' : ''}`}
                    onClick={() => setTransferTarget(option.id)}
                  >
                    <strong>{option.name}</strong>
                    <span>
                      {option.name.toLowerCase() === 'geral'
                        ? 'Mantem o atendimento na fila principal.'
                        : `Direciona a conversa para o setor ${option.name.toLowerCase()}.`}
                    </span>
                  </button>
                ))}
                {transferOptions.length === 0 && (
                  <div className="status-banner">
                    <strong>Nenhum departamento disponivel</strong>
                    <span>Cadastre ao menos um departamento ativo para transferir.</span>
                  </div>
                )}
              </div>
            </div>
          </ConversationSidePanel>
        )}

        {activePanel === 'schedule' && (
          <ConversationSidePanel
            title="Agendamento de mensagens"
            subtitle="Escolha quando a proxima resposta deve ser enviada"
            onClose={() => setActivePanel(null)}
            footer={
              <button className="primary-button compact" onClick={handleSaveSchedule}>
                Salvar agendamento
              </button>
            }
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
                  <input
                    type="text"
                    value={scheduleDraft.date}
                    onChange={(event) =>
                      setScheduleDraft((current) => ({ ...current, date: event.target.value }))
                    }
                  />
                </label>
                <label className="form-field mini">
                  <span>Hora</span>
                  <input
                    type="text"
                    value={scheduleDraft.time}
                    onChange={(event) =>
                      setScheduleDraft((current) => ({ ...current, time: event.target.value }))
                    }
                  />
                </label>
              </div>
            </div>
            {schedulerBlocks.map((block) => (
              <div key={block.title} className="panel-section">
                <h3>{block.title}</h3>
                <div className="panel-chip-row">
                  {block.options.map((option) => (
                    <button
                      key={option}
                      className="chip interactive"
                      onClick={() => setScheduleDraft((current) => ({ ...current, time: option }))}
                    >
                      {option}
                    </button>
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
            <input
              type="checkbox"
              checked={composerSubscribed}
              onChange={(event) => setComposerSubscribed(event.target.checked)}
            />
            <span />
          </label>
          <strong>{composerSubscribed ? selectedConversation.assignedTo : 'Assinar'}</strong>
        </div>
        <div className="composer-tabs">
          <button
            className={`pill ${composerMode === 'message' ? 'active' : 'ghost'}`}
            onClick={() => setComposerMode('message')}
          >
            Mensagem
          </button>
          <button
            className={`pill ${composerMode === 'note' ? 'active' : 'ghost'}`}
            onClick={() => setComposerMode('note')}
          >
            Notas
          </button>
        </div>
        <textarea
          placeholder={
            composerMode === 'message'
              ? 'Digite sua mensagem ou arraste um arquivo...'
              : 'Registre uma observacao interna para a equipe...'
          }
          rows={3}
          value={composerText}
          onChange={(event) => setComposerText(event.target.value)}
        />
        <div className="composer-tools">
          <div className="tool-row">
            <button
              className={`icon-button tiny ${showAttachments ? 'is-active' : ''}`}
              onClick={() => {
                setShowAttachments((current) => !current)
                setAttachmentTab('new')
                setShowEmojiPicker(false)
                setShowStickerDrawer(false)
              }}
              title="Anexar"
            >
              <Glyph name="paperclip" />
            </button>
            <button
              className={`icon-button tiny ${showEmojiPicker ? 'is-active' : ''}`}
              onClick={() => {
                setShowEmojiPicker((current) => !current)
                setShowAttachments(false)
                setShowStickerDrawer(false)
              }}
              title="Emojis"
            >
              <Glyph name="smile" />
            </button>
            <button
              className={`icon-button tiny ${showStickerDrawer ? 'is-active' : ''}`}
              onClick={() => {
                setShowStickerDrawer((current) => !current)
                setShowAttachments(false)
                setShowEmojiPicker(false)
              }}
              title="Figurinhas"
            >
              <Glyph name="image" />
            </button>
            <button
              className={`icon-button tiny ${activePanel === 'quickReplies' ? 'is-active' : ''}`}
              onClick={() => openPanel('quickReplies')}
              title="Respostas rapidas"
            >
              <Glyph name="spark" />
            </button>
            <button
              className={`icon-button tiny ${activePanel === 'chatbot' ? 'is-active' : ''}`}
              onClick={() => openPanel('chatbot')}
              title="Executar chatbot"
            >
              <Glyph name="bot" />
            </button>
            <button
              className={`icon-button tiny ${activePanel === 'schedule' ? 'is-active' : ''}`}
              onClick={() => openPanel('schedule')}
              title="Agendar"
            >
              <Glyph name="calendar" />
            </button>
          </div>
          <button
            className="voice-button"
            onClick={handleSendComposer}
            disabled={!hasConversation || isSending}
          >
            <Glyph name={composerMode === 'message' ? 'send' : 'note'} />
          </button>
        </div>
        <input ref={fileInputRef} type="file" multiple className="hidden-file-input" onChange={handleFileSelection} />
      </footer>
    </div>
  )
}
