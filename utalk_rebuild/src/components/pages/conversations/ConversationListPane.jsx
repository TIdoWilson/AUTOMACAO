import { Glyph } from '../../shared/Glyph'
import { bulkActionItems, queueActionItems, sortActionItems, tabDefinitions } from '../../../lib/appData'
import { useConversationContext } from './ConversationContext'

export function ConversationListPane() {
  const {
    activeQueueMenuId,
    activeTab,
    conversations,
    handleBulkAction,
    handleOpenConversation,
    handleQueueAction,
    isLoadingConversations,
    listPopover,
    searchTerm,
    selectedConversationId,
    setActiveQueueMenuId,
    setActiveTab,
    setListPopover,
    setNotice,
    setSearchTerm,
    setShowContactPicker,
    setSortMode,
    sortMode,
    toggleConversationSelected,
    visibleConversations,
  } = useConversationContext()

  return (
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
                onChange={(event) => toggleConversationSelected(item.id, event.target.checked)}
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
            <strong>{isLoadingConversations ? 'Carregando conversas...' : 'Nenhuma conversa encontrada'}</strong>
            <span>
              {isLoadingConversations
                ? 'Aguarde enquanto sincronizamos a fila.'
                : 'Ajuste a busca, troque a aba ou abra um contato para iniciar uma conversa.'}
            </span>
          </div>
        )}
      </div>
    </aside>
  )
}
