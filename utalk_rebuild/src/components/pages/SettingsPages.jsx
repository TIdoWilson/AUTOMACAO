import { useState } from 'react'
import { Glyph } from '../shared/Glyph'
import { buildChannelSettingsState, buildQuickReplyState, buildTagState, settingRoutes, settingsGroups } from '../../lib/appData'
import { navigate } from '../../lib/navigation'

export function SettingsPage() {
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
              <span className="settings-arrow">{'>'}</span>
            </button>
          )
        })}
      </div>
    </section>
  )
}

export function ChannelsPage() {
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

export function TagsPage() {
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

export function QuickRepliesPage() {
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
              <small>{reply.shortcut} - {reply.scope}</small>
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
