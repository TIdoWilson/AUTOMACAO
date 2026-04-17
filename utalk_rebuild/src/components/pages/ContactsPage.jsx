import { useState } from 'react'
import { Glyph } from '../shared/Glyph'
import { buildContactState, contactItems, contactTagOptions } from '../../lib/appData'

export function ContactsPage() {
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
