import { useState } from 'react'
import { Glyph } from '../../shared/Glyph'

function colorFromText(text) {
  const palette = ['#d48377', '#47a6ff', '#ffd577', '#8bbd68', '#5b80ff', '#9f74de', '#44b6a8']
  const input = String(text || '')
  let hash = 0
  for (let index = 0; index < input.length; index += 1) {
    hash = (hash << 5) - hash + input.charCodeAt(index)
    hash |= 0
  }
  return palette[Math.abs(hash) % palette.length]
}

export function ContactPickerModal({ contacts, onClose, onOpenContact }) {
  const [query, setQuery] = useState('')

  const normalizedQuery = query.toLowerCase().trim()
  const filteredContacts =
    normalizedQuery === ''
      ? contacts
      : contacts.filter((contact) => {
          const haystack = `${contact.displayName} ${contact.channelIdentifier} ${
            contact.phoneMasked || ''
          }`.toLowerCase()
          return haystack.includes(normalizedQuery)
        })

  return (
    <div className="modal-scrim" onClick={onClose}>
      <div className="contact-picker-modal" onClick={(event) => event.stopPropagation()}>
        <div className="side-panel-header">
          <div>
            <h2>Escolha um contato</h2>
            <p>Selecione um contato real para abrir ou iniciar uma conversa.</p>
          </div>
          <button className="icon-button small" onClick={onClose}>
            <Glyph name="x" />
          </button>
        </div>
        <div className="search-field">
          <Glyph name="search" />
          <input
            type="text"
            placeholder="Pesquisar contato"
            value={query}
            onChange={(event) => setQuery(event.target.value)}
          />
        </div>
        <div className="contact-picker-list">
          {filteredContacts.map((item) => (
            <button key={item.id} className="contact-picker-row" onClick={() => onOpenContact(item)}>
              <div className="contact-cell">
                <div className="avatar small" style={{ '--avatar': colorFromText(item.channelIdentifier) }}>
                  {item.displayName?.[0] || 'C'}
                </div>
                <div>
                  <strong>{item.displayName}</strong>
                  <span>{item.phoneMasked || 'Telefone oculto'}</span>
                </div>
              </div>
              <span className="chip interactive">Abrir</span>
            </button>
          ))}
          {filteredContacts.length === 0 && (
            <div className="empty-inline-state">
              <strong>Nenhum contato encontrado</strong>
              <span>Ajuste os termos de pesquisa para localizar um contato.</span>
            </div>
          )}
        </div>
      </div>
    </div>
  )
}
