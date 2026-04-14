import { useState } from 'react'
import { Glyph } from '../shared/Glyph'
import { buildCampaignState, chatbotCards, dashboardCards } from '../../lib/appData'

export function DashboardPage() {
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

export function ChatbotsPage() {
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

export function CampaignsPage() {
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
            <button
              className="status-tag interactive-tag"
              onClick={() => setCampaigns((current) => current.map((item) => item.id === row.id ? { ...item, status: item.status === 'Programada' ? 'Concluida' : item.status === 'Concluida' ? 'Rascunho' : 'Programada' } : item))}
            >
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
