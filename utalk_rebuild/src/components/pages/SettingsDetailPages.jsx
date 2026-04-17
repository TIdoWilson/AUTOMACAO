import { useState } from 'react'
import { Glyph } from '../shared/Glyph'
import {
  agentPermissionOptions,
  agentReassignmentOptions,
  buildAgentState,
  buildTemplateState,
} from '../../lib/appData'
import { navigate } from '../../lib/navigation'

export function AgentsPage() {
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

export function TemplatesPage() {
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
              <button
                className="status-tag interactive-tag"
                onClick={() =>
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
                }
              >
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
