import { useEffect, useState } from 'react'
import { Sidebar } from './components/layout/Sidebar'
import { Topbar } from './components/layout/Topbar'
import { CampaignsPage, ChatbotsPage, DashboardPage } from './components/pages/AnalyticsPages'
import { ContactsPage } from './components/pages/ContactsPage'
import { ConversationsPage } from './components/pages/ConversationsPage'
import { AgentsPage, TemplatesPage } from './components/pages/SettingsDetailPages'
import { ChannelsPage, QuickRepliesPage, SettingsPage, TagsPage } from './components/pages/SettingsPages'
import {
  clearSession,
  getStoredAccessToken,
  getStoredUser,
  loginRequest,
  logoutRequest,
  meRequest,
  refreshRequest,
  storeSession,
} from './lib/api'
import { navItems } from './lib/appData'
import { navigate, usePathname } from './lib/navigation'
import './App.css'

const titleByPath = {
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

const settingsLabelByPath = {
  '/settings/channels': 'Canais de atendimento',
  '/settings/agents': 'Atendentes',
  '/settings/tags': 'Etiquetas',
  '/settings/quick-replies': 'Respostas rapidas',
  '/settings/templates': 'Templates WhatsApp Business API',
}

const knownWorkspaceRoutes = new Set([
  '/',
  '/contacts',
  '/settings',
  '/settings/channels',
  '/settings/agents',
  '/settings/tags',
  '/settings/quick-replies',
  '/settings/templates',
  '/dashboard',
  '/chatbots',
  '/bulksend',
])

function buildSession(result, fallbackUser = null) {
  return {
    accessToken: result.accessToken,
    user: result.user ?? fallbackUser,
  }
}

function userInitial(name = '') {
  const text = String(name).trim()
  return text ? text[0].toUpperCase() : 'U'
}

function App() {
  const pathname = usePathname()
  const [authState, setAuthState] = useState({
    status: 'checking',
    session: null,
    error: '',
  })
  const [isLogoutLoading, setIsLogoutLoading] = useState(false)

  useEffect(() => {
    document.title = `${titleByPath[pathname] ?? 'Umbler Talk'} | Rebuild`
  }, [pathname])

  useEffect(() => {
    let cancelled = false

    async function restoreSession() {
      const accessToken = getStoredAccessToken()
      const storedUser = getStoredUser()

      if (!accessToken || !storedUser) {
        clearSession()
        if (!cancelled) {
          setAuthState({ status: 'anonymous', session: null, error: '' })
        }
        return
      }

      try {
        const me = await meRequest(accessToken)
        const session = { accessToken, user: me?.user ?? storedUser }
        storeSession(session)
        if (!cancelled) {
          setAuthState({ status: 'authenticated', session, error: '' })
        }
        return
      } catch {
        // Falls back to refresh when access token expires.
      }

      try {
        const refreshed = await refreshRequest()
        const session = buildSession(refreshed, storedUser)
        storeSession(session)
        if (!cancelled) {
          setAuthState({ status: 'authenticated', session, error: '' })
        }
      } catch (error) {
        clearSession()
        if (!cancelled) {
          setAuthState({
            status: 'anonymous',
            session: null,
            error: error?.message ?? 'Sua sessao expirou. Faca login novamente.',
          })
        }
      }
    }

    restoreSession()
    return () => {
      cancelled = true
    }
  }, [])

  useEffect(() => {
    if (authState.status === 'checking') return
    if (authState.status === 'anonymous' && pathname !== '/login') {
      navigate('/login')
      return
    }
    if (authState.status === 'authenticated' && pathname === '/login') {
      navigate('/')
    }
  }, [authState.status, pathname])

  async function handleLogin({ email, password }) {
    const result = await loginRequest({
      email: String(email || '').trim().toLowerCase(),
      password,
    })
    const session = buildSession(result)
    storeSession(session)
    setAuthState({ status: 'authenticated', session, error: '' })
    navigate('/')
  }

  async function handleLogout() {
    if (!authState.session?.accessToken) {
      clearSession()
      setAuthState({ status: 'anonymous', session: null, error: '' })
      navigate('/login')
      return
    }

    setIsLogoutLoading(true)
    try {
      await logoutRequest(authState.session.accessToken)
    } catch {
      // Session cleanup should happen even when backend logout fails.
    } finally {
      clearSession()
      setAuthState({ status: 'anonymous', session: null, error: '' })
      setIsLogoutLoading(false)
      navigate('/login')
    }
  }

  if (authState.status === 'checking') {
    return (
      <div className="login-shell">
        <section className="login-panel">
          <div className="brand-row">
            <img src="/assets/favicon.svg" alt="" className="brand-favicon" />
            <span className="brand-wordmark">umbler talk</span>
          </div>
          <div className="status-banner">
            <strong>Validando sessao</strong>
            <span>Aguarde enquanto carregamos seu ambiente.</span>
          </div>
        </section>
        <section className="promo-panel" />
      </div>
    )
  }

  if (pathname === '/login') {
    return <LoginPage onLogin={handleLogin} initialError={authState.error} />
  }

  if (authState.status !== 'authenticated' || !authState.session) {
    return null
  }

  return (
    <WorkspacePage
      pathname={pathname}
      session={authState.session}
      onLogout={handleLogout}
      isLogoutLoading={isLogoutLoading}
    />
  )
}

function LoginPage({ onLogin, initialError = '' }) {
  const [email, setEmail] = useState('')
  const [password, setPassword] = useState('')
  const [keepConnected, setKeepConnected] = useState(true)
  const [isSubmitting, setIsSubmitting] = useState(false)
  const [error, setError] = useState(initialError)

  useEffect(() => {
    if (initialError) {
      setError(initialError)
    }
  }, [initialError])

  async function handleSubmit(event) {
    event.preventDefault()
    setError('')
    setIsSubmitting(true)
    try {
      await onLogin({ email, password, keepConnected })
    } catch (requestError) {
      setError(requestError?.message ?? 'Nao foi possivel realizar o login.')
    } finally {
      setIsSubmitting(false)
    }
  }

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
          <button className="social-card" type="button" disabled>
            <span className="mini-avatar">M</span>
            <span className="social-meta">
              <strong>Login social desabilitado</strong>
              <small>Use e-mail e senha</small>
            </span>
            <span className="social-badge google">G</span>
          </button>
          <button className="social-facebook" type="button" disabled>
            Facebook indisponivel
          </button>
        </div>

        <div className="divider">ou</div>

        <form className="login-form" onSubmit={handleSubmit}>
          <label className="form-field">
            <span>E-mail</span>
            <input
              type="email"
              autoComplete="username"
              value={email}
              onChange={(event) => setEmail(event.target.value)}
              required
            />
          </label>

          <label className="form-field">
            <span>Senha</span>
            <input
              type="password"
              autoComplete="current-password"
              value={password}
              onChange={(event) => setPassword(event.target.value)}
              required
            />
          </label>

          <div className="login-actions">
            <label className="checkbox-line">
              <input
                type="checkbox"
                checked={keepConnected}
                onChange={(event) => setKeepConnected(event.target.checked)}
              />
              <span>Manter-me conectado</span>
            </label>
            <a href="/login">Esqueci minha senha</a>
          </div>

          {error && (
            <div className="status-banner" style={{ marginTop: '8px' }}>
              <strong>Falha no login</strong>
              <span>{error}</span>
            </div>
          )}

          <button className="primary-button" type="submit" disabled={isSubmitting}>
            {isSubmitting ? 'Entrando...' : 'Entrar'}
          </button>
        </form>
      </section>

      <section className="promo-panel">
        <div className="promo-hero">
          <div className="promo-copy">
            <h2>Todo seu time oferecendo suporte em um so WhatsApp</h2>
            <p>
              Com o Umbler Talk, voce coloca quantos operadores quiser atendendo simultaneamente em varios
              computadores usando um unico numero.
            </p>
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
                <button type="button">Criar etiqueta</button>
              </div>
              <div className="promo-floating-input" />
              <div className="promo-floating-chip">suporte</div>
            </div>
          </div>
        </div>
      </section>
    </div>
  )
}

function WorkspacePage({ pathname, session, onLogout, isLogoutLoading }) {
  const isConversationRoute = pathname === '/' || pathname.startsWith('/chats')
  const currentLabel =
    settingsLabelByPath[pathname] ?? navItems.find((item) => item.path === pathname)?.label ?? 'Configuracoes'

  return (
    <div className={`workspace-shell ${isConversationRoute ? 'is-chat' : ''}`}>
      <Sidebar
        pathname={pathname}
        user={session.user}
        onLogout={onLogout}
        isLogoutLoading={isLogoutLoading}
      />
      <main className="workspace-main">
        <Topbar
          currentLabel={currentLabel}
          pathname={pathname}
          userInitial={userInitial(session.user?.name)}
        />
        <div className="workspace-body">
          {pathname === '/' && <ConversationsPage session={session} />}
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
          {!knownWorkspaceRoutes.has(pathname) && <SettingsPage />}
        </div>
      </main>
    </div>
  )
}

export default App
