import { useState } from 'react'
import { navItems } from '../../lib/appData'
import { navigate } from '../../lib/navigation'
import { Glyph } from '../shared/Glyph'

export function Sidebar({ pathname }) {
  const [showProfileMenu, setShowProfileMenu] = useState(false)

  return (
    <aside className="sidebar">
      <button className="sidebar-logo" onClick={() => navigate('/')}>
        <img src="/assets/favicon.svg" alt="" />
      </button>

      <nav className="sidebar-nav">
        {navItems.map((item) => (
          <button
            key={item.key}
            className={`sidebar-link ${pathname === item.path ? 'active' : ''}`}
            title={item.label}
            onClick={() => navigate(item.path)}
          >
            <Glyph name={item.icon} />
          </button>
        ))}
      </nav>

      <div className="sidebar-footer">
        <button className="sidebar-link muted"><Glyph name="bell" /></button>
        <button className="sidebar-link muted"><Glyph name="spark" /></button>
        <button className="presence-avatar profile-trigger" onClick={() => setShowProfileMenu((current) => !current)}>M</button>
        {showProfileMenu && (
          <div className="profile-menu-card">
            <div className="profile-menu-header">
              <div className="presence-avatar large-profile">M</div>
              <strong>Operador atual</strong>
              <span>Minha conta</span>
            </div>
            <div className="profile-menu-tabs">
              <button className="active">Minha conta</button>
              <button>Preferencias</button>
            </div>
            <div className="profile-menu-links">
              <button>Meu perfil</button>
              <button>Assinaturas e planos</button>
              <button>Indique e ganhe</button>
            </div>
            <div className="profile-menu-orgs">
              <strong>Minhas organizacoes</strong>
              <button className="org-pill">Empresa interna principal</button>
            </div>
            <button className="profile-menu-logout">Sair</button>
          </div>
        )}
      </div>
    </aside>
  )
}
