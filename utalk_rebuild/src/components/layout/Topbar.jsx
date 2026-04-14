import { Glyph } from '../shared/Glyph'

export function Topbar({ currentLabel, pathname }) {
  return (
    <header className="topbar">
      <div className="topbar-brand">
        <span className="brand-wordmark dark">umbler talk</span>
        <span className="breadcrumb-label">{pathname.startsWith('/settings/') ? `Configuracoes / ${currentLabel}` : currentLabel}</span>
      </div>
      <div className="topbar-actions">
        <button className="icon-button"><Glyph name="search" /></button>
        <button className="icon-button"><Glyph name="chat" /></button>
        <button className="icon-button"><Glyph name="send" /></button>
        <div className="topbar-avatar">M</div>
      </div>
    </header>
  )
}
