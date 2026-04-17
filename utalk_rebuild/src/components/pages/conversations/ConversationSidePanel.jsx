import { Glyph } from '../../shared/Glyph'

export function ConversationSidePanel({ title, subtitle, children, footer, onClose }) {
  return (
    <aside className="conversation-side-panel">
      <div className="side-panel-header">
        <div>
          <h2>{title}</h2>
          <p>{subtitle}</p>
        </div>
        <button className="icon-button small" onClick={onClose}><Glyph name="x" /></button>
      </div>
      <div className="side-panel-body">{children}</div>
      {footer && <div className="side-panel-footer">{footer}</div>}
    </aside>
  )
}
