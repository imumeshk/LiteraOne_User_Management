import type { Page } from '../types/ui'

interface SidebarProps {
  activePage: Page
  onNav: (page: Page) => void
  connected: boolean
  orgName: string
  tenantId: string
  onAuthClick: () => void
  consoleOpen: boolean
  onToggleConsole: () => void
}

export default function Sidebar ({
  activePage, onNav, connected, orgName, tenantId,
  onAuthClick, consoleOpen, onToggleConsole
}: SidebarProps) {
  const nav = [
    { id: 'dashboard', icon: '⊞', label: 'Dashboard' },
    { id: 'apps', icon: '🗔', label: 'Entra Apps' },
    { id: 'deploy', icon: '☁', label: 'Deployments' },
    ...(connected ? [{ id: 'users-management', icon: '👥', label: 'Users Management' }] : [])
  ]

  const initials = orgName ? orgName[0].toUpperCase() : '?'

  return (
    <aside className="sidebar">
      <div className="sidebar-logo">
        <h1>Litera One</h1>
        <p>Manager Portal</p>
      </div>

      <nav className="sidebar-nav">
        <div className="nav-section-label">Navigation</div>
        {nav.map(n => (
          <button
            key={n.id}
            className={`nav-btn ${activePage === n.id ? 'active' : ''}`}
            onClick={() => onNav(n.id)}
          >
            <span className="nav-icon">{n.icon}</span>
            {n.label}
          </button>
        ))}
        <div className="nav-section-label" style={{ marginTop: 16 }}>Tools</div>
        <button
          className={`nav-btn ${consoleOpen ? 'active' : ''}`}
          onClick={onToggleConsole}
        >
          <span className="nav-icon">📋</span>
          Console Logs
        </button>
        <button
          className={`nav-btn ${activePage === 'enablement' ? 'active' : ''}`}
          onClick={() => onNav('enablement')}
        >
          <span className="nav-icon">📖</span>
          Enablement Hub
        </button>
      </nav>

      <div className="sidebar-footer">
        <div className="user-card">
          <div className="user-avatar">{initials}</div>
          <div className="user-info">
            <div className="user-name">{connected ? orgName : 'Not Signed In'}</div>
            <div className="user-role">{connected ? tenantId : 'Connect to get started'}</div>
          </div>
        </div>

        <select className="tenant-select" defaultValue="organizations">
          <option value="organizations">organizations</option>
          <option value="common">common</option>
          <option value="consumers">consumers</option>
        </select>

        <button
          className={`btn-signin ${connected ? 'connected' : ''}`}
          onClick={onAuthClick}
        >
          {connected ? '✕  Sign Out' : '→  Sign In'}
        </button>
      </div>
    </aside>
  )
}
