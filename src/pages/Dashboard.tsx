import type { Page } from '../types/ui'

interface DashboardProps {
  connected: boolean
  orgName: string
  tenantId: string
  lastSync: string
  onNav: (page: Page) => void
}

export default function Dashboard ({ connected, orgName, tenantId, lastSync, onNav }: DashboardProps) {
  return (
    <div className="page-container page-enter">
      <div className="page-header">
        <div>
          <div className="page-title">
            {connected ? `Welcome, ${orgName}` : 'Welcome'}
          </div>
          <div className="page-subtitle">Litera One Manager Portal — Microsoft Entra Administration</div>
        </div>
      </div>

      <div className="tiles-grid">
        <div className="tile" onClick={() => onNav('apps')}>
          <div className="tile-icon ti-blue">🗔</div>
          <h3>Entra Apps</h3>
          <p>Manage app permissions, users, groups, and SCIM provisioning for Litera One enterprise applications.</p>
        </div>
        <div className="tile" onClick={() => onNav('deploy')}>
          <div className="tile-icon ti-purple">☁</div>
          <h3>Deploy Add-ins</h3>
          <p>Roll out Outlook and Word add-ins to your organization via AppSource or manifest deployment.</p>
        </div>
        <div className="tile" onClick={() => onNav('enablement')}>
          <div className="tile-icon ti-green">📖</div>
          <h3>Enablement Hub</h3>
          <p>Browse and export training resources, guides, videos, and templates for end-users.</p>
        </div>
        {connected && (
          <div className="tile" onClick={() => onNav('users-management')}>
            <div className="tile-icon ti-blue">👥</div>
            <h3>Users Management</h3>
            <p>Assign or remove users/groups across multiple Litera One apps and export history.</p>
          </div>
        )}
      </div>

      <div className="status-card">
        <span className="section-label">System Status</span>
        <div className="status-row">
          <span className={`dot ${connected ? 'dot-green' : 'dot-red'}`} />
          <span className="status-text">
            Microsoft Graph API: {connected ? 'Connected' : 'Disconnected'}
          </span>
          {connected && (
            <span className="badge badge-green" style={{ marginLeft: 8 }}>Active</span>
          )}
        </div>
        <div style={{ height: 4, borderRadius: 20, background: 'var(--surface3)', overflow: 'hidden', marginBottom: 10 }}>
          <div style={{
            height: '100%',
            width: connected ? '100%' : '0%',
            background: connected
              ? 'linear-gradient(90deg, #22c55e, #10b981)'
              : '#ef4444',
            borderRadius: 20,
            transition: 'all 0.6s'
          }} />
        </div>
        {connected && (
          <div className="info-grid" style={{ marginBottom: 0 }}>
            <div className="info-item">
              <label>Organization</label>
              <span>{orgName}</span>
            </div>
            <div className="info-item">
              <label>Tenant ID</label>
              <span style={{ fontFamily: 'DM Mono, monospace', fontSize: 12 }}>{tenantId}</span>
            </div>
            <div className="info-item">
              <label>Last Checked</label>
              <span>{lastSync}</span>
            </div>
            <div className="info-item">
              <label>Auth Method</label>
              <span>Client Credentials</span>
            </div>
          </div>
        )}
        {!connected && (
          <div className="text-xs" style={{ color: 'var(--text3)' }}>
            Click Sign In to connect using the Entra app configured in your .env file.
          </div>
        )}
      </div>
    </div>
  )
}
