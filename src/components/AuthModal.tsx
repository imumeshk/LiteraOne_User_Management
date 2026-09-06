import { useState, type KeyboardEvent } from 'react'

interface AuthModalProps {
  onConnect: () => Promise<void> | void
  onClose: () => void
  tenantAuthority: string
  configured: boolean
}

export default function AuthModal ({ onConnect, onClose, tenantAuthority, configured }: AuthModalProps) {
  const [loading, setLoading] = useState(false)
  const [error, setError] = useState('')

  const handleConnect = async () => {
    setError('')
    if (!configured) {
      setError('Missing Entra app configuration in .env. Add ENTRA/VITE_ENTRA_CLIENT_ID and ENTRA/VITE_ENTRA_CLIENT_SECRET, then restart.')
      return
    }
    setLoading(true)
    try {
      await onConnect()
    } catch (e: any) {
      setError(e.message)
    } finally {
      setLoading(false)
    }
  }

  const handleKey = (e: KeyboardEvent<HTMLInputElement>) => { if (e.key === 'Enter') handleConnect() }

  return (
    <div className="modal-overlay" onClick={e => e.target === e.currentTarget && onClose()}>
      <div className="modal">
        <div className="modal-title">🔐 Connect with Microsoft Entra</div>

        <p className="text-sm mb-3" style={{ color: 'var(--text3)', lineHeight: 1.5 }}>
          This portal uses your configured multi-tenant Entra app from <code>.env</code>.
          Authentication uses OAuth 2.0 <strong style={{ color: 'var(--text2)' }}>Client Credentials</strong> with the
          <strong style={{ color: 'var(--text2)' }}> {tenantAuthority}</strong> authority.
        </p>

        <div className="form-row">
          <label className="label">Tenant Authority</label>
          <input className="inp inp-mono" value={tenantAuthority} readOnly onKeyDown={handleKey} autoFocus />
          <div className="hint">Using ENTRA/VITE_ENTRA_TENANT_ID (default: organizations)</div>
        </div>

        <div className="info-box mb-3">
          <strong>Required Application Permissions (application type):</strong><br />
          Directory.Read.All · Application.ReadWrite.All · AppRoleAssignment.ReadWrite.All<br />
          Group.Read.All · Synchronization.ReadWrite.All · DelegatedPermissionGrant.ReadWrite.All
        </div>

        {error && <div className="err-box mb-3">{error}</div>}

        <div className="modal-footer" style={{ marginTop: 0 }}>
          <button className="btn btn-secondary" onClick={onClose} disabled={loading}>Cancel</button>
          <button className="btn btn-primary" onClick={handleConnect} disabled={loading || !configured}>
            {loading && <span className="spinner" />}
            Connect
          </button>
        </div>
      </div>
    </div>
  )
}
