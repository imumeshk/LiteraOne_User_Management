import { useState, useEffect } from 'react'
import graphClient from '../services/graphClient'

interface UserAppsModalProps {
  userId: string
  userName: string
  literaAppIds: string[]
  onClose: () => void
}

interface UserAssignment {
  resourceDisplayName?: string
  resourceId: string
}

export default function UserAppsModal ({ userId, userName, literaAppIds, onClose }: UserAppsModalProps) {
  const [apps, setApps] = useState<UserAssignment[]>([])
  const [loading, setLoading] = useState(true)
  const [error, setError] = useState('')

  useEffect(() => {
    if (!userId) return
    setLoading(true)
    graphClient.getUserAppAssignments(userId)
      .then(res => {
        const spSet = new Set(literaAppIds)
        const mine = (res?.value || []).filter(a => spSet.has(a.resourceId))
        setApps(mine)
      })
      .catch((e: Error) => setError(e.message))
      .finally(() => setLoading(false))
  }, [userId])

  return (
    <div className="modal-overlay" onClick={e => e.target === e.currentTarget && onClose()}>
      <div className="modal modal-lg">
        <div className="modal-title">👤 Litera Apps for "{userName}"</div>
        {loading && <div style={{ textAlign: 'center', padding: 24 }}><span className="spinner spinner-dark" /></div>}
        {error && <div className="err-box mb-3">{error}</div>}
        {!loading && (
          <div className="tbl-wrap" style={{ maxHeight: 360, overflowY: 'auto' }}>
            <table className="tbl">
              <thead><tr><th>Application</th></tr></thead>
              <tbody>
                {apps.length === 0 && (
                  <tr><td className="tbl-empty">No Litera app assignments found</td></tr>
                )}
                {apps.map((a, i) => (
                  <tr key={i}><td style={{ color: 'var(--text)' }}>{a.resourceDisplayName}</td></tr>
                ))}
              </tbody>
            </table>
          </div>
        )}
        <div className="modal-footer">
          <button className="btn btn-secondary" onClick={onClose}>Close</button>
        </div>
      </div>
    </div>
  )
}
