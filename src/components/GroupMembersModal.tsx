import { useState, useEffect } from 'react'
import graphClient from '../services/graphClient'

interface GroupMembersModalProps {
  groupId: string
  groupName: string
  onClose: () => void
}

interface GroupMember {
  id: string
  displayName: string
  userPrincipalName?: string
  ['@odata.type']?: string
}

export default function GroupMembersModal ({ groupId, groupName, onClose }: GroupMembersModalProps) {
  const [members, setMembers] = useState<GroupMember[]>([])
  const [loading, setLoading] = useState(true)
  const [error, setError] = useState('')

  useEffect(() => {
    if (!groupId) return
    setLoading(true)
    graphClient.getGroupMembers(groupId)
      .then(res => {
        const users = (res?.value || []).filter(m => m['@odata.type'] === '#microsoft.graph.user')
        setMembers(users)
      })
      .catch((e: Error) => setError(e.message))
      .finally(() => setLoading(false))
  }, [groupId])

  return (
    <div className="modal-overlay" onClick={e => e.target === e.currentTarget && onClose()}>
      <div className="modal modal-lg">
        <div className="modal-title">👥 Members of "{groupName}"</div>
        {loading && (
          <div style={{ textAlign: 'center', padding: 24 }}><span className="spinner spinner-dark" /></div>
        )}
        {error && <div className="err-box mb-3">{error}</div>}
        {!loading && (
          <div className="tbl-wrap" style={{ maxHeight: 360, overflowY: 'auto' }}>
            <table className="tbl">
              <thead>
                <tr><th>Display Name</th><th>User Principal Name</th></tr>
              </thead>
              <tbody>
                {members.length === 0 && (
                  <tr><td colSpan={2} className="tbl-empty">No members found</td></tr>
                )}
                {members.map(m => (
                  <tr key={m.id}>
                    <td style={{ color: 'var(--text)' }}>{m.displayName}</td>
                    <td className="text-xs" style={{ fontFamily: 'DM Mono, monospace' }}>{m.userPrincipalName}</td>
                  </tr>
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
