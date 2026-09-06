import { useState, useCallback } from 'react'
import graphClient from '../services/graphClient'
import AssignModal, { type AssignSelection } from '../components/AssignModal'
import GroupMembersModal from '../components/GroupMembersModal'
import UserAppsModal from '../components/UserAppsModal'
import ScimModal from '../components/ScimModal'
import { downloadCsv, assignmentsToCsv } from '../utils/exportHelpers'
import type { GraphApp, GraphAssignment } from '../types/graph'
import type { LogFn, ToastFn } from '../types/ui'

interface EntraAppsProps {
  log: LogFn
  toast: ToastFn
  onAppsLoaded?: (apps: GraphApp[]) => void
}

interface PrincipalModal {
  id: string
  name: string
}

export default function EntraApps ({ log, toast, onAppsLoaded }: EntraAppsProps) {
  const [searchQuery, setSearchQuery] = useState('Litera')
  const [apps, setApps] = useState<GraphApp[]>([])
  const [loading, setLoading] = useState(false)
  const [selectedApp, setSelectedApp] = useState<GraphApp | null>(null)

  // Detail state
  const [perms, setPerms] = useState<string[]>([])
  const [hasConsent, setHasConsent] = useState(false)
  const [assignments, setAssignments] = useState<GraphAssignment[]>([])
  const [assignFilter, setAssignFilter] = useState('')
  const [appHidden, setAppHidden] = useState(false)
  const [appTags, setAppTags] = useState<string[]>([])
  const [detailLoading, setDetailLoading] = useState(false)
  const [syncJobs, setSyncJobs] = useState<any[]>([])

  // Modals
  const [showAssignModal, setShowAssignModal] = useState(false)
  const [showScimModal, setShowScimModal] = useState(false)
  const [groupModal, setGroupModal] = useState<PrincipalModal | null>(null)
  const [userModal, setUserModal] = useState<PrincipalModal | null>(null)

  // ── Search ─────────────────────────────────────────────────────────────────
  const searchApps = useCallback(async () => {
    if (!graphClient.isAuthenticated) { toast('Please sign in first', 'error'); return }
    setLoading(true)
    setSelectedApp(null)
    log(`Searching apps: ${searchQuery}`)
    try {
      const res = await graphClient.searchApps(searchQuery)
      const appList = res?.value || []
      setApps(appList)
      onAppsLoaded?.(appList)
      log(`Found ${res?.value?.length || 0} app(s).`)
    } catch (e) {
      log(`ERROR: ${e.message}`)
      toast(e.message, 'error')
    } finally {
      setLoading(false)
    }
  }, [searchQuery, log, toast])

  // ── Select App ────────────────────────────────────────────────────────────
  const selectApp = useCallback(async (app: GraphApp) => {
    setSelectedApp(app)
    setDetailLoading(true)
    setPerms([])
    setAssignments([])
    setSyncJobs([])
    log(`Loading details for: ${app.displayName}`)
    try {
      await Promise.all([
        loadPerms(app),
        loadAssignments(app.id),
        loadVisibility(app.id),
        loadSyncJobs(app.id)
      ])
    } catch (e) {
      log(`Detail load error: ${e.message}`)
    } finally {
      setDetailLoading(false)
    }
  }, [log])

  // ── Permissions ───────────────────────────────────────────────────────────
  const loadPerms = async (app: GraphApp) => {
    try {
      const appQ = await graphClient.getAppPermissions(app.appId)
      const reqAccess = appQ?.value?.[0]?.requiredResourceAccess || []
      const chips = []
      for (const rra of reqAccess) {
        try {
          const spR = await graphClient.get(
            `/servicePrincipals?$filter=appId eq '${rra.resourceAppId}'&$select=displayName,oauth2PermissionScopes,appRoles`
          )
          const sp = spR?.value?.[0]
          if (!sp) continue
          for (const ra of (rra.resourceAccess || [])) {
            let name = ra.id
            if (ra.type === 'Scope') {
              const s = sp.oauth2PermissionScopes?.find(x => x.id === ra.id)
              if (s) name = s.value
            } else {
              const r = sp.appRoles?.find(x => x.id === ra.id)
              if (r) name = r.value || r.displayName
            }
            chips.push(name)
          }
        } catch { /* partial ok */ }
      }
      setPerms(chips)

      const grants = await graphClient.getPermissionGrants(app.id)
      setHasConsent((grants?.value?.length || 0) > 0)
    } catch (e) {
      log(`Perms error: ${e.message}`)
    }
  }

  // ── Assignments ───────────────────────────────────────────────────────────
  const loadAssignments = async (spId: string) => {
    try {
      const raw = await graphClient.getAssignments(spId)
      const enriched = await graphClient.enrichAssignments(raw)
      setAssignments(enriched)
      log(`Loaded ${enriched.length} assignment(s).`)
    } catch (e) {
      log(`Assignments error: ${e.message}`)
    }
  }

  // ── Visibility ────────────────────────────────────────────────────────────
  const loadVisibility = async (spId: string) => {
    try {
      const sp = await graphClient.getSpTags(spId)
      const tags = sp?.tags || []
      setAppTags(tags)
      setAppHidden(tags.includes('HideApp'))
    } catch (e) {
      log(`Visibility error: ${e.message}`)
    }
  }

  const toggleVisibility = async () => {
    if (!selectedApp) return
    const tags = [...appTags]
    const HIDE = 'HideApp'
    const INTEGRATED = 'WindowsAzureActiveDirectoryIntegratedApp'
    if (!tags.includes(INTEGRATED)) tags.push(INTEGRATED)
    let newTags
    if (appHidden) {
      newTags = tags.filter(t => t !== HIDE)
      log('Making app visible in MyApps…')
    } else {
      if (!tags.includes(HIDE)) tags.push(HIDE)
      newTags = tags
      log('Hiding app from MyApps…')
    }
    try {
      await graphClient.setSpTags(selectedApp.id, newTags)
      await graphClient.setAppTagsByAppId(selectedApp.appId, newTags)
      setAppTags(newTags)
      setAppHidden(!appHidden)
      toast('Visibility updated', 'success')
      log('Visibility updated: HideApp=' + (!appHidden))
    } catch (e) {
      log(`Visibility error: ${e.message}`)
      toast(e.message, 'error')
    }
  }

  // ── Sync Jobs ─────────────────────────────────────────────────────────────
  const loadSyncJobs = async (spId: string) => {
    try {
      const jobs = await graphClient.getSyncJobs(spId)
      setSyncJobs(jobs || [])
    } catch { /* non-critical */ }
  }

  const checkSync = async () => {
    if (!selectedApp) return
    log(`Checking sync for: ${selectedApp.displayName}`)
    const jobs = await graphClient.getSyncJobs(selectedApp.id)
    setSyncJobs(jobs || [])
    if (jobs?.length) {
      toast(`Sync: ${jobs[0].status?.code || 'Unknown'}`, 'success')
      log(`Sync job: ${jobs[0].id} — ${jobs[0].status?.code}`)
    } else {
      toast('No sync jobs found')
      log('No sync jobs found.')
    }
  }

  const restartProvisioning = async () => {
    if (!selectedApp || !syncJobs.length) return
    try {
      await graphClient.restartSyncJob(selectedApp.id, syncJobs[0].id)
      toast('Provisioning restart requested', 'success')
      log('Provisioning restart requested.')
    } catch (e) {
      toast(e.message, 'error')
      log(`Restart error: ${e.message}`)
    }
  }

  // ── Consent ───────────────────────────────────────────────────────────────
  const grantConsent = async () => {
    if (!selectedApp) return
    log(`Granting admin consent for: ${selectedApp.displayName}`)
    try {
      const results = await graphClient.grantAdminConsent(selectedApp)
      results.forEach(r => log('  ' + r))
      setHasConsent(true)
      toast('Admin consent granted', 'success')
    } catch (e) {
      log(`Consent error: ${e.message}`)
      toast(e.message, 'error')
    }
  }

  // ── Remove Assignments ────────────────────────────────────────────────────
  const removeChecked = async () => {
    const checked = assignments.filter(a => a._checked)
    if (!checked.length) { toast('No assignments checked', 'error'); return }
    if (!confirm(`Remove ${checked.length} assignment(s)?`)) return
    log(`Removing ${checked.length} assignment(s)…`)
    let ok = 0, fail = 0
    for (const a of checked) {
      try {
        await graphClient.removeAssignment(selectedApp.id, a.id)
        log(`  ✓ Removed: ${a.principalDisplayName}`)
        ok++
      } catch (e) {
        log(`  ✗ Remove error for ${a.principalDisplayName}: ${e.message}`)
        fail++
      }
    }
    await loadAssignments(selectedApp.id)
    toast(`Removed ${ok}${fail ? ` (${fail} failed)` : ''}`, ok ? 'success' : 'error')
  }

  // ── Confirm Assign ────────────────────────────────────────────────────────
  const confirmAssign = async (selected: AssignSelection[]) => {
    setShowAssignModal(false)
    if (!selected.length) return
    log(`Assigning ${selected.length} principal(s) to ${selectedApp.displayName}…`)
    let ok = 0, fail = 0
    for (const p of selected) {
      try {
        await graphClient.addAssignment(selectedApp.id, p.id, p.type)
        log(`  ✓ Assigned: ${p.name}`)
        ok++
      } catch (e) {
        log(`  ✗ ${p.name}: ${e.message}`)
        fail++
      }
    }
    await loadAssignments(selectedApp.id)
    toast(`Assigned ${ok}${fail ? ` (${fail} failed)` : ''}`, ok ? 'success' : 'error')
  }

  // ── Delete App ────────────────────────────────────────────────────────────
  const deleteApp = async () => {
    if (!selectedApp) return
    if (!confirm(`Permanently delete "${selectedApp.displayName}"? This cannot be undone.`)) return
    try {
      await graphClient.deleteServicePrincipal(selectedApp.id)
      log(`Deleted: ${selectedApp.displayName}`)
      toast('App deleted', 'success')
      setSelectedApp(null)
      await searchApps()
    } catch (e) {
      log(`Delete error: ${e.message}`)
      toast(e.message, 'error')
    }
  }

  // ── Filtered assignments ──────────────────────────────────────────────────
  const filteredAssignments = assignments.filter((a) => {
    if (!assignFilter) return true
    const q = assignFilter.toLowerCase()
    return (a.principalDisplayName || '').toLowerCase().includes(q) ||
      (a.emailOrId || '').toLowerCase().includes(q)
  })

  const toggleCheck = (id: string) => {
    setAssignments(prev => prev.map(a => a.principalId === id ? { ...a, _checked: !a._checked } : a))
  }
  const toggleAllCheck = (checked: boolean) => {
    setAssignments(prev => prev.map(a => ({ ...a, _checked: checked })))
  }

  return (
    <div className="page-container page-enter">
      <div className="page-header">
        <div>
          <div className="page-title">Enterprise Applications</div>
          <div className="page-subtitle">Manage Litera apps in Microsoft Entra</div>
        </div>
        <div className="flex gap-2">
          <button className="btn btn-primary" onClick={() => setShowScimModal(true)}>＋ New SCIM App</button>
          <button className="btn btn-secondary" onClick={searchApps}>↻ Refresh</button>
        </div>
      </div>

      <div className="split-layout">
        {/* ── Left: App List ── */}
        <div className="split-left">
          <div className="card card-sm mb-2">
            <div className="inp-group mb-2">
              <span className="inp-icon">🔍</span>
              <input
                className="inp"
                value={searchQuery}
                onChange={e => setSearchQuery(e.target.value)}
                onKeyDown={e => e.key === 'Enter' && searchApps()}
                placeholder="Search apps (e.g. Litera)"
              />
            </div>
            <button className="btn btn-secondary w-full" onClick={searchApps} disabled={loading}>
              {loading ? <><span className="spinner" /> Searching…</> : 'Search'}
            </button>
          </div>

          <div style={{ maxHeight: 'calc(100vh - 230px)', overflowY: 'auto' }}>
            {loading && (
              <div style={{ textAlign: 'center', padding: 24 }}>
                <span className="spinner spinner-dark" />
              </div>
            )}
            {!loading && apps.length === 0 && (
              <div className="empty-state">
                <div className="empty-icon">🗔</div>
                <div>Sign in and search for apps</div>
              </div>
            )}
            {apps.map(app => (
              <div
                key={app.id}
                className={`app-item ${selectedApp?.id === app.id ? 'active' : ''}`}
                onClick={() => selectApp(app)}
              >
                <div className="aname">{app.displayName}</div>
                <div className="aid">{app.id}</div>
              </div>
            ))}
          </div>
        </div>

        {/* ── Right: App Detail ── */}
        {!selectedApp && (
          <div className="card" style={{ display: 'flex', alignItems: 'center', justifyContent: 'center', minHeight: 320 }}>
            <div className="empty-state">
              <div className="empty-icon" style={{ fontSize: 28 }}>←</div>
              <div>Select an app to view details</div>
            </div>
          </div>
        )}

        {selectedApp && (
          <div className="split-right" style={{ maxHeight: 'calc(100vh - 130px)', overflowY: 'auto' }}>
            <div className="card">
              {detailLoading && (
                <div style={{ display: 'flex', justifyContent: 'center', padding: 16 }}>
                  <span className="spinner spinner-dark" />
                </div>
              )}

              {/* Header */}
              <div className="flex items-start justify-between mb-3">
                <div style={{ minWidth: 0, flex: 1 }}>
                  <div className="flex items-center gap-2 mb-1">
                    <span style={{ fontFamily: 'Syne, sans-serif', fontSize: 20, fontWeight: 700 }}>
                      {selectedApp.displayName}
                    </span>
                    {hasConsent && <span title="Admin consent granted">✅</span>}
                    {!hasConsent && <span title="Consent not granted" style={{ opacity: 0.7 }}>⚠️</span>}
                  </div>
                  <div style={{ fontFamily: 'DM Mono, monospace', fontSize: 11, color: 'var(--text3)' }}>
                    SP: {selectedApp.id}
                  </div>
                  <div style={{ fontFamily: 'DM Mono, monospace', fontSize: 11, color: 'var(--text3)' }}>
                    AppID: {selectedApp.appId}
                  </div>
                </div>
                <div className="flex flex-wrap gap-2">
                  {!hasConsent && (
                    <button className="btn btn-warn btn-sm" onClick={grantConsent}>🛡 Grant Consent</button>
                  )}
                  <button className="btn btn-secondary btn-sm" onClick={checkSync}>🔄 Check Sync</button>
                  {syncJobs.length > 0 && (
                    <button className="btn btn-warn btn-sm" onClick={restartProvisioning}>⚡ Restart Prov.</button>
                  )}
                  <button className="btn btn-secondary btn-sm" onClick={() => downloadCsv(assignmentsToCsv(assignments), `${selectedApp.displayName}_assignments.csv`)}>
                    📥 CSV
                  </button>
                  <button className="btn btn-danger btn-sm" onClick={deleteApp}>🗑 Delete</button>
                </div>
              </div>

              {/* Visibility */}
              <div className="vis-panel">
                <button
                  className={`btn btn-sm ${appHidden ? 'btn-green' : 'btn-warn'}`}
                  onClick={toggleVisibility}
                >
                  {appHidden ? '👁 Show in MyApps' : '🙈 Hide from MyApps'}
                </button>
                <div>
                  <div style={{ fontSize: 13, fontWeight: 600, color: appHidden ? '#34d399' : '#fbbf24' }}>
                    {appHidden ? 'Hidden (Recommended for SCIM apps)' : 'Visible in MyApps'}
                  </div>
                  <div className="text-xs">HideApp tag {appHidden ? 'present' : 'absent'}</div>
                </div>
              </div>

              {/* Sync status */}
              {syncJobs.length > 0 && (
                <div className="flex items-center gap-2 mb-3" style={{ background: 'var(--surface2)', border: '1px solid var(--border)', borderRadius: 6, padding: '8px 12px' }}>
                  <span className="dot dot-green" />
                  <span style={{ fontSize: 12.5 }}>Sync job: <strong>{syncJobs[0].status?.code || 'Unknown'}</strong></span>
                  <span className="text-xs" style={{ marginLeft: 'auto' }}>{syncJobs[0].id}</span>
                </div>
              )}

              {/* Permissions */}
              <span className="section-label">API Permissions</span>
              <div className="flex flex-wrap mb-3">
                {perms.length === 0 && !detailLoading && (
                  <span className="perm-chip">Loading…</span>
                )}
                {perms.map((p, i) => <span key={i} className="perm-chip">{p}</span>)}
              </div>

              <hr className="divider" />

              {/* Assignments */}
              <div className="flex items-center justify-between mb-2">
                <span className="section-label" style={{ marginBottom: 0 }}>
                  Assigned Users / Groups ({assignments.length})
                </span>
                <div className="flex gap-2">
                  <button className="btn btn-secondary btn-sm" onClick={() => setShowAssignModal(true)}>＋ Assign</button>
                  <button className="btn btn-danger btn-sm" onClick={removeChecked}>🗑 Remove Checked</button>
                </div>
              </div>

              <div className="inp-group mb-2">
                <span className="inp-icon">🔍</span>
                <input
                  className="inp inp-sm"
                  value={assignFilter}
                  onChange={e => setAssignFilter(e.target.value)}
                  placeholder="Filter assignments…"
                />
              </div>

              <div className="tbl-wrap" style={{ maxHeight: 280, overflowY: 'auto' }}>
                <table className="tbl">
                  <thead>
                    <tr>
                      <th className="chk-col">
                        <input type="checkbox" onChange={e => toggleAllCheck(e.target.checked)} />
                      </th>
                      <th>Name</th>
                      <th>Email / ID</th>
                      <th>Type</th>
                    </tr>
                  </thead>
                  <tbody>
                    {filteredAssignments.length === 0 && (
                      <tr><td colSpan={4} className="tbl-empty">No assignments found</td></tr>
                    )}
                    {filteredAssignments.map(a => (
                      <tr key={a.id} className={a._checked ? 'selected' : ''}>
                        <td className="chk-col">
                          <input type="checkbox" checked={!!a._checked} onChange={() => toggleCheck(a.principalId)} />
                        </td>
                        <td>
                          <span
                            style={{ cursor: 'pointer', color: 'var(--text)', fontWeight: 600 }}
                            onClick={() => {
                              if (a.principalType === 'Group') {
                                setGroupModal({ id: a.principalId, name: a.principalDisplayName })
                              } else {
                                setUserModal({ id: a.principalId, name: a.principalDisplayName })
                              }
                            }}
                          >
                            {a.principalDisplayName || '(unknown)'}
                          </span>
                        </td>
                        <td className="text-xs truncate" style={{ maxWidth: 200, fontFamily: 'DM Mono, monospace' }}>
                          {a.emailOrId}
                        </td>
                        <td>
                          <span className={`badge ${a.principalType === 'User' ? 'badge-blue' : 'badge-purple'}`}>
                            {a.principalType}
                          </span>
                        </td>
                      </tr>
                    ))}
                  </tbody>
                </table>
              </div>
            </div>
          </div>
        )}
      </div>

      {/* Modals */}
      {showAssignModal && (
        <AssignModal
          title={`Assign to ${selectedApp?.displayName}`}
          onConfirm={confirmAssign}
          onClose={() => setShowAssignModal(false)}
        />
      )}
      {showScimModal && (
        <ScimModal
          onCreated={() => { searchApps(); log('SCIM app created — refreshed app list.') }}
          onClose={() => setShowScimModal(false)}
        />
      )}
      {groupModal && (
        <GroupMembersModal
          groupId={groupModal.id}
          groupName={groupModal.name}
          onClose={() => setGroupModal(null)}
        />
      )}
      {userModal && (
        <UserAppsModal
          userId={userModal.id}
          userName={userModal.name}
          literaAppIds={apps.map(a => a.id)}
          onClose={() => setUserModal(null)}
        />
      )}
    </div>
  )
}
