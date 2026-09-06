import { useCallback, useEffect, useMemo, useRef, useState } from 'react'
import { downloadCsv } from '../utils/exportHelpers'
import type { GraphPrincipal, GraphApp } from '../types/graph'
import type { LogFn, ToastFn } from '../types/ui'

type PrincipalSearchResult = GraphPrincipal

type SelectedPrincipal = {
  id: string
  name: string
  emailOrId: string
  type: 'User' | 'Group'
}

type LiteraApp = GraphApp

type ActionHistoryItem = {
  time: string
  action: string
  appName: string
  appId: string
  principalName: string
  principalType: string
  principalId: string
  emailOrId: string
  status: string
  error?: string
}

interface LiteraUsersManagementProps {
  connected: boolean
  log: LogFn
  toast: ToastFn
}

async function api (method: string, path: string, body?: unknown) {
  const opts = {
    method,
    headers: {
      Accept: 'application/json',
      'Content-Type': 'application/json'
    },
    credentials: 'same-origin'
  }
  if (body !== undefined) opts.body = JSON.stringify(body)
  const res = await fetch('/api' + path, opts)
  const data = await res.json().catch(() => ({}))
  if (!res.ok) throw new Error(data.error || `HTTP ${res.status}`)
  return data
}

const GET = (p) => api('GET', p)
const POST = (p, b) => api('POST', p, b)
const DEL = (p) => api('DELETE', p)

export default function LiteraUsersManagement ({ connected, log, toast }: LiteraUsersManagementProps) {
  const [mode, setMode] = useState<'assign' | 'remove'>('assign')
  const [tab, setTab] = useState<'users' | 'groups'>('users')
  const [query, setQuery] = useState('')
  const [appQuery, setAppQuery] = useState('')
  const [results, setResults] = useState<PrincipalSearchResult[]>([])
  const [selectedPrincipals, setSelectedPrincipals] = useState<SelectedPrincipal[]>([])
  const [apps, setApps] = useState<LiteraApp[]>([])
  const [selectedAppIds, setSelectedAppIds] = useState<string[]>([])
  const [history, setHistory] = useState<ActionHistoryItem[]>([])
  const [searching, setSearching] = useState(false)
  const [running, setRunning] = useState(false)
  const searchTimer = useRef<ReturnType<typeof setTimeout> | null>(null)
  const selectAllResultsRef = useRef<HTMLInputElement | null>(null)
  const selectAllAppsRef = useRef<HTMLInputElement | null>(null)

  const selectedResultCount = useMemo(
    () => results.filter(r => selectedPrincipals.some(s => s.id === r.id)).length,
    [results, selectedPrincipals]
  )
  const allResultsSelected = results.length > 0 && selectedResultCount === results.length
  const someResultsSelected = selectedResultCount > 0 && !allResultsSelected
  const allAppsSelected = apps.length > 0 && selectedAppIds.length === apps.length
  const someAppsSelected = selectedAppIds.length > 0 && !allAppsSelected
  const filteredApps = useMemo(() => {
    const q = appQuery.trim().toLowerCase()
    if (!q) return apps
    return apps.filter(app =>
      (app.displayName || '').toLowerCase().includes(q) ||
      (app.id || '').toLowerCase().includes(q)
    )
  }, [appQuery, apps])
  const historySuccessCount = useMemo(
    () => history.filter(item => item.status === 'Success').length,
    [history]
  )
  const historyFailureCount = history.length - historySuccessCount

  useEffect(() => {
    if (!selectAllResultsRef.current) return
    selectAllResultsRef.current.indeterminate = someResultsSelected
  }, [someResultsSelected])

  useEffect(() => {
    if (!selectAllAppsRef.current) return
    selectAllAppsRef.current.indeterminate = someAppsSelected
  }, [someAppsSelected])

  const searchPrincipals = useCallback(async (q: string, t: 'users' | 'groups') => {
    if (!connected) return
    setSearching(true)
    try {
      const qs = new URLSearchParams({ type: t, top: '20' })
      if (q.trim()) qs.set('q', q.trim())
      const data = await GET(`/principals/search?${qs.toString()}`)
      setResults(data.principals || [])
    } catch (e) {
      setResults([])
      toast(e.message, 'error')
    } finally {
      setSearching(false)
    }
  }, [connected, toast])

  const loadApps = useCallback(async () => {
    if (!connected) return
    try {
      const data = await GET('/apps/search?q=Litera')
      setApps(data.apps || [])
      setSelectedAppIds([])
    } catch (e) {
      setApps([])
      toast(e.message, 'error')
    }
  }, [connected, toast])

  useEffect(() => {
    if (!connected) return
    loadApps()
  }, [connected, loadApps])

  useEffect(() => {
    if (!connected) return
    clearTimeout(searchTimer.current)
    searchTimer.current = setTimeout(() => {
      searchPrincipals(query, tab)
    }, 280)
    return () => clearTimeout(searchTimer.current)
  }, [connected, query, tab, searchPrincipals])

  const togglePrincipal = (item: PrincipalSearchResult, explicitTab: 'users' | 'groups' = tab) => {
    const sub = item.userPrincipalName || item.mail || item.id
    const type = explicitTab === 'groups' ? 'Group' : 'User'
    setSelectedPrincipals(prev => {
      const exists = prev.some(p => p.id === item.id)
      if (exists) return prev.filter(p => p.id !== item.id)
      return [...prev, { id: item.id, name: item.displayName, emailOrId: sub, type }]
    })
  }

  const toggleAllResults = (checked: boolean) => {
    if (!results.length) return
    setSelectedPrincipals(prev => {
      if (!checked) {
        const resultIds = new Set(results.map(r => r.id))
        return prev.filter(p => !resultIds.has(p.id))
      }
      const next = [...prev]
      const existing = new Set(prev.map(p => p.id))
      for (const r of results) {
        if (existing.has(r.id)) continue
        const sub = r.userPrincipalName || r.mail || r.id
        const type = tab === 'groups' ? 'Group' : 'User'
        next.push({ id: r.id, name: r.displayName, emailOrId: sub, type })
      }
      return next
    })
  }

  const toggleApp = (appId: string) => {
    setSelectedAppIds(prev => prev.includes(appId) ? prev.filter(id => id !== appId) : [...prev, appId])
  }

  const toggleAllApps = (checked: boolean) => {
    setSelectedAppIds(checked ? apps.map(a => a.id) : [])
  }

  const clearPrincipals = () => setSelectedPrincipals([])
  const clearApps = () => setSelectedAppIds([])
  const record = (entry: ActionHistoryItem) => setHistory(prev => [...prev, entry])

  const runAction = async () => {
    if (!selectedPrincipals.length) { toast('Select at least one user/group', 'error'); return }
    if (!selectedAppIds.length) { toast('Select at least one Litera app', 'error'); return }

    setRunning(true)
    const action = mode === 'assign' ? 'Assigned' : 'Removed'
    let ok = 0
    let fail = 0

    try {
      for (const appId of selectedAppIds) {
        const app = apps.find(a => a.id === appId)
        const appName = app?.displayName || appId

        if (mode === 'assign') {
          for (const p of selectedPrincipals) {
            try {
              await POST(`/apps/${appId}/assignments`, { principalId: p.id, principalType: p.type })
              ok++
              record({
                time: new Date().toLocaleString(),
                action,
                appName,
                appId,
                principalName: p.name,
                principalType: p.type,
                principalId: p.id,
                emailOrId: p.emailOrId,
                status: 'Success',
                error: ''
              })
            } catch (e) {
              fail++
              record({
                time: new Date().toLocaleString(),
                action,
                appName,
                appId,
                principalName: p.name,
                principalType: p.type,
                principalId: p.id,
                emailOrId: p.emailOrId,
                status: 'Failed',
                error: e.message
              })
            }
          }
          continue
        }

        let appAssignments = []
        try {
          const data = await GET(`/apps/${appId}/assignments`)
          appAssignments = data.assignments || []
        } catch (e) {
          for (const p of selectedPrincipals) {
            fail++
            record({
              time: new Date().toLocaleString(),
              action,
              appName,
              appId,
              principalName: p.name,
              principalType: p.type,
              principalId: p.id,
              emailOrId: p.emailOrId,
              status: 'Failed',
              error: e.message
            })
          }
          continue
        }

        for (const p of selectedPrincipals) {
          const matches = appAssignments.filter(a => a.principalId === p.id)
          if (!matches.length) {
            fail++
            record({
              time: new Date().toLocaleString(),
              action,
              appName,
              appId,
              principalName: p.name,
              principalType: p.type,
              principalId: p.id,
              emailOrId: p.emailOrId,
              status: 'Failed',
              error: 'Assignment not found'
            })
            continue
          }

          for (const a of matches) {
            try {
              await DEL(`/apps/${appId}/assignments/${a.id}`)
              ok++
              record({
                time: new Date().toLocaleString(),
                action,
                appName,
                appId,
                principalName: p.name,
                principalType: p.type,
                principalId: p.id,
                emailOrId: p.emailOrId,
                status: 'Success',
                error: ''
              })
            } catch (e) {
              fail++
              record({
                time: new Date().toLocaleString(),
                action,
                appName,
                appId,
                principalName: p.name,
                principalType: p.type,
                principalId: p.id,
                emailOrId: p.emailOrId,
                status: 'Failed',
                error: e.message
              })
            }
          }
        }
      }

      log(`${action} users/groups: ${ok} success${fail ? `, ${fail} failed` : ''}`)
      toast(`${action}: ${ok} success${fail ? `, ${fail} failed` : ''}`, ok ? 'success' : 'error')
    } finally {
      setRunning(false)
    }
  }

  const exportHistory = () => {
    if (!history.length) { toast('No history to export', 'error'); return }
    const rows = [
      ['Time', 'Action', 'App Name', 'App ID', 'Principal Name', 'Principal Type', 'Principal ID', 'Email / ID', 'Status', 'Error'],
      ...history.map(h => [h.time, h.action, h.appName, h.appId, h.principalName, h.principalType, h.principalId, h.emailOrId, h.status, h.error || ''])
    ]
    downloadCsv(rows, 'litera_one_users_management_history.csv')
  }

  if (!connected) {
    return (
      <div className="page-container page-enter">
        <div className="card" style={{ minHeight: 240, display: 'flex', alignItems: 'center', justifyContent: 'center' }}>
          <div className="empty-state">
            <div className="empty-icon">🔐</div>
            <div>Sign in to Microsoft to use Litera One Users Management</div>
          </div>
        </div>
      </div>
    )
  }

  return (
    <div className="page-container page-enter">
      <div className="users-mgmt-breadcrumbs">
        <span>Dashboard</span>
        <span>/</span>
        <span>Users</span>
        <span>/</span>
        <strong>Manage Users</strong>
      </div>

      <div className="page-header users-mgmt-page-head">
        <div>
          <div className="page-title">Manage Users</div>
          <div className="page-subtitle">Assign or remove users and groups across Litera One applications.</div>
        </div>
        <div className="users-mgmt-toolbar-meta">
          <button className="btn btn-secondary" onClick={loadApps}>↻ Refresh Apps</button>
          <button className="btn btn-secondary" onClick={exportHistory}>📥 Export History</button>
        </div>
      </div>

      <div className="users-mgmt-commandbar card mb-3">
        <div className="users-mgmt-commandbar-left">
          <div className="users-mgmt-button-group">
            <button className={`users-mgmt-solid-btn ${mode === 'assign' ? 'green' : 'ghost'}`} onClick={() => setMode('assign')}>
              Add User
            </button>
            <button className={`users-mgmt-solid-btn ${mode === 'remove' ? 'blue' : 'ghost'}`} onClick={() => setMode('remove')}>
              Remove User
            </button>
          </div>
          <div className="users-mgmt-button-group">
            <button className={`users-mgmt-solid-btn ${tab === 'users' ? 'active-filter' : 'ghost'}`} onClick={() => setTab('users')}>
              Users
            </button>
            <button className={`users-mgmt-solid-btn ${tab === 'groups' ? 'active-filter' : 'ghost'}`} onClick={() => setTab('groups')}>
              Groups
            </button>
          </div>
        </div>
        <div className="users-mgmt-commandbar-right">
          <div className="inp-group users-mgmt-search">
            <span className="inp-icon">🔍</span>
            <input
              className="inp"
              value={query}
              onChange={e => setQuery(e.target.value)}
              placeholder={`Search ${tab}...`}
            />
          </div>
          <button className="users-mgmt-filter-btn" onClick={clearPrincipals}>Clear People</button>
          <button className="users-mgmt-filter-btn" onClick={clearApps}>Clear Apps</button>
        </div>
      </div>

      <div className="users-mgmt-layout">
        <div className="users-mgmt-main">
          <div className="card users-mgmt-panel">
            <div className="users-mgmt-panel-head">
              <div>
                <span className="section-label" style={{ marginBottom: 0 }}>Directory Results</span>
                <div className="users-mgmt-panel-title">{tab === 'users' ? 'Available users' : 'Available groups'}</div>
              </div>
              <div className="users-mgmt-toolbar-meta">
                <label className="chk-label">
                  <input
                    ref={selectAllResultsRef}
                    type="checkbox"
                    checked={allResultsSelected}
                    onChange={e => toggleAllResults(e.target.checked)}
                  />
                  Select visible
                </label>
                <span className="badge badge-gray">{selectedResultCount} selected</span>
              </div>
            </div>

            <div className="tbl-wrap users-mgmt-table-wrap">
              <table className="tbl users-mgmt-table">
                <thead>
                  <tr>
                    <th className="chk-col"></th>
                    <th>Name</th>
                    <th>{tab === 'users' ? 'Email / Username' : 'Email / Id'}</th>
                    <th>Type</th>
                    <th>Status</th>
                  </tr>
                </thead>
                <tbody>
                  {searching && (
                    <tr><td colSpan={5} className="tbl-empty"><span className="spinner spinner-dark" /></td></tr>
                  )}
                  {!searching && results.length === 0 && (
                    <tr><td colSpan={5} className="tbl-empty">No results found</td></tr>
                  )}
                  {!searching && results.map(item => {
                    const sub = item.userPrincipalName || item.mail || item.id
                    const sel = selectedPrincipals.some(s => s.id === item.id)
                    return (
                      <tr key={item.id} className={sel ? 'users-mgmt-row-selected' : ''}>
                        <td className="chk-col">
                          <input
                            type="checkbox"
                            checked={sel}
                            onChange={() => togglePrincipal(item)}
                          />
                        </td>
                        <td>{item.displayName}</td>
                        <td>{sub}</td>
                        <td>
                          <span className={`badge ${tab === 'groups' ? 'badge-purple' : 'badge-blue'}`}>{tab === 'groups' ? 'Group' : 'User'}</span>
                        </td>
                        <td>{sel ? 'Selected' : 'Available'}</td>
                      </tr>
                    )
                  })}
                </tbody>
              </table>
            </div>
          </div>

          <div className="card users-mgmt-panel">
            <div className="users-mgmt-panel-head">
              <div>
                <span className="section-label" style={{ marginBottom: 0 }}>Applications</span>
                <div className="users-mgmt-panel-title">Litera app assignments</div>
              </div>
              <div className="users-mgmt-toolbar-meta">
                <div className="inp-group users-mgmt-search users-mgmt-search-sm">
                  <span className="inp-icon">🔎</span>
                  <input
                    className="inp"
                    value={appQuery}
                    onChange={e => setAppQuery(e.target.value)}
                    placeholder="Filter apps"
                  />
                </div>
                <label className="chk-label">
                  <input
                    ref={selectAllAppsRef}
                    type="checkbox"
                    checked={allAppsSelected}
                    onChange={e => toggleAllApps(e.target.checked)}
                  />
                  Select all
                </label>
              </div>
            </div>

            <div className="tbl-wrap users-mgmt-table-wrap">
              <table className="tbl users-mgmt-table">
                <thead>
                  <tr>
                    <th className="chk-col"></th>
                    <th>Application</th>
                    <th>Service Principal Id</th>
                    <th>Selection</th>
                  </tr>
                </thead>
                <tbody>
                  {filteredApps.length === 0 && (
                    <tr><td colSpan={4} className="tbl-empty">No Litera apps found</td></tr>
                  )}
                  {filteredApps.map(app => {
                    const checked = selectedAppIds.includes(app.id)
                    return (
                      <tr key={app.id} className={checked ? 'users-mgmt-row-selected' : ''}>
                        <td className="chk-col">
                          <input
                            type="checkbox"
                            checked={checked}
                            onChange={() => toggleApp(app.id)}
                          />
                        </td>
                        <td>{app.displayName}</td>
                        <td className="text-xs" style={{ fontFamily: 'DM Mono, monospace' }}>{app.id}</td>
                        <td>{checked ? 'Selected' : 'Not selected'}</td>
                      </tr>
                    )
                  })}
                </tbody>
              </table>
            </div>
          </div>
        </div>

        <div className="users-mgmt-side">
          <div className="card users-mgmt-panel users-mgmt-summary-panel users-mgmt-sticky">
            <span className="section-label">Action Summary</span>
            <div className="users-mgmt-panel-title">{mode === 'assign' ? 'Ready to assign' : 'Ready to remove'}</div>
            <p className="users-mgmt-summary-copy">
              {mode === 'assign'
                ? 'Selected users or groups will be added to every selected Litera app.'
                : 'Existing assignments for the selected users or groups will be removed from every selected Litera app.'}
            </p>

            <div className="users-mgmt-stat-strip">
              <div className="users-mgmt-stat-box">
                <span>People</span>
                <strong>{selectedPrincipals.length}</strong>
              </div>
              <div className="users-mgmt-stat-box">
                <span>Apps</span>
                <strong>{selectedAppIds.length}</strong>
              </div>
              <div className="users-mgmt-stat-box">
                <span>History</span>
                <strong>{history.length}</strong>
              </div>
            </div>

            <button className="btn btn-primary users-mgmt-run-btn" onClick={runAction} disabled={running}>
              {running
                ? <><span className="spinner" /> Processing…</>
                : (mode === 'assign' ? 'Assign Selected' : 'Remove Selected')}
            </button>

            <div className="users-mgmt-mini-panel">
              <div className="users-mgmt-mini-head">
                <span className="section-label">Selected People</span>
                <span>{selectedPrincipals.length}</span>
              </div>
              <div className="users-mgmt-chip-list">
                {selectedPrincipals.length === 0 && <div className="users-mgmt-empty-inline">No users or groups selected</div>}
                {selectedPrincipals.map(s => (
                  <span key={s.id} className="chip">
                    {s.name}
                    <button type="button" className="chip-remove" onClick={() => togglePrincipal({ id: s.id, displayName: s.name, mail: s.emailOrId }, s.type === 'Group' ? 'groups' : 'users')}>×</button>
                  </span>
                ))}
              </div>
            </div>

            <div className="users-mgmt-mini-panel">
              <div className="users-mgmt-mini-head">
                <span className="section-label">Selected Apps</span>
                <span>{selectedAppIds.length}</span>
              </div>
              <div className="users-mgmt-chip-list">
                {selectedAppIds.length === 0 && <div className="users-mgmt-empty-inline">No apps selected</div>}
                {selectedAppIds.map(appId => {
                  const app = apps.find(item => item.id === appId)
                  return (
                    <span key={appId} className="chip">
                      {app?.displayName || appId}
                      <button type="button" className="chip-remove" onClick={() => toggleApp(appId)}>×</button>
                    </span>
                  )
                })}
              </div>
            </div>
          </div>
        </div>
      </div>

      <div className="card users-mgmt-history">
        <div className="flex items-center justify-between mb-2">
          <div>
            <span className="section-label" style={{ marginBottom: 0 }}>Activity</span>
            <div className="users-mgmt-panel-title users-mgmt-history-title">Assignment history</div>
          </div>
          <div className="users-mgmt-toolbar-meta">
            <span className="badge badge-green">{historySuccessCount} success</span>
            <span className="badge badge-red">{historyFailureCount} failed</span>
          </div>
          <button className="btn btn-secondary btn-sm" onClick={exportHistory}>Export CSV</button>
        </div>
        <div className="tbl-wrap" style={{ maxHeight: 320, overflowY: 'auto' }}>
          <table className="tbl">
            <thead>
              <tr>
                <th>Time</th>
                <th>Action</th>
                <th>App Name</th>
                <th>Principal</th>
                <th>Type</th>
                <th>Status</th>
              </tr>
            </thead>
            <tbody>
              {history.length === 0 && (
                <tr><td colSpan={6} className="tbl-empty">No history yet</td></tr>
              )}
              {history.slice().reverse().map((h, idx) => (
                <tr key={`${h.time}-${h.appId}-${h.principalId}-${idx}`}>
                  <td className="text-xs">{h.time}</td>
                  <td>{h.action}</td>
                  <td>{h.appName}</td>
                  <td>
                    <div>{h.principalName}</div>
                    <div className="text-xs" style={{ fontFamily: 'DM Mono, monospace' }}>{h.emailOrId}</div>
                  </td>
                  <td>{h.principalType}</td>
                  <td>
                    <span className={`badge ${h.status === 'Success' ? 'badge-green' : 'badge-red'}`}>
                      {h.status}
                    </span>
                  </td>
                </tr>
              ))}
            </tbody>
          </table>
        </div>
      </div>
    </div>
  )
}
