import { useState, useCallback } from 'react'
import graphClient from '../services/graphClient'
import AssignModal, { type AssignSelection } from '../components/AssignModal'
import { downloadCsv } from '../utils/exportHelpers'
import type { GraphApp } from '../types/graph'
import type { LogFn, ToastFn } from '../types/ui'

const DEFAULT_URLS = {
  Outlook: 'https://marketplace.microsoft.com/en-us/product/office/wa200008417',
  Word: 'https://marketplace.microsoft.com/en-us/product/office/wa200008417'
}

interface DeploymentsProps {
  log: LogFn
  toast: ToastFn
  apps: GraphApp[]
}

type DeployTarget = AssignSelection & {
  id: string
  name: string
  type: string
  emailOrId?: string
}

type StatusItem = {
  name: string
  state: string
  time?: string
  ok?: boolean
}

export default function Deployments ({ log, toast, apps }: DeploymentsProps) {
  const [product, setProduct] = useState('Outlook')
  const [source, setSource] = useState('appsource')
  const [method, setMethod] = useState('prepare')
  const [sourceUrl, setSourceUrl] = useState(DEFAULT_URLS['Outlook'])
  const [manifestInput, setManifestInput] = useState('')
  const [targets, setTargets] = useState<DeployTarget[]>([])
  const [copyTargetsEnabled, setCopyTargetsEnabled] = useState(false)
  const [copySourceAppId, setCopySourceAppId] = useState('')
  const [statusItems, setStatusItems] = useState<StatusItem[]>([])
  const [loadingStatus, setLoadingStatus] = useState(false)
  const [showAssignModal, setShowAssignModal] = useState(false)

  const changeProduct = (p: string) => {
    setProduct(p)
    setSourceUrl(DEFAULT_URLS[p] || '')
  }

  const addStatusItem = (name: string, state: string, ok = false) => {
    setStatusItems(prev => [{
      name, state,
      time: new Date().toLocaleTimeString(),
      ok
    }, ...prev])
  }

  const loadDeployments = async () => {
    if (!graphClient.isAuthenticated) { toast('Sign in first', 'error'); return }
    setLoadingStatus(true)
    log('Loading deployment status…')
    try {
      const res = await graphClient.getDeployments()
      const items = res?.value || []
      if (!items.length) {
        setStatusItems([{ name: 'No deployments found', state: 'This API may not be available for your tenant.', time: new Date().toLocaleTimeString() }])
        log('No deployments found (API may not be supported in this tenant).')
        return
      }
      setStatusItems(items.map(d => ({
        name: d.displayName || d.description || d.id || '(unnamed)',
        state: d.state || d.status || d.lifecycleStatus || 'Status unavailable',
        time: new Date().toLocaleTimeString()
      })))
      log(`Loaded ${items.length} deployment(s).`)
    } catch (e) {
      log(`Deployments API error: ${e.message}`)
      setStatusItems([{
        name: 'Deployment status unavailable',
        state: e.message + ' — Try "M365 Admin Center" method.',
        time: new Date().toLocaleTimeString()
      }])
    } finally {
      setLoadingStatus(false)
    }
  }

  const copyTargetsFromApp = async () => {
    if (!copySourceAppId) { toast('Select a source app', 'error'); return }
    try {
      const raw = await graphClient.getAssignments(copySourceAppId)
      const enriched = await graphClient.enrichAssignments(raw)
      const newTargets = enriched.map(a => ({
        id: a.principalId,
        name: a.principalDisplayName,
        type: a.principalType,
        emailOrId: a.emailOrId
      }))
      setTargets(prev => {
        const existing = new Set(prev.map(t => t.id))
        return [...prev, ...newTargets.filter(t => !existing.has(t.id))]
      })
      toast(`Copied ${newTargets.length} target(s)`, 'success')
      log(`Copied ${newTargets.length} targets from source app.`)
    } catch (e) {
      toast(e.message, 'error')
    }
  }

  const removeTarget = (id: string) => {
    setTargets(prev => prev.filter(t => t.id !== id))
  }

  const submitDeploy = () => {
    const src = source === 'manifest' ? manifestInput : sourceUrl
    if (!src.trim()) { toast('Source URL is required', 'error'); return }
    if (!targets.length) { toast('Add deployment targets first', 'error'); return }
    log(`DEPLOY REQUEST — Product: ${product} | Source: ${source} | URL: ${src} | Targets: ${targets.length}`)
    addStatusItem(
      `🚀 Deployment — ${product} (${source})`,
      `${targets.length} targets · ${new Date().toLocaleTimeString()} · ${method === 'integrated' ? 'M365 Admin Center' : 'Prepared via Portal'}`,
      true
    )
    toast('Deployment request logged', 'success')
  }

  const confirmDeployAssign = useCallback((selected: AssignSelection[]) => {
    setShowAssignModal(false)
    setTargets(prev => {
      const existing = new Set(prev.map(t => t.id))
      return [...prev, ...selected.filter(s => !existing.has(s.id))]
    })
    toast(`Added ${selected.length} target(s)`, 'success')
  }, [toast])

  return (
    <div className="page-container page-enter">
      <div className="page-header">
        <div>
          <div className="page-title">Add-in Deployment</div>
          <div className="page-subtitle">Deploy and monitor Litera One add-ins for Outlook and Word</div>
        </div>
        <button className="btn btn-secondary" onClick={loadDeployments} disabled={loadingStatus}>
          {loadingStatus ? <><span className="spinner" /> Loading…</> : '↻ Check Status'}
        </button>
      </div>

      <div style={{ display: 'grid', gridTemplateColumns: '1fr 1fr', gap: 18, alignItems: 'start' }}>
        {/* ── Left: Config ── */}
        <div className="card">
          {/* Product */}
          <span className="section-label">Add-in Product</span>
          <div className="radio-group mb-3">
            {['Outlook', 'Word'].map(p => (
              <label key={p} className="radio-opt">
                <input type="radio" name="product" value={p} checked={product === p}
                  onChange={() => changeProduct(p)} />
                {p === 'Outlook' ? '✉ ' : '📄 '} Litera One for {p}
              </label>
            ))}
          </div>

          {/* Source */}
          <span className="section-label">Deployment Source</span>
          <div className="radio-group mb-3">
            <label className="radio-opt">
              <input type="radio" name="source" value="appsource" checked={source === 'appsource'}
                onChange={() => setSource('appsource')} />
              Microsoft AppSource (URL)
            </label>
            <label className="radio-opt">
              <input type="radio" name="source" value="manifest" checked={source === 'manifest'}
                onChange={() => setSource('manifest')} />
              Manifest File (XML / URL)
            </label>
          </div>

          {source === 'appsource' && (
            <div className="form-row">
              <label className="label">AppSource URL</label>
              <input className="inp inp-mono" value={sourceUrl} onChange={e => setSourceUrl(e.target.value)} />
            </div>
          )}
          {source === 'manifest' && (
            <div className="form-row">
              <label className="label">Manifest URL or File Path</label>
              <input className="inp" value={manifestInput} onChange={e => setManifestInput(e.target.value)}
                placeholder="https://…/manifest.xml" />
            </div>
          )}

          {/* Method */}
          <span className="section-label">Deployment Method</span>
          <div className="radio-group mb-3">
            <label className="radio-opt">
              <input type="radio" name="method" value="prepare" checked={method === 'prepare'}
                onChange={() => setMethod('prepare')} />
              Prepare via this tool
            </label>
            <label className="radio-opt">
              <input type="radio" name="method" value="integrated" checked={method === 'integrated'}
                onChange={() => setMethod('integrated')} />
              Deploy via M365 Admin Center
            </label>
          </div>

          {method === 'integrated' && (
            <div className="info-box mb-3">
              <strong>Integrated Apps (manual):</strong> Use the M365 admin center to deploy the add-in.
              Export your targets to CSV and paste them in the admin center assignment dialog.
              <div className="flex gap-2 mt-2">
                <a className="btn btn-secondary btn-sm"
                  href="https://admin.microsoft.com/adminportal/home#/Settings/IntegratedApps"
                  target="_blank" rel="noreferrer">
                  ↗ Open Integrated Apps
                </a>
                <button className="btn btn-secondary btn-sm"
                  onClick={() => downloadCsv([['Name','Email/ID','Type'], ...targets.map(t => [t.name, t.emailOrId || t.id, t.type])], 'deploy_targets.csv')}>
                  Export CSV
                </button>
              </div>
            </div>
          )}

          <hr className="divider" />

          {/* Targets */}
          <div className="flex items-center justify-between mb-2">
            <span className="section-label" style={{ marginBottom: 0 }}>
              Targets ({targets.length})
            </span>
            <div className="flex gap-2">
              <button className="btn btn-secondary btn-sm" onClick={() => setShowAssignModal(true)}>
                ＋ Add Users/Groups
              </button>
              {targets.length > 0 && (
                <button className="btn btn-danger btn-xs" onClick={() => setTargets([])}>Clear</button>
              )}
            </div>
          </div>

          {/* Copy from app */}
          <label className="chk-label mb-2">
            <input type="checkbox" checked={copyTargetsEnabled} onChange={e => setCopyTargetsEnabled(e.target.checked)} />
            Copy targets from a Litera app
          </label>
          {copyTargetsEnabled && (
            <div className="flex gap-2 mb-3">
              <select className="inp inp-sm" value={copySourceAppId} onChange={e => setCopySourceAppId(e.target.value)} style={{ flex: 1 }}>
                <option value="">— Select app —</option>
                {(apps || []).map(a => <option key={a.id} value={a.id}>{a.displayName}</option>)}
              </select>
              <button className="btn btn-secondary btn-sm" onClick={copyTargetsFromApp}>Copy</button>
            </div>
          )}

          {/* Target table */}
          <div className="tbl-wrap mb-3" style={{ maxHeight: 200, overflowY: 'auto' }}>
            <table className="tbl">
              <thead>
                <tr><th>Name</th><th>Type</th><th style={{ width: 32 }} /></tr>
              </thead>
              <tbody>
                {targets.length === 0 && (
                  <tr><td colSpan={3} className="tbl-empty">No targets selected</td></tr>
                )}
                {targets.map(t => (
                  <tr key={t.id}>
                    <td>
                      <div style={{ color: 'var(--text)', fontWeight: 600, fontSize: 13 }}>{t.name}</div>
                      <div className="text-xs" style={{ fontFamily: 'DM Mono, monospace' }}>{t.emailOrId || t.id}</div>
                    </td>
                    <td><span className={`badge ${t.type === 'User' ? 'badge-blue' : 'badge-purple'}`}>{t.type}</span></td>
                    <td>
                      <button className="btn btn-danger btn-xs" onClick={() => removeTarget(t.id)}>×</button>
                    </td>
                  </tr>
                ))}
              </tbody>
            </table>
          </div>

          <button className="btn btn-primary w-full" onClick={submitDeploy}>
            ☁ {method === 'integrated' ? 'Log Deployment Request' : 'Prepare Deployment'}
          </button>
        </div>

        {/* ── Right: Status ── */}
        <div className="card">
          <div className="card-header">Deployment Status</div>
          {loadingStatus && (
            <div style={{ textAlign: 'center', padding: 24 }}>
              <span className="spinner spinner-dark" />
            </div>
          )}
          {!loadingStatus && statusItems.length === 0 && (
            <div className="empty-state">
              <div className="empty-icon">☁</div>
              <div>Click "Check Status" to load deployments</div>
              <div className="text-xs mt-2">Requires admin/officeConfiguration Graph access</div>
            </div>
          )}
          {statusItems.map((item, i) => (
            <div key={i} className="deploy-status-item">
              <div className="deploy-status-name">{item.name}</div>
              <div className="deploy-status-state">{item.state}</div>
              {item.time && <div className="text-xs mt-2" style={{ marginTop: 4 }}>{item.time}</div>}
            </div>
          ))}
        </div>
      </div>

      {showAssignModal && (
        <AssignModal
          title="Add Deployment Targets"
          onConfirm={confirmDeployAssign}
          onClose={() => setShowAssignModal(false)}
        />
      )}
    </div>
  )
}
