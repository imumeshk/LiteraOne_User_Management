import { useState, useRef, useCallback } from 'react'
import graphClient from '../services/graphClient'
import AssignModal, { type AssignSelection } from './AssignModal'

interface ScimModalProps {
  onClose: () => void
  onCreated?: () => void
}

export default function ScimModal ({ onClose, onCreated }: ScimModalProps) {
  const [step, setStep] = useState(1)
  const [appName, setAppName] = useState('Litera One SCIM')
  const [scimUrl, setScimUrl] = useState('')
  const [scimToken, setScimToken] = useState('')
  const [showToken, setShowToken] = useState(false)
  const [assignSelected, setAssignSelected] = useState<AssignSelection[]>([])
  const [showAssignPicker, setShowAssignPicker] = useState(false)
  const [logLines, setLogLines] = useState<string[]>(['Waiting to start…'])
  const [progress, setProgress] = useState(0)
  const [creating, setCreating] = useState(false)
  const [done, setDone] = useState(false)
  const cancelled = useRef(false)
  const logRef = useRef<HTMLDivElement | null>(null)

  const addLog = useCallback((msg: string) => {
    const line = `${new Date().toLocaleTimeString()} — ${msg}`
    setLogLines(prev => {
      const next = [...prev, line]
      setTimeout(() => { if (logRef.current) logRef.current.scrollTop = logRef.current.scrollHeight }, 50)
      return next
    })
  }, [])

  const validateStep1 = () => {
    if (!appName.trim()) return 'Application name is required.'
    if (!scimUrl.trim()) return 'SCIM tenant URL is required.'
    if (!scimToken.trim()) return 'SCIM secret token is required.'
    return null
  }

  const handleNext = async () => {
    if (step === 1) {
      const err = validateStep1()
      if (err) { alert(err); return }
      setStep(2)
    } else if (step === 2) {
      setStep(3)
      await createApp()
    }
  }

  const createApp = async () => {
    cancelled.current = false
    setCreating(true)
    setProgress(5)
    addLog(`Starting SCIM app creation: ${appName}`)

    try {
      // Step A: Create app + SP
      addLog('Creating application and service principal…')
      const { spId } = await graphClient.createScimApp(appName)
      if (!spId) throw new Error('Failed to obtain service principal ID')
      addLog(`✓ Service Principal ID: ${spId}`)
      setProgress(30)
      if (cancelled.current) { addLog('Cancelled.'); return }

      // Step B: Assign users/groups
      if (assignSelected.length) {
        addLog(`Assigning ${assignSelected.length} principal(s)…`)
        for (const p of assignSelected) {
          if (cancelled.current) break
          try {
            await graphClient.addAssignment(spId, p.id, p.type)
            addLog(`  ✓ Assigned: ${p.name}`)
          } catch (e) {
            addLog(`  ⚠ Assign note for ${p.name}: ${e.message}`)
          }
        }
      }
      setProgress(55)
      if (cancelled.current) { addLog('Cancelled.'); return }

      // Step C: Create sync job
      addLog('Creating SCIM synchronization job…')
      const job = await graphClient.createSyncJob(spId)
      if (job?.id) {
        addLog(`✓ Sync job created: ${job.id}`)
      } else {
        addLog('⚠ Sync job creation may require manual setup in Azure Portal')
      }
      setProgress(70)

      // Step D: Set SCIM secrets
      if (scimUrl && scimToken && job?.id) {
        addLog('Setting SCIM provisioning secrets…')
        const secrets = [
          { key: 'BaseAddress', value: scimUrl },
          { key: 'SecretToken', value: scimToken }
        ]
        const ok = await graphClient.setSyncSecrets(spId, secrets)
        addLog(ok ? '✓ SCIM secrets saved' : '⚠ Could not save secrets via API — set them manually in Azure Portal')
      }
      setProgress(90)

      // Step E: Hide app from MyApps (recommended for SCIM)
      try {
        await graphClient.setSpTags(spId, ['HideApp', 'WindowsAzureActiveDirectoryIntegratedApp'])
        addLog('✓ App hidden from My Apps (recommended for SCIM)')
      } catch { /* non-critical */ }

      setProgress(100)
      addLog(`\n✅ SCIM app "${appName}" created successfully!`)
      addLog('Refresh the Entra Apps tab to see the new app.')
      setDone(true)
      onCreated?.()
    } catch (e) {
      addLog(`\n❌ Error: ${e.message}`)
      setProgress(0)
    } finally {
      setCreating(false)
    }
  }

  return (
    <>
      <div className="modal-overlay" onClick={e => e.target === e.currentTarget && !creating && onClose()}>
        <div className="modal modal-lg">
          <div className="modal-title">Create Litera One SCIM App</div>

          <div className="steps">
            <div className={`step ${step === 1 ? 'active' : step > 1 ? 'done' : ''}`}>1. App Details</div>
            <div className={`step ${step === 2 ? 'active' : step > 2 ? 'done' : ''}`}>2. Assignments</div>
            <div className={`step ${step === 3 ? 'active' : ''}`}>3. Create</div>
          </div>

          {/* ── Step 1 ── */}
          {step === 1 && (
            <div>
              <div className="form-row">
                <label className="label">Application Name <span className="required">*</span></label>
                <input className="inp" value={appName} onChange={e => setAppName(e.target.value)} placeholder="e.g. Litera One SCIM" autoFocus />
              </div>
              <div className="form-row">
                <label className="label">SCIM Tenant URL <span className="required">*</span></label>
                <input className="inp inp-mono" value={scimUrl} onChange={e => setScimUrl(e.target.value)} placeholder="https://scim.litera.com/v2/{tenantId}" />
                <div className="hint">Provided by Litera — the SCIM endpoint for your Litera tenant</div>
              </div>
              <div className="form-row">
                <label className="label">Secret Token <span className="required">*</span></label>
                <div className="flex gap-2">
                  <input
                    className="inp inp-mono"
                    type={showToken ? 'text' : 'password'}
                    value={scimToken}
                    onChange={e => setScimToken(e.target.value)}
                    placeholder="Bearer token provided by Litera"
                  />
                  <button className="btn btn-secondary btn-sm" onClick={() => setShowToken(s => !s)} tabIndex={-1}>
                    {showToken ? '🙈' : '👁'}
                  </button>
                </div>
              </div>
              <div className="info-box">
                This wizard will create an Enterprise Application in your Azure AD, configure SCIM provisioning, and optionally assign initial users/groups.
                Admin consent for required permissions will be applied automatically.
              </div>
            </div>
          )}

          {/* ── Step 2 ── */}
          {step === 2 && (
            <div>
              <div className="section-label">Initial Assignments (optional)</div>
              <p className="text-sm mb-3" style={{ color: 'var(--text3)' }}>
                Add users or groups to the SCIM app now. You can also add them later via the Entra Apps tab.
              </p>
              <button className="btn btn-secondary mb-2" onClick={() => setShowAssignPicker(true)}>
                ＋ Add Users / Groups
              </button>
              {assignSelected.length > 0 && (
                <div className="flex flex-wrap gap-1 mt-2">
                  {assignSelected.map(s => (
                    <span key={s.id} className="chip">
                      <span className={`badge ${s.type === 'Group' ? 'badge-purple' : 'badge-blue'}`} style={{ marginRight: 4, padding: '1px 5px' }}>{s.type}</span>
                      {s.name}
                      <button className="chip-remove" onClick={() => setAssignSelected(prev => prev.filter(p => p.id !== s.id))}>×</button>
                    </span>
                  ))}
                </div>
              )}
              {assignSelected.length === 0 && (
                <div className="empty-state" style={{ padding: '20px 0' }}>
                  <div style={{ fontSize: 13, color: 'var(--text3)' }}>No assignments yet — you can skip this step</div>
                </div>
              )}
            </div>
          )}

          {/* ── Step 3 ── */}
          {step === 3 && (
            <div>
              <div className="creation-log" ref={logRef}>
                {logLines.join('\n')}
              </div>
              {(creating || done) && (
                <div className="mt-2">
                  <div className="text-xs mb-1" style={{ color: 'var(--text3)' }}>
                    {done ? 'Completed' : `Creating app…`}
                  </div>
                  <div className="progress-wrap"><div className="progress-bar" style={{ width: progress + '%' }} /></div>
                </div>
              )}
            </div>
          )}

          <div className="modal-footer">
            {step < 3 && (
              <button className="btn btn-secondary" onClick={onClose}>Cancel</button>
            )}
            {step === 3 && (
              <>
                {creating && (
                  <button className="btn btn-danger btn-sm" onClick={() => { cancelled.current = true }}>Cancel</button>
                )}
                <button className="btn btn-secondary" onClick={onClose} disabled={creating && !cancelled.current}>
                  {done ? 'Close' : 'Dismiss'}
                </button>
              </>
            )}
            {step === 2 && (
              <button className="btn btn-secondary" onClick={() => setStep(1)}>← Back</button>
            )}
            {step < 3 && (
              <button className="btn btn-primary" onClick={handleNext}>
                {step === 2 ? '🚀 Create App' : 'Next →'}
              </button>
            )}
          </div>
        </div>
      </div>

      {showAssignPicker && (
        <AssignModal
          title="Add Initial Assignments"
          onConfirm={(sel: AssignSelection[]) => {
            setAssignSelected(prev => {
              const merged = [...prev]
              for (const s of sel) { if (!merged.find(p => p.id === s.id)) merged.push(s) }
              return merged
            })
            setShowAssignPicker(false)
          }}
          onClose={() => setShowAssignPicker(false)}
        />
      )}
    </>
  )
}
