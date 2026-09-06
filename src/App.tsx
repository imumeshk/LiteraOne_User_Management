import { useState, useCallback, useEffect } from 'react'
import Sidebar from './components/Sidebar'
import ConsolePanel from './components/ConsolePanel'
import Notifications from './components/Notifications'
import ToastContainer from './components/Toast'
import Dashboard from './pages/Dashboard'
import EntraApps from './pages/EntraApps'
import Deployments from './pages/Deployments'
import EnablementHub from './pages/EnablementHub'
import LiteraUsersManagement from './pages/LiteraUsersManagement'
import { useToast } from './hooks/useToast'
import { useLogger } from './hooks/useLogger'
import type { AuthStatus } from './types/auth'
import type { Page } from './types/ui'

export default function App () {
  const [page, setPage] = useState<Page>('dashboard')
  const [consoleOpen, setConsoleOpen] = useState(false)

  // Auth state
  const [connected, setConnected] = useState(false)
  const [orgName, setOrgName] = useState('')
  const [tenantId, setTenantId] = useState('')
  const [lastSync, setLastSync] = useState('—')

  // Shared app list (passed to Deployments for copy-targets feature)
  const [apps, setApps] = useState<unknown[]>([])

  const { toasts, toast } = useToast()
  const { logs, log, clearLogs, exportLogs } = useLogger()

  useEffect(() => {
    let cancelled = false
    const loadStatus = async () => {
      try {
        const res = await fetch('/auth/status', { credentials: 'same-origin' })
        const data: AuthStatus = await res.json()
        if (cancelled) return
        setConnected(!!data.authenticated)
        setOrgName(data.orgName || '')
        setTenantId(data.tenantId || '')
        setLastSync(data.lastSync ? new Date(data.lastSync).toLocaleString() : '—')
      } catch {
        if (cancelled) return
        setConnected(false)
      }
    }
    loadStatus()
    return () => { cancelled = true }
  }, [])

  useEffect(() => {
    if (!connected && page === 'users-management') setPage('dashboard')
  }, [connected, page])

  // ── Sign Out ──────────────────────────────────────────────────────────────
  const handleSignOut = useCallback(() => {
    if (!confirm('Sign out and clear session?')) return
    const form = document.createElement('form')
    form.method = 'POST'
    form.action = '/auth/logout'
    document.body.appendChild(form)
    form.submit()
  }, [toast, log])

  const handleAuthClick = () => {
    if (connected) handleSignOut()
    else window.location.href = '/auth/login'
  }

  const navigate = (p: Page) => {
    if (p === 'users-management' && !connected) return
    setPage(p)
  }

  const pageProps = { log, toast }

  return (
    <div className="app-shell">
      <Sidebar
        activePage={page}
        onNav={navigate}
        connected={connected}
        orgName={orgName}
        tenantId={tenantId}
        onAuthClick={handleAuthClick}
        consoleOpen={consoleOpen}
        onToggleConsole={() => setConsoleOpen(o => !o)}
      />

      <div className="main-content">
        <Notifications logs={logs} />

        {/* Pages */}
        {page === 'dashboard' && (
          <Dashboard
            connected={connected}
            orgName={orgName}
            tenantId={tenantId}
            lastSync={lastSync}
            onNav={navigate}
          />
        )}
        {page === 'apps' && (
          <EntraApps {...pageProps} onAppsLoaded={setApps} />
        )}
        {page === 'deploy' && (
          <Deployments {...pageProps} apps={apps} />
        )}
        {page === 'enablement' && (
          <EnablementHub {...pageProps} />
        )}
        {page === 'users-management' && (
          <LiteraUsersManagement connected={connected} {...pageProps} />
        )}

        {/* Console */}
        <ConsolePanel
          logs={logs}
          open={consoleOpen}
          onClear={clearLogs}
          onExport={exportLogs}
        />
      </div>
      <ToastContainer toasts={toasts} />
    </div>
  )
}
