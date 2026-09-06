// ─── Credential Store ─────────────────────────────────────────────────────────
// Persists auth credentials in localStorage (user opt-in only)
import type { EnablementDataPayload } from '../types/enablement'

const KEY = 'lo_creds_v2'
const ENABLEMENT_KEY = 'lo_enablement'

export interface StoredCredentials {
  tenantId: string
  clientId: string
  clientSecret: string
}

export function saveCredentials (tenantId: string, clientId: string, clientSecret: string) {
  try {
    localStorage.setItem(KEY, JSON.stringify({ tenantId, clientId, clientSecret }))
  } catch { /* storage unavailable */ }
}

export function loadCredentials (): StoredCredentials | null {
  try {
    const raw = localStorage.getItem(KEY)
    return raw ? JSON.parse(raw) as StoredCredentials : null
  } catch { return null }
}

export function clearCredentials () {
  try { localStorage.removeItem(KEY) } catch { /* noop */ }
}

// Enablement hub – load custom JSON from localStorage if user uploaded one
export function saveEnablementData (data: EnablementDataPayload) {
  try { localStorage.setItem(ENABLEMENT_KEY, JSON.stringify(data)) } catch { /* noop */ }
}

export function loadEnablementData (): EnablementDataPayload | null {
  try {
    const raw = localStorage.getItem(ENABLEMENT_KEY)
    return raw ? JSON.parse(raw) as EnablementDataPayload : null
  } catch { return null }
}
