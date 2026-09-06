import 'express-session'

export interface AuthSessionData {
  accessToken: string
  refreshToken?: string | null
  expiresAt?: number
  tenantId: string
  clientId: string
  orgName: string
  orgId?: string
  userName?: string
  lastSync?: string | null
  connectedAt?: string
}

declare module 'express-session' {
  interface SessionData {
    auth?: AuthSessionData | null
    oauthState?: string
    oauthTenant?: string
    returnTo?: string
    csrfToken?: string
    enablementAdmin?: {
      loggedIn: boolean
      username: string
      loggedInAt?: string
    } | null
  }
}
