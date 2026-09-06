export interface AuthStatus {
  authenticated: boolean
  orgName?: string
  tenantId?: string
  lastSync?: string
}
