// ─── Microsoft Graph API Client ──────────────────────────────────────────────
// Uses Client Credentials (OAuth 2.0) flow: Client ID + Client Secret → Bearer token

const GRAPH_BASE = 'https://graph.microsoft.com/v1.0'
const GRAPH_BETA = 'https://graph.microsoft.com/beta'
const NON_GALLERY_TEMPLATE_ID = '8adf8e6e-67b2-4cf2-a259-e3dc5476c621'

class GraphClient {
  constructor () {
    this.token = null
    this.tenantId = null
    this.clientId = null
  }

  // ── Authenticate ─────────────────────────────────────────────────────────
  async authenticate (tenantId, clientId, clientSecret) {
    const body = new URLSearchParams({
      grant_type: 'client_credentials',
      client_id: clientId,
      client_secret: clientSecret,
      scope: 'https://graph.microsoft.com/.default'
    })

    const res = await fetch(
      `https://login.microsoftonline.com/${tenantId}/oauth2/v2.0/token`,
      {
        method: 'POST',
        headers: { 'Content-Type': 'application/x-www-form-urlencoded' },
        body: body.toString()
      }
    )

    const data = await res.json()
    if (!res.ok) {
      const msg = data.error_description || data.error || `HTTP ${res.status}`
      throw new Error(msg)
    }

    this.token = data.access_token
    this.tenantId = tenantId
    this.clientId = clientId
    return data
  }

  signOut () {
    this.token = null
    this.tenantId = null
    this.clientId = null
  }

  get isAuthenticated () { return !!this.token }

  // ── Core request ─────────────────────────────────────────────────────────
  async request (method, path, body = null, { beta = false } = {}) {
    if (!this.token) throw new Error('Not authenticated')
    const base = beta ? GRAPH_BETA : GRAPH_BASE
    const url = path.startsWith('http') ? path : `${base}${path.startsWith('/') ? path : '/' + path}`

    const opts = {
      method,
      headers: {
        Authorization: `Bearer ${this.token}`,
        'Content-Type': 'application/json',
        ConsistencyLevel: 'eventual' // needed for $search queries
      }
    }
    if (body) opts.body = JSON.stringify(body)

    const res = await fetch(url, opts)
    if (res.status === 204) return true

    const data = await res.json()
    if (!res.ok) {
      const msg = data.error?.message || data.error_description || `HTTP ${res.status}`
      throw new Error(msg)
    }
    return data
  }

  get (path, opts) { return this.request('GET', path, null, opts) }
  post (path, body, opts) { return this.request('POST', path, body, opts) }
  patch (path, body, opts) { return this.request('PATCH', path, body, opts) }
  del (path, opts) { return this.request('DELETE', path, null, opts) }

  // ── Paginated GET (follows @odata.nextLink) ───────────────────────────────
  async getAll (path, opts = {}) {
    let results = []
    let url = path
    while (url) {
      const data = await this.request('GET', url, null, opts)
      if (data?.value) results.push(...data.value)
      url = data?.['@odata.nextLink'] ?? null
    }
    return results
  }

  // ── Service Principals ────────────────────────────────────────────────────
  async searchApps (query) {
    const q = query || 'Litera'
    return this.get(`/servicePrincipals?$filter=startswith(displayName,'${encodeURIComponent(q)}')&$select=id,displayName,appId,tags,accountEnabled`)
  }

  async getServicePrincipal (spId) {
    return this.get(`/servicePrincipals/${spId}?$select=id,displayName,appId,tags,accountEnabled,appRoles,appRoleAssignmentRequired`)
  }

  async getSpTags (spId) {
    try {
      return await this.get(`/servicePrincipals/${spId}?$select=id,appId,tags,servicePrincipalType,appOwnerOrganizationId`)
    } catch {
      return this.request('GET', `/servicePrincipals/${spId}?$select=id,appId,tags,servicePrincipalType,appOwnerOrganizationId`, null, { beta: true })
    }
  }

  async setSpTags (spId, tags) {
    const normalized = [...new Set(tags.filter(Boolean))]
    try {
      await this.patch(`/servicePrincipals/${spId}`, { tags: normalized })
      return true
    } catch {
      try {
        await this.request('PATCH', `/servicePrincipals/${spId}`, { tags: normalized }, { beta: true })
        return true
      } catch { return false }
    }
  }

  async setAppTagsByAppId (appId, tags) {
    try {
      const res = await this.get(`/applications(appId='${appId}')?$select=id,tags`)
      if (!res?.id) return false
      await this.patch(`/applications/${res.id}`, { tags: [...new Set(tags.filter(Boolean))] })
      return true
    } catch { return false }
  }

  async deleteServicePrincipal (spId) {
    return this.del(`/servicePrincipals/${spId}`)
  }

  // ── Assignments ───────────────────────────────────────────────────────────
  async getAssignments (spId) {
    return this.getAll(`/servicePrincipals/${spId}/appRoleAssignedTo?$top=999`)
  }

  async enrichAssignments (assignments) {
    const ids = [...new Set(
      assignments
        .filter(a => ['User', 'Group'].includes(a.principalType) && a.principalId)
        .map(a => a.principalId)
    )]
    const lookup = {}
    const CHUNK = 900
    for (let i = 0; i < ids.length; i += CHUNK) {
      const chunk = ids.slice(i, i + CHUNK)
      try {
        const res = await this.post('/directoryObjects/getByIds', { ids: chunk, types: ['user', 'group'] })
        for (const obj of (res?.value || [])) {
          lookup[obj.id] = obj.mail || obj.userPrincipalName || obj.id
        }
      } catch { /* partial enrichment ok */ }
    }
    return assignments.map(a => ({
      ...a,
      emailOrId: lookup[a.principalId] || a.principalId,
      isChecked: false
    }))
  }

  async addAssignment (spId, principalId, principalType) {
    // Resolve best appRoleId
    let appRoleId = '00000000-0000-0000-0000-000000000000'
    try {
      const sp = await this.get(`/servicePrincipals/${spId}?$select=appRoles,appRoleAssignmentRequired`)
      const memberType = principalType === 'Application' ? 'Application' : 'User'
      const roles = (sp?.appRoles || []).filter(r =>
        r.isEnabled && (r.allowedMemberTypes || []).includes(memberType)
      )
      if (roles.length) {
        const defaultRole = roles.find(r => !r.value) || roles[0]
        appRoleId = defaultRole.id
      } else if (sp?.appRoleAssignmentRequired) {
        throw new Error('No assignable app role found for ' + principalType)
      }
    } catch (e) {
      if (e.message.includes('No assignable')) throw e
    }
    return this.post(`/servicePrincipals/${spId}/appRoleAssignedTo`, {
      principalId, resourceId: spId, appRoleId
    })
  }

  async removeAssignment (spId, assignmentId) {
    return this.del(`/servicePrincipals/${spId}/appRoleAssignedTo/${assignmentId}`)
  }

  // ── Permissions / Consent ─────────────────────────────────────────────────
  async getAppPermissions (appId) {
    return this.get(`/applications?$filter=appId eq '${appId}'&$select=requiredResourceAccess`)
  }

  async getPermissionGrants (spId) {
    return this.get(`/oauth2PermissionGrants?$filter=clientId eq '${spId}'`)
  }

  async grantAdminConsent (sp) {
    const appQ = await this.get(`/applications?$filter=appId eq '${sp.appId}'&$select=requiredResourceAccess`)
    const app = appQ?.value?.[0]
    if (!app) throw new Error('Application registration not found')
    const results = []
    for (const rra of (app.requiredResourceAccess || [])) {
      const spR = await this.get(
        `/servicePrincipals?$filter=appId eq '${rra.resourceAppId}'&$select=id,displayName,oauth2PermissionScopes,appRoles`
      )
      const resSp = spR?.value?.[0]
      if (!resSp) continue
      const scopes = [], roles = []
      for (const ra of (rra.resourceAccess || [])) {
        if (ra.type === 'Scope') {
          const s = resSp.oauth2PermissionScopes?.find(x => x.id === ra.id)
          if (s) scopes.push(s.value)
        } else {
          roles.push({ id: ra.id, resourceId: resSp.id })
        }
      }
      if (scopes.length) {
        const existing = await this.get(
          `/oauth2PermissionGrants?$filter=clientId eq '${sp.id}' and resourceId eq '${resSp.id}'`
        )
        const grant = existing?.value?.[0]
        if (grant) {
          const merged = [...new Set([...grant.scope.split(' '), ...scopes])].join(' ')
          if (merged !== grant.scope) {
            await this.patch(`/oauth2PermissionGrants/${grant.id}`, { scope: merged })
          }
        } else {
          await this.post('/oauth2PermissionGrants', {
            clientId: sp.id, consentType: 'AllPrincipals',
            resourceId: resSp.id, scope: scopes.join(' ')
          })
        }
        results.push(`Delegated scopes: ${scopes.join(', ')} on ${resSp.displayName}`)
      }
      for (const role of roles) {
        try {
          await this.post(`/servicePrincipals/${sp.id}/appRoleAssignments`, {
            principalId: sp.id, resourceId: role.resourceId, appRoleId: role.id
          })
          results.push(`App role ${role.id} on ${resSp.displayName}`)
        } catch (e) { results.push(`Role note: ${e.message}`) }
      }
    }
    return results
  }

  // ── Search Users/Groups ───────────────────────────────────────────────────
  async searchPrincipals (type, query) {
    if (!query.trim()) {
      const f = type === 'groups'
        ? '/groups?$select=id,displayName,mail&$top=30'
        : '/users?$select=id,displayName,userPrincipalName,mail&$top=30'
      return this.get(f)
    }
    const fields = type === 'groups' ? 'id,displayName,mail' : 'id,displayName,userPrincipalName,mail'
    return this.get(`/${type}?$search="displayName:${query}"&$select=${fields}&$top=25`)
  }

  // ── Organization ──────────────────────────────────────────────────────────
  async getOrg () {
    return this.get('/organization?$select=displayName,onPremisesLastSyncDateTime')
  }

  // ── Group Members ─────────────────────────────────────────────────────────
  async getGroupMembers (groupId) {
    return this.get(`/groups/${groupId}/members?$select=displayName,userPrincipalName,@odata.type`)
  }

  // ── User App Assignments ──────────────────────────────────────────────────
  async getUserAppAssignments (userId) {
    return this.get(`/users/${userId}/appRoleAssignments?$select=resourceDisplayName,resourceId`)
  }

  // ── Synchronization / SCIM ────────────────────────────────────────────────
  async getSyncJobs (spId) {
    try {
      const v1 = await this.get(`/servicePrincipals/${spId}/synchronization/jobs`)
      if (v1?.value?.length) return v1.value
    } catch { /* fall through */ }
    try {
      const beta = await this.request('GET', `/servicePrincipals/${spId}/synchronization/jobs`, null, { beta: true })
      return beta?.value || []
    } catch { return [] }
  }

  async restartSyncJob (spId, jobId) {
    try {
      await this.post(`/servicePrincipals/${spId}/synchronization/jobs/${jobId}/restart`, {
        criteria: { resetScope: 'Full' }
      })
      return true
    } catch {
      return this.request('POST', `/servicePrincipals/${spId}/synchronization/jobs/${jobId}/restart`,
        { criteria: { resetScope: 'Full' } }, { beta: true })
    }
  }

  async createScimApp (appName) {
    let spId, appObjId
    try {
      const inst = await this.post(
        `/applicationTemplates/${NON_GALLERY_TEMPLATE_ID}/instantiate`,
        { displayName: appName }
      )
      spId = inst?.servicePrincipal?.id
      appObjId = inst?.application?.id
    } catch { /* fall through */ }

    if (!spId) {
      // Fallback: create custom app + SP
      const app = await this.post('/applications', {
        displayName: appName,
        signInAudience: 'AzureADMyOrg',
        tags: ['HideApp', 'WindowsAzureActiveDirectoryIntegratedApp']
      })
      appObjId = app?.id
      await new Promise(r => setTimeout(r, 2000))
      const sp = await this.post('/servicePrincipals', {
        appId: app.appId,
        tags: ['HideApp', 'WindowsAzureActiveDirectoryIntegratedApp']
      })
      spId = sp?.id
    }

    return { spId, appObjId }
  }

  async createSyncJob (spId) {
    // Try v1 with scim template, fall back to beta
    const attempts = [
      () => this.post(`/servicePrincipals/${spId}/synchronization/jobs`, { templateId: 'scim' }),
      () => this.request('POST', `/servicePrincipals/${spId}/synchronization/jobs`, { templateId: 'scim' }, { beta: true })
    ]
    for (const attempt of attempts) {
      try { return await attempt() } catch { /* try next */ }
      await new Promise(r => setTimeout(r, 3000))
    }
    return null
  }

  async setSyncSecrets (spId, secrets) {
    const body = { value: secrets }
    const uris = [
      `/servicePrincipals/${spId}/synchronization/secrets`,
    ]
    const methods = ['PATCH', 'PUT']
    for (let a = 0; a < 4; a++) {
      for (const uri of uris) {
        for (const method of methods) {
          try {
            await this.request(method, uri, body)
            return true
          } catch { /* try next */ }
        }
      }
      await new Promise(r => setTimeout(r, 3000))
    }
    return false
  }

  // ── Deployments ───────────────────────────────────────────────────────────
  async getDeployments () {
    return this.request('GET', '/admin/officeConfiguration/clientConfigurations', null, { beta: true })
  }
}

export const graphClient = new GraphClient()
export default graphClient
