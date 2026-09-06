import rawGraphClient from './graphClient.js'
import type { GraphApp, GraphAssignment, GraphPrincipal } from '../types/graph'

type AnyRecord = Record<string, any>

export interface TypedGraphClient {
  readonly isAuthenticated: boolean
  searchApps: (query?: string) => Promise<{ value?: GraphApp[] }>
  getAppPermissions: (appId: string) => Promise<{ value?: AnyRecord[] }>
  get: (path: string, opts?: AnyRecord) => Promise<any>
  getPermissionGrants: (spId: string) => Promise<{ value?: AnyRecord[] }>
  getAssignments: (spId: string) => Promise<GraphAssignment[]>
  enrichAssignments: (assignments: GraphAssignment[]) => Promise<GraphAssignment[]>
  getSpTags: (spId: string) => Promise<{ tags?: string[] }>
  setSpTags: (spId: string, tags: string[]) => Promise<boolean>
  setAppTagsByAppId: (appId: string, tags: string[]) => Promise<boolean>
  getSyncJobs: (spId: string) => Promise<any[]>
  restartSyncJob: (spId: string, jobId: string) => Promise<any>
  grantAdminConsent: (sp: GraphApp) => Promise<string[]>
  removeAssignment: (spId: string, assignmentId: string) => Promise<any>
  addAssignment: (spId: string, principalId: string, principalType: string) => Promise<any>
  deleteServicePrincipal: (spId: string) => Promise<any>
  searchPrincipals: (type: 'users' | 'groups', query: string) => Promise<{ value?: GraphPrincipal[] }>
  getGroupMembers: (groupId: string) => Promise<{ value?: AnyRecord[] }>
  getUserAppAssignments: (userId: string) => Promise<{ value?: AnyRecord[] }>
  createScimApp: (appName: string) => Promise<{ spId?: string, appObjId?: string }>
  createSyncJob: (spId: string) => Promise<{ id?: string } | null>
  setSyncSecrets: (spId: string, secrets: Array<{ key: string, value: string }>) => Promise<boolean>
  getDeployments: () => Promise<{ value?: AnyRecord[] }>
}

const graphClient = rawGraphClient as TypedGraphClient

export default graphClient
