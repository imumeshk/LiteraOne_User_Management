export interface GraphApp {
  id: string
  appId: string
  displayName: string
  tags?: string[]
  accountEnabled?: boolean
}

export interface GraphAssignment {
  id: string
  principalId: string
  principalType?: string
  principalDisplayName?: string
  emailOrId?: string
  _checked?: boolean
}

export interface GraphPrincipal {
  id: string
  displayName: string
  mail?: string
  userPrincipalName?: string
}
