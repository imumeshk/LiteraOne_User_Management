export type EnablementType = 'video' | 'guide' | 'doc' | 'template'

export interface EnablementResource {
  id: string | number
  title: string
  type: EnablementType | string
  description: string
  url: string
  icon?: string
  tags?: string[]
}

export interface EnablementDataPayload {
  resources: EnablementResource[]
}
