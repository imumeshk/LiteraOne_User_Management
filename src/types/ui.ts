export type Page = 'dashboard' | 'apps' | 'deploy' | 'enablement' | 'users-management'

export type ToastFn = (message: string, type?: string) => void
export type LogFn = (message: string) => void
