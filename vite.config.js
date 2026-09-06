import { defineConfig, loadEnv } from 'vite'
import react from '@vitejs/plugin-react'

export default defineConfig(({ mode }) => {
  const env = loadEnv(mode, process.cwd(), '')
  const entraClientId = env.VITE_ENTRA_CLIENT_ID || env.ENTRA_CLIENT_ID || ''
  const entraClientSecret = env.VITE_ENTRA_CLIENT_SECRET || env.ENTRA_CLIENT_SECRET || ''
  const entraTenantId = env.VITE_ENTRA_TENANT_ID || env.ENTRA_TENANT_ID || 'organizations'

  return {
    plugins: [react()],
    server: { port: 3000 },
    define: {
      __ENTRA_CLIENT_ID__: JSON.stringify(entraClientId),
      __ENTRA_CLIENT_SECRET__: JSON.stringify(entraClientSecret),
      __ENTRA_TENANT_ID__: JSON.stringify(entraTenantId)
    }
  }
})
