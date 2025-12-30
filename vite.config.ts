import { defineConfig, loadEnv } from 'vite'
import react from '@vitejs/plugin-react'

export default defineConfig(({ mode }) => {
  const env = loadEnv(mode, process.cwd(), '')
  const DEV_PORT = Number(env.VITE_DEV_PORT || 5173)
  const PROXY_TARGET = env.VITE_PROXY_TARGET || 'http://127.0.0.1:5000'
  return {
    plugins: [react()],
    server: {
      host: true, // listen on 0.0.0.0 for LAN access
      port: DEV_PORT,
      proxy: {
        '/api': {
          target: PROXY_TARGET,
          changeOrigin: true
        }
      }
    }
  }
})
