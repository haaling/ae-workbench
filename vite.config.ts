import { defineConfig } from 'vite'
import react from '@vitejs/plugin-react'

// https://vite.dev/config/
export default defineConfig({
  plugins: [react()],
  server: {
    port: 5173,
    proxy: {
      '/api': {
        target: process.env.VITE_LOCAL_API_PROXY_TARGET || 'http://127.0.0.1:3100',
        changeOrigin: true
      }
    }
  }
})
