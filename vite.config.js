import { defineConfig } from 'vite'
import react from '@vitejs/plugin-react'

// https://vitejs.dev/config/
export default defineConfig({
  plugins: [react()],
  server: {
    allowedHosts: true // ปลดล็อกให้ LocalTunnel เข้าใช้งานได้ทุกโดเมน
  }
})