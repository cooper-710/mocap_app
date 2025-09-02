import { defineConfig } from 'vite'
import react from '@vitejs/plugin-react'

export default defineConfig({
  plugins: [react()],
  base: '/mocap_app/',     // <-- MUST match the GH Pages subpath
})
