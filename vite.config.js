import { defineConfig } from 'vite'
import react from '@vitejs/plugin-react'
import { resolve } from 'path'

// https://vitejs.dev/config/
export default defineConfig({
  plugins: [react()],
  build: {
    rollupOptions: {
      input: {
        // The main entry point (your existing system)
        main: resolve(__dirname, 'index.html'),
        // The new beta entry point
        beta: resolve(__dirname, 'beta/index.html'),
        test: resolve(__dirname, 'test/index.html'),
        alpha: resolve(__dirname, 'alpha/index.html'),
      },
    },
  },
})