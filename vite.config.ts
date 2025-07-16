import path from 'node:path'
import { defineConfig } from 'vitest/config'
import react from '@vitejs/plugin-react'
import tailwindcss from '@tailwindcss/vite'

export default defineConfig({
  base: './',
  plugins: [react(), tailwindcss()],
  resolve: {
    alias: {
      '@gmi/excel-core': path.resolve(__dirname, 'packages/excel-core/src'),
    },
  },
  test: { environment: 'node', include: ['tests/**/*.test.ts'], passWithNoTests: true },
})
