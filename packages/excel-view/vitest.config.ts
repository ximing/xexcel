import path from 'node:path'
import { defineConfig } from 'vitest/config'

export default defineConfig({
  resolve: {
    alias: { '@gmi/excel-core': path.resolve(__dirname, '../excel-core/src/index.ts') },
  },
  test: { environment: 'node', include: ['tests/**/*.test.ts'] },
})
