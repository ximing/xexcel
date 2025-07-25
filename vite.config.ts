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
      '@gmi/excel-view': path.resolve(__dirname, 'packages/excel-view/src'),
      '@gmi/excel-react': path.resolve(__dirname, 'packages/excel-react/src'),
      // node 环境下 vitest 走 konva 的 main（index-node.js，硬 require 'canvas'），
      // 别名为浏览器构建，与 browser 字段一致，app 构建行为不变。
      konva: path.resolve(__dirname, 'packages/excel-view/node_modules/konva/lib/index.js'),
    },
  },
  test: { environment: 'node', include: ['tests/**/*.test.ts'], passWithNoTests: true },
})
