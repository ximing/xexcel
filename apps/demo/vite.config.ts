import path from 'node:path'
import { defineConfig } from 'vitest/config'
import react from '@vitejs/plugin-react'
import tailwindcss from '@tailwindcss/vite'

// 包名直指源码：demo dev/build 免预构建 SDK dist；外部消费才走 tsup 产物
export default defineConfig({
  base: './',
  plugins: [react(), tailwindcss()],
  resolve: {
    dedupe: ['react', 'react-dom'],
    alias: {
      '@xexcel/core': path.resolve(import.meta.dirname, '../../packages/excel-core/src'),
      '@xexcel/view': path.resolve(import.meta.dirname, '../../packages/excel-view/src'),
      '@xexcel/react': path.resolve(import.meta.dirname, '../../packages/excel-react/src'),
      // node 环境下 vitest 走 konva 的 main（index-node.js，硬 require 'canvas'），
      // 别名为浏览器构建，与 browser 字段一致，app 构建行为不变。
      konva: path.resolve(import.meta.dirname, '../../packages/excel-view/node_modules/konva/lib/index.js'),
    },
  },
  test: { environment: 'node', include: ['tests/**/*.test.ts'], passWithNoTests: true },
})
