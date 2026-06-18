import path from 'node:path'
import { defineConfig } from 'vitest/config'

export default defineConfig({
  resolve: {
    alias: {
      '@xexcel/core': path.resolve(__dirname, '../excel-core/src/index.ts'),
      '@xexcel/view': path.resolve(__dirname, '../excel-view/src/index.ts'),
      // node 环境下 barrel import 会命中 konva 的 main（index-node.js 硬 require('canvas')），强制指到浏览器版
      konva: path.resolve(__dirname, '../excel-view/node_modules/konva/lib/index.js'),
    },
  },
  test: { environment: 'node', include: ['tests/**/*.test.ts'] },
})
