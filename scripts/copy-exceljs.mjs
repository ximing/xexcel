// 把 exceljs UMD dist 拷到 public/vendor（浏览器经 <script> 直载，绕开 vite pre-bundle：
// esbuild CJS wrap 后的 exceljs 在浏览器 xlsx.load() 永不 resolve，见 M4b 验收）
import { copyFileSync, mkdirSync } from 'node:fs'
import { dirname, join } from 'node:path'
import { fileURLToPath } from 'node:url'

const root = join(dirname(fileURLToPath(import.meta.url)), '..')
const src = join(root, 'node_modules/exceljs/dist/exceljs.min.js')
const destDir = join(root, 'public/vendor')
mkdirSync(destDir, { recursive: true })
copyFileSync(src, join(destDir, 'exceljs.min.js'))
console.log('copied exceljs.min.js → public/vendor/')
