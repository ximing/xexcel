import { describe, it, expect } from 'vitest'
import { readFileSync, readdirSync } from 'node:fs'
import { fileURLToPath } from 'node:url'
import { join } from 'node:path'

const srcRoot = fileURLToPath(new URL('../src', import.meta.url))

const scan = (dir: string): string[] =>
  readdirSync(dir, { withFileTypes: true }).flatMap(e =>
    e.isDirectory() ? scan(join(dir, e.name)) : e.name.endsWith('.ts') ? [join(dir, e.name)] : [])

// 禁用 import 目标：任何上层包与 DOM/视图库
const FORBIDDEN = ['@xexcel/view', '@xexcel/react', 'konva', 'react', 'react-dom']

describe('core/formula 零上层依赖', () => {
  for (const dir of ['core', 'formula']) {
    it(`${dir} 不 import 上层包/DOM/view/react/konva`, () => {
      for (const f of scan(join(srcRoot, dir))) {
        const src = readFileSync(f, 'utf8')
        expect(src, f).not.toMatch(/from\s+['"].*(view|react|app)\//)
        for (const target of FORBIDDEN) {
          expect(src, f).not.toMatch(new RegExp(`from\\s+['"]${target}([/'"])`))
        }
      }
    })
  }
})
