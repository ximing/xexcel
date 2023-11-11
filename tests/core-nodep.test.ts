import { describe, it, expect } from 'vitest'
import { readFileSync, readdirSync } from 'node:fs'
import { join } from 'node:path'

const scan = (dir: string): string[] =>
  readdirSync(dir, { withFileTypes: true }).flatMap(e =>
    e.isDirectory() ? scan(join(dir, e.name)) : e.name.endsWith('.ts') ? [join(dir, e.name)] : [])

describe('core/formula 零上层依赖', () => {
  for (const dir of ['src/core', 'src/formula']) {
    it(`${dir} 不 import DOM/view/react/konva`, () => {
      for (const f of scan(dir)) {
        const src = readFileSync(f, 'utf8')
        expect(src, f).not.toMatch(/from\s+['"].*(view|react|app)\//)
        expect(src, f).not.toMatch(/from\s+['"]konva/)
        expect(src, f).not.toMatch(/from\s+['"]react/)
      }
    })
  }
})
