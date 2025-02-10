// theme.css @theme 与 view/theme.ts 的同值守护：防两处漂移
import { readFileSync } from 'node:fs'
import { describe, expect, it } from 'vitest'
import { THEME } from '../src/view/theme'

function cssColors(): Record<string, string> {
  const css = readFileSync('src/app/theme.css', 'utf8')
  const out: Record<string, string> = {}
  for (const m of css.matchAll(/--color-([a-z0-9-]+):\s*(#[0-9a-fA-F]{6});/g)) {
    const camel = m[1].replace(/-([a-z0-9])/g, (_, c: string) => c.toUpperCase())
    out[camel] = m[2].toLowerCase()
  }
  return out
}

describe('theme 镜像', () => {
  it('THEME 每个键值与 theme.css 同值', () => {
    const css = cssColors()
    for (const [k, v] of Object.entries(THEME)) {
      expect(css[k], `token ${k}`).toBe(v.toLowerCase())
    }
  })
  it('THEME 覆盖画布用到的全部色 token（11 个）', () => {
    expect(Object.keys(THEME).sort()).toEqual(
      ['hover', 'ink', 'ink2', 'ink3', 'line', 'lineStrong', 'primary', 'primarySoft', 'scrollbar', 'surface', 'surface2'].sort(),
    )
  })
})
