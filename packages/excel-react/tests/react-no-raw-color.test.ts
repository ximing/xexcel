// 规范回潮守卫：react 层禁止十六进制色值（token 唯一来源 theme.css）
import { readdirSync, readFileSync } from 'node:fs'
import { join } from 'node:path'
import { fileURLToPath } from 'node:url'
import { describe, expect, it } from 'vitest'

// 白名单：cfstyle.ts 的色值是写入文档的样式数据（CF 预设/color input 兜底），非 UI chrome
const WHITELIST = new Set(['cfstyle.ts'])

function tsxFiles(dir: string): string[] {
  return readdirSync(dir, { withFileTypes: true }).flatMap((e) =>
    e.isDirectory() ? tsxFiles(join(dir, e.name)) : e.name.endsWith('.tsx') || e.name.endsWith('.ts') ? [join(dir, e.name)] : [],
  )
}

describe('裸色守卫', () => {
  it('src 无十六进制色值', () => {
    const bad: string[] = []
    for (const f of tsxFiles(fileURLToPath(new URL('../src', import.meta.url)))) {
      if (WHITELIST.has(f.split('/').pop()!)) continue
      const hits = readFileSync(f, 'utf8').match(/#[0-9a-fA-F]{3,8}\b/g)
      if (hits) bad.push(`${f}: ${hits.join(', ')}`)
    }
    expect(bad).toEqual([])
  })
})
