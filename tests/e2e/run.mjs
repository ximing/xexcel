// e2e 回归 runner：node tests/e2e/run.mjs [suite...]（默认全部）
import { cleanup, ensureApp } from './lib/env.mjs'

const suites = {
  m4c: () => import('./suites/m4c.mjs'),
  m4a: () => import('./suites/m4a.mjs'),
  m4b: () => import('./suites/m4b.mjs'),
  m3c: () => import('./suites/m3c.mjs'),
}

// daemon evaluate 返回值经序列化后对象 key 序不稳定（实测按字母序），比较前须规范化
function canon(v) {
  if (Array.isArray(v)) return `[${v.map(canon).join(',')}]`
  if (v && typeof v === 'object') {
    return `{${Object.keys(v).sort().map((k) => `${JSON.stringify(k)}:${canon(v[k])}`).join(',')}}`
  }
  return JSON.stringify(v)
}

export function assertEq(actual, expected, label) {
  const a = canon(actual)
  const e = canon(expected)
  if (a !== e) throw new Error(`${label}\n  expected: ${e}\n  actual:   ${a}`)
}

const only = process.argv.slice(2)
const names = (only.length ? only : Object.keys(suites)).filter((n) => suites[n])
let pass = 0, fail = 0
await ensureApp()
for (const name of names) {
  const { default: run } = await suites[name]()
  try {
    await run({ assertEq })
    console.log(`✅ ${name}`)
    pass++
  } catch (e) {
    console.error(`❌ ${name}: ${e.message}`)
    fail++
  }
}
cleanup()
console.log(`suites: ${pass} 通过, ${fail} 失败`)
process.exit(fail ? 1 : 0)
