// 条件格式命中判定（渲染时求值）：值规则复用 matchCriteria；duplicate 需 range 级预扫描。
import { CFStyle, CondFormatRule } from '../core/model'
import { matchCriteria } from './criteria'
import type { CellEvaluator } from './engine'

const VALUE_OP_PREFIX: Record<string, string> = {
  gt: '>', gte: '>=', lt: '<', lte: '<=', eq: '=', neq: '<>',
}

// 单格命中：按数组序（优先级）返回首个命中规则的样式
export function condFormatStyle(
  rules: readonly CondFormatRule[],
  sheetId: string,
  row: number,
  col: number,
  ev: CellEvaluator,
  dupSets: ReadonlyMap<string, ReadonlySet<string>>,
): CFStyle | undefined {
  for (const rule of rules) {
    const r = rule.range
    if (row < r.sr || row > r.er || col < r.sc || col > r.ec) continue
    switch (rule.type) {
      case 'value': {
        const v = ev.get(sheetId, row, col)
        if (rule.op === 'between') {
          const lo = Number(rule.v1)
          const hi = Number(rule.v2)
          if (typeof v === 'number' && !Number.isNaN(lo) && !Number.isNaN(hi) && v >= lo && v <= hi) {
            return rule.style
          }
          continue
        }
        const prefix = VALUE_OP_PREFIX[rule.op]
        if (!prefix) continue // 文本域 op 不适于值规则
        if (matchCriteria(prefix + rule.v1, v === undefined ? '' : (v as number | string | boolean))) {
          return rule.style
        }
        break
      }
      case 'textContains': {
        if (rule.text === '') break
        const t = ev.displayText(sheetId, row, col)
        if (t !== '' && t.toLowerCase().includes(rule.text.toLowerCase())) return rule.style
        break
      }
      case 'duplicate': {
        const set = dupSets.get(rule.id)
        if (!set) break
        const t = ev.displayText(sheetId, row, col)
        if (t !== '' && set.has(t.toLowerCase())) return rule.style
        break
      }
    }
  }
  return undefined
}

// duplicate 预扫描：rule.id → range 内出现 ≥2 次的显示文本集合（小写）
export function duplicateSets(
  rules: readonly CondFormatRule[],
  sheetId: string,
  ev: CellEvaluator,
): Map<string, Set<string>> {
  const out = new Map<string, Set<string>>()
  for (const rule of rules) {
    if (rule.type !== 'duplicate') continue
    const counts = new Map<string, number>()
    for (let r = rule.range.sr; r <= rule.range.er; r++) {
      for (let c = rule.range.sc; c <= rule.range.ec; c++) {
        const t = ev.displayText(sheetId, r, c).toLowerCase()
        if (t === '') continue
        counts.set(t, (counts.get(t) ?? 0) + 1)
      }
    }
    const s = new Set<string>()
    for (const [t, n] of counts) if (n >= 2) s.add(t)
    out.set(rule.id, s)
  }
  return out
}
