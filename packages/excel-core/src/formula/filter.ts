// 筛选隐藏推导：对筛选区域数据区逐行评估各列 criteria（AND），不匹配的行入隐藏集。
// 实时推导（几何构建时重算），手动隐藏与本集取并集生效；表头行（range.sr）不参与。
import { FilterConditionCriteria, FilterCriteria, SheetData, SheetId } from '../core/model'
import { CellEvaluator } from './engine'
import { FormulaValue } from './eval'

function matchCondition(c: FilterConditionCriteria, value: FormulaValue, text: string): boolean {
  if (c.field === 'text') {
    const t = text.toLowerCase()
    const v = c.v1.toLowerCase()
    switch (c.op) {
      case 'contains': return t.includes(v)
      case 'notContains': return !t.includes(v)
      case 'startsWith': return t.startsWith(v)
      case 'endsWith': return t.endsWith(v)
      case 'eq': return t === v
      case 'neq': return t !== v
      default: return false
    }
  }
  // 数值条件：求值结果必须是数字，v1/v2 必须可解析
  if (typeof value !== 'number') return false
  const n1 = Number(c.v1)
  if (Number.isNaN(n1)) return false
  switch (c.op) {
    case 'eq': return value === n1
    case 'neq': return value !== n1
    case 'gt': return value > n1
    case 'gte': return value >= n1
    case 'lt': return value < n1
    case 'lte': return value <= n1
    case 'between': {
      const n2 = Number(c.v2)
      return !Number.isNaN(n2) && value >= n1 && value <= n2
    }
    default: return false
  }
}

function matchCriteriaOf(crit: FilterCriteria, value: FormulaValue, text: string): boolean {
  if (crit.type === 'values') return !crit.excluded.includes(text)
  return matchCondition(crit, value, text)
}

export function filterHiddenRows(sheetId: SheetId, data: SheetData, ev: CellEvaluator): Set<number> {
  const out = new Set<number>()
  const f = data.filter
  if (!f) return out
  const active = Object.entries(f.criteria).filter(
    ([, c]) => c.type !== 'values' || c.excluded.length > 0,
  )
  if (active.length === 0) return out
  for (let row = f.range.sr + 1; row <= f.range.er; row++) {
    let visible = true
    for (const [colStr, crit] of active) {
      const col = Number(colStr)
      const cell = data.getCell(row, col)
      const value: FormulaValue = cell ? ev.get(sheetId, row, col) : ''
      const text = cell ? ev.displayText(sheetId, row, col) : ''
      if (!matchCriteriaOf(crit, value, text)) {
        visible = false
        break
      }
    }
    if (!visible) out.add(row)
  }
  return out
}
