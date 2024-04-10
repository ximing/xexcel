// 查找/替换的纯扫描：匹配显示文本（经 evaluator）与公式原文 raw；每格至多一条匹配。
// 替换只作用于 raw（显示值命中但 raw 不含查询串的格由调用侧跳过）。
import { fromA1 } from '../core/addr'
import { SheetId, Workbook } from '../core/model'
import { CellEvaluator, isFormula } from './engine'
import { shiftFormula } from './transform'

export interface FindQuery {
  text: string
  caseSensitive: boolean
  wholeCell: boolean
  workbook: boolean // true=全簿；false=当前表
}

export interface FindMatch {
  sheet: SheetId
  row: number
  col: number
}

function matches(hay: string, q: FindQuery): boolean {
  const h = q.caseSensitive ? hay : hay.toLowerCase()
  const n = q.caseSensitive ? q.text : q.text.toLowerCase()
  return q.wholeCell ? h === n : h.includes(n)
}

export function findAll(doc: Workbook, ev: CellEvaluator, q: FindQuery): FindMatch[] {
  const out: FindMatch[] = []
  if (q.text === '') return out
  const sheets = q.workbook ? doc.order : [doc.active]
  for (const sid of sheets) {
    const data = doc.sheet(sid)
    const r = data.usedRange()
    for (let row = 0; row <= r.er; row++) {
      for (let col = 0; col <= r.ec; col++) {
        const cell = data.getCell(row, col)
        if (!cell || cell.raw === '') continue
        if (matches(ev.displayText(sid, row, col), q) || (isFormula(cell.raw) && matches(cell.raw, q))) {
          out.push({ sheet: sid, row, col })
        }
      }
    }
  }
  return out
}

function escapeRegExp(s: string): string {
  return s.replace(/[.*+?^${}()|[\]\\]/g, '\\$&')
}

// raw 级替换：命中返回新 raw，未命中返回 null
export function replaceInRaw(raw: string, q: FindQuery, replacement: string): string | null {
  if (q.text === '') return null
  if (q.wholeCell) return matches(raw, q) ? replacement : null
  // 公式 + 双方均为单元格引用（A1→B1）：按位移平移全部引用（=SUM(A1:A3) → =SUM(B1:B3)）
  if (isFormula(raw)) {
    const from = fromA1(q.text)
    const to = fromA1(replacement)
    if (from && to) {
      const shifted = shiftFormula(raw, to.row - from.row, to.col - from.col)
      return shifted === raw ? null : shifted
    }
  }
  if (q.caseSensitive) return raw.includes(q.text) ? raw.split(q.text).join(replacement) : null
  const re = new RegExp(escapeRegExp(q.text), 'gi')
  return re.test(raw) ? raw.replace(re, replacement) : null
}
