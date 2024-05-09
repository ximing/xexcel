// 查找/替换的纯扫描：匹配显示文本（经 evaluator）与公式原文 raw；每格至多一条匹配。
// 替换只作用于 raw（显示值命中但 raw 不含查询串的格由调用侧跳过）。
import { SheetId, Workbook } from '../core/model'
import { CellEvaluator, isFormula } from './engine'

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

// 替换当前匹配后的游标前进：取旧序列中「被替换格的下一匹配」（循环），
// 返回它在新序列中的下标；找不到或新序列为空 → 0。
// 替换文本仍命中查询串（查 a 换 ab）时游标离开本格，不再反复替换同格。
export function indexAfterReplace(oldMatches: FindMatch[], oldIdx: number, newMatches: FindMatch[]): number {
  if (oldMatches.length === 0 || newMatches.length === 0) return 0
  const i = Math.min(Math.max(oldIdx, 0), oldMatches.length - 1)
  const target = oldMatches[(i + 1) % oldMatches.length]
  const j = newMatches.findIndex((m) => m.sheet === target.sheet && m.row === target.row && m.col === target.col)
  return j >= 0 ? j : 0
}

function escapeRegExp(s: string): string {
  return s.replace(/[.*+?^${}()|[\]\\]/g, '\\$&')
}

// raw 级替换：命中返回新 raw，未命中返回 null
export function replaceInRaw(raw: string, q: FindQuery, replacement: string): string | null {
  if (q.text === '') return null
  if (q.wholeCell) return matches(raw, q) ? replacement : null
  if (q.caseSensitive) return raw.includes(q.text) ? raw.split(q.text).join(replacement) : null
  const re = new RegExp(escapeRegExp(q.text), 'gi')
  return re.test(raw) ? raw.replace(re, replacement) : null
}
