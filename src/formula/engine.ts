// 公式引擎公共面：CellEvaluator（带 memo + 循环检测）与 evaluatorFor（按 Workbook 缓存）。
// 语义见 spec §4.8。
import { SheetId, Workbook } from '../core/model'
import {
  EvalCtx,
  FormulaError,
  FormulaValue,
  evalNode,
  formatNumber,
  isBlank,
  isError,
} from './eval'
import { parseFormula } from './parser'
import { formatValue } from './format'

export { isError }
export type { FormulaError, FormulaValue }

export function isFormula(raw: string): boolean {
  return raw.startsWith('=')
}

// 非公式单元格：trim 后纯数字（含科学计数法）→ number；"50%" → 0.5；其余保留原串
const NUMERIC_RE = /^-?\d+(\.\d+)?([eE][+-]?\d+)?$/
const PERCENT_RE = /^-?\d+(\.\d+)?%$/

function convertRaw(raw: string): FormulaValue {
  const t = raw.trim()
  if (NUMERIC_RE.test(t)) return Number(t)
  if (PERCENT_RE.test(t)) return Number(t.slice(0, -1)) / 100
  return raw
}

export class CellEvaluator {
  private memo = new Map<string, FormulaValue>()
  private visiting = new Set<string>()

  constructor(private workbook: Workbook) {}

  get(sheet: SheetId, row: number, col: number): FormulaValue {
    const data = this.workbook.sheets.get(sheet)
    // 引用越界（含未知表）→ #REF!
    if (!data || row < 0 || col < 0 || row >= data.rowCount || col >= data.colCount) {
      return { error: '#REF!' }
    }
    const key = `${sheet}!${row},${col}`
    const cached = this.memo.get(key)
    if (cached !== undefined) return cached
    // 正在求值链上再次遇到 → 循环引用；缓存 #CYCLE! 使链上各格一致
    if (this.visiting.has(key)) {
      const cyc: FormulaError = { error: '#CYCLE!' }
      this.memo.set(key, cyc)
      return cyc
    }
    const raw = data.getCell(row, col)?.raw ?? ''
    let value: FormulaValue
    if (isFormula(raw)) {
      this.visiting.add(key)
      try {
        value = this.evalFormula(raw, sheet, row, col)
      } finally {
        this.visiting.delete(key)
      }
    } else {
      value = convertRaw(raw)
    }
    this.memo.set(key, value)
    return value
  }

  displayText(sheet: SheetId, row: number, col: number): string {
    const v = this.get(sheet, row, col)
    if (isError(v)) return v.error
    if (typeof v === 'boolean') return v ? 'TRUE' : 'FALSE'
    if (typeof v === 'number') {
      const fmt = this.workbook.sheets.get(sheet)?.getCell(row, col)?.style?.numFmt
      if (fmt) {
        const s = formatValue(fmt, v)
        if (s !== null) return s
      }
      return formatNumber(v)
    }
    return v
  }

  private evalFormula(raw: string, sheet: SheetId, row: number, col: number): FormulaValue {
    const ctx: EvalCtx = {
      sheet,
      row,
      col,
      get: (s, r, c) => this.get(s, r, c),
      resolveSheet: (name) => {
        const lower = name.toLowerCase()
        for (const [id, n] of this.workbook.names) if (n.toLowerCase() === lower) return id
        return null
      },
    }
    try {
      const v = evalNode(parseFormula(raw.slice(1)), ctx)
      // BLANK 哨兵不越过公式边界：对外仍是 ''
      return isBlank(v) ? '' : v
    } catch {
      // 词法/语法错误（含未知裸名字）→ #NAME?
      return { error: '#NAME?' }
    }
  }
}

const cache = new WeakMap<Workbook, CellEvaluator>()

export function evaluatorFor(doc: Workbook): CellEvaluator {
  let ev = cache.get(doc)
  if (!ev) {
    ev = new CellEvaluator(doc)
    cache.set(doc, ev)
  }
  return ev
}
