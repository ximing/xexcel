// 排序计算：按求值结果对 range 数据区行重排，产出整 band 的 SetCells 条目。
// 移动语义：公式引用不改写；行高不动；undo 由 SetCellsStep.invert 快照保证。
// 比较序对齐 Excel：数字 < 文本（不区分大小写） < FALSE < TRUE < 错误；空格恒排末尾。
import { CellRange, rangesIntersect } from '../core/addr'
import { Cell, SheetData, SheetId } from '../core/model'
import { CellEvaluator } from './engine'
import { FormulaValue, isError } from './eval'

export interface SortKey {
  col: number // 绝对列号（须在 range 内）
  asc: boolean
}

function rank(v: FormulaValue): number {
  if (isError(v)) return 3
  if (typeof v === 'number') return 0
  if (typeof v === 'string') return 1
  return 2 // boolean
}

function compareVal(a: FormulaValue, b: FormulaValue): number {
  const ra = rank(a)
  const rb = rank(b)
  if (ra !== rb) return ra - rb
  if (typeof a === 'number' && typeof b === 'number') return a - b
  if (typeof a === 'string' && typeof b === 'string') {
    return a.toLowerCase().localeCompare(b.toLowerCase())
  }
  if (isError(a) && isError(b)) return a.error.localeCompare(b.error)
  return a === b ? 0 : a ? 1 : -1 // boolean
}

export function computeSortEntries(
  data: SheetData,
  sheetId: SheetId,
  ev: CellEvaluator,
  range: CellRange,
  keys: SortKey[],
  hasHeader: boolean,
): { row: number; col: number; cell: Cell | null }[] {
  const firstData = hasHeader ? range.sr + 1 : range.sr
  // blank 判定：无 cell 或 raw===''（不污染 evaluator memo）；blank 恒排末尾。
  // 公式格求值为 '' 不算 blank，按文本 '' 参与比较
  const valueOf = (r: number, c: number): { v: FormulaValue; blank: boolean } => {
    const cell = data.getCell(r, c)
    if (!cell || cell.raw === '') return { v: '', blank: true }
    return { v: ev.get(sheetId, r, c), blank: false }
  }
  const decorated: { r: number; i: number }[] = []
  for (let r = firstData; r <= range.er; r++) decorated.push({ r, i: r - firstData })
  decorated.sort((x, y) => {
    for (const k of keys) {
      const a = valueOf(x.r, k.col)
      const b = valueOf(y.r, k.col)
      if (a.blank || b.blank) {
        if (a.blank && b.blank) continue
        return a.blank ? 1 : -1 // 空格恒尾，不受 asc 影响
      }
      const c = compareVal(a.v, b.v)
      if (c !== 0) return k.asc ? c : -c
    }
    return x.i - y.i // 稳定
  })
  const entries: { row: number; col: number; cell: Cell | null }[] = []
  decorated.forEach(({ r: srcRow }, i) => {
    const destRow = firstData + i
    for (let col = range.sc; col <= range.ec; col++) {
      entries.push({ row: destRow, col, cell: data.getCell(srcRow, col) ?? null })
    }
  })
  return entries
}

// 排序区域与 merge 相交 → 拒绝（Excel 同）
export function sortBlockedByMerges(data: SheetData, range: CellRange): boolean {
  return data.merges.some((m) => rangesIntersect(m, range))
}
