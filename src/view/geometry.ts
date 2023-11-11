// 网格几何：行高列宽前缀和 + 二分反查。不 import konva，可在 node 环境单测。
import { CellRange } from '../core/addr'
import { SheetData } from '../core/model'
import { Rect } from './types'

// 二分：最后一个 tops[i] <= v 的 i，clamp 到 [0, count-1]
function bsearch(tops: number[], v: number): number {
  const n = tops.length - 1 // 单元格个数
  if (v <= 0) return 0
  if (v >= tops[n]) return n - 1
  let lo = 0
  let hi = n // 不变式：tops[lo] <= v < tops[hi]
  while (hi - lo > 1) {
    const mid = (lo + hi) >> 1
    if (tops[mid] <= v) lo = mid
    else hi = mid
  }
  return lo
}

export class GridGeometry {
  private rowTops: number[] // 长度 rowCount+1，rowTops[i]=第 i 行顶 y
  private colLefts: number[] // 长度 colCount+1
  // 冻结窗格预留（M3 用，M1 恒 0）：渲染/hitTest 加偏移即可，此处零返工
  readonly frozenRows: number
  readonly frozenCols: number

  constructor(readonly sheet: SheetData, frozenRows = 0, frozenCols = 0) {
    this.frozenRows = frozenRows
    this.frozenCols = frozenCols
    this.rowTops = [0]
    for (let r = 0; r < sheet.rowCount; r++) this.rowTops.push(this.rowTops[r] + sheet.rowHeight(r))
    this.colLefts = [0]
    for (let c = 0; c < sheet.colCount; c++) this.colLefts.push(this.colLefts[c] + sheet.colWidth(c))
  }

  get contentWidth(): number {
    return this.colLefts[this.colLefts.length - 1]
  }

  get contentHeight(): number {
    return this.rowTops[this.rowTops.length - 1]
  }

  rowTop(row: number): number {
    return this.rowTops[row]
  }

  colLeft(col: number): number {
    return this.colLefts[col]
  }

  rowAt(y: number): number {
    return bsearch(this.rowTops, y)
  }

  colAt(x: number): number {
    return bsearch(this.colLefts, x)
  }

  cellRect(row: number, col: number): Rect {
    return { x: this.colLefts[col], y: this.rowTops[row], w: this.sheet.colWidth(col), h: this.sheet.rowHeight(row) }
  }

  rangeRect(r: CellRange): Rect {
    const a = this.cellRect(r.sr, r.sc)
    const b = this.cellRect(r.er, r.ec)
    return { x: a.x, y: a.y, w: b.x + b.w - a.x, h: b.y + b.h - a.y }
  }

  visibleRange(scrollX: number, scrollY: number, viewW: number, viewH: number): CellRange {
    return {
      sr: this.rowAt(scrollY),
      sc: this.colAt(scrollX),
      er: Math.max(this.rowAt(scrollY), this.rowAt(scrollY + viewH - 1)),
      ec: Math.max(this.colAt(scrollX), this.colAt(scrollX + viewW - 1)),
    }
  }
}
