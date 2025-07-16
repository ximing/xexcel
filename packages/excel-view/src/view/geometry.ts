// 网格几何：行高列宽前缀和 + 二分反查。不 import konva，可在 node 环境单测。
import { CellAddr, CellRange } from '@gmi/excel-core'
import { SheetData } from '@gmi/excel-core'
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
  readonly frozenRows: number
  readonly frozenCols: number

  constructor(
    readonly sheet: SheetData,
    frozenRows = 0,
    frozenCols = 0,
    extraHiddenRows?: ReadonlySet<number>,
    autoRows?: ReadonlyMap<number, number>,
    readonly zoom = 1,
  ) {
    this.frozenRows = frozenRows
    this.frozenCols = frozenCols
    this.rowTops = [0]
    for (let r = 0; r < sheet.rowCount; r++) {
      // 行高优先级：隐藏 0 > 手动 > max(自动推导, 默认)
      // base 为 0 即模型层隐藏（rowHeight 对隐藏行返回 0），自动行高不得撑开
      const base = sheet.rowHeight(r)
      const h =
        extraHiddenRows?.has(r) || base === 0
          ? 0
          : sheet.customRowHeights.has(r)
            ? base
            : Math.max(base, autoRows?.get(r) ?? 0)
      this.rowTops.push(this.rowTops[r] + h * zoom)
    }
    this.colLefts = [0]
    for (let c = 0; c < sheet.colCount; c++) this.colLefts.push(this.colLefts[c] + sheet.colWidth(c) * zoom)
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

  rowHeight(row: number): number {
    return this.rowTops[row + 1] - this.rowTops[row]
  }

  colWidth(col: number): number {
    return this.colLefts[col + 1] - this.colLefts[col]
  }

  rowAt(y: number): number {
    return bsearch(this.rowTops, y)
  }

  colAt(x: number): number {
    return bsearch(this.colLefts, x)
  }

  // 宽高用前缀和差值：筛选隐藏（extraHiddenRows）塌缩的行/列返回 0，而非模型层的原始尺寸
  cellRect(row: number, col: number): Rect {
    return { x: this.colLefts[col], y: this.rowTops[row], w: this.colWidth(col), h: this.rowHeight(row) }
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

  // 防御性钳位：冻结数超出表尺寸（陈旧数据）时不越界读前缀和，避免 NaN 渲染
  get frozenWidth(): number {
    return this.colLefts[Math.min(this.frozenCols, this.colLefts.length - 1)]
  }

  get frozenHeight(): number {
    return this.rowTops[Math.min(this.frozenRows, this.rowTops.length - 1)]
  }

  // 内容区坐标（已减表头）→ 单元格；冻结区不吃 scroll。未 clamp（调用侧负责）
  cellAtContent(cx: number, cy: number, scrollX: number, scrollY: number): CellAddr {
    const col =
      cx < this.frozenWidth ? this.colAt(cx) : this.colAt(this.frozenWidth + scrollX + (cx - this.frozenWidth))
    const row =
      cy < this.frozenHeight ? this.rowAt(cy) : this.rowAt(this.frozenHeight + scrollY + (cy - this.frozenHeight))
    return { row, col }
  }
}
