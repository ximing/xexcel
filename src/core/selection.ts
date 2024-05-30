import { CellAddr, CellRange, normalizeRange } from './addr'

// 选区：多区域 + 活动格。ranges 长度恒 ≥1；activeCell 必落在 ranges[last]（活动区域）内。
// 不可变：所有变更经 Transaction.setSelection 构造全新对象，禁止就地改 ranges/activeCell
// （HistoryGroup.selection 按引用快照，就地改会静默破坏 undo）。
export interface Selection {
  ranges: CellRange[]
  activeCell: CellAddr
}

// 活动区域 = 含 activeCell 的 range，约定 = ranges[last]（最后加入者）；单区域即 ranges[0]。
export function selectionRange(sel: Selection): CellRange {
  return sel.ranges[sel.ranges.length - 1]
}

export function activeRange(sel: Selection): CellRange {
  return sel.ranges[sel.ranges.length - 1]
}

// 全部区域遍历（CAT-A：格式/清除/合并/复制源/渲染高亮 作用于全部区域）
export function forEachSelectionRange(sel: Selection, cb: (r: CellRange) => void): void {
  for (const r of sel.ranges) cb(r)
}

export function singleCell(row: number, col: number): Selection {
  return { ranges: [{ sr: row, sc: col, er: row, ec: col }], activeCell: { row, col } }
}

// 一个 range 作单区域选区；activeCell 缺省取其归一左上
export function rangeSelection(r: CellRange, activeCell?: CellAddr): Selection {
  const n = normalizeRange(r)
  return { ranges: [n], activeCell: activeCell ?? { row: n.sr, col: n.sc } }
}

// Ctrl+追加区域（新造 ranges 数组，末项为新增区域）；activeCell 落在新增区域
export function appendRange(sel: Selection, r: CellRange, activeCell?: CellAddr): Selection {
  const n = normalizeRange(r)
  return { ranges: [...sel.ranges, n], activeCell: activeCell ?? { row: n.sr, col: n.sc } }
}

// Ctrl+click 反选：移除最后加入的、含该格的 range（LIFO）；移除最后一个 → 回到该格单选
export function toggleRange(sel: Selection, row: number, col: number): Selection {
  for (let i = sel.ranges.length - 1; i >= 0; i--) {
    const r = sel.ranges[i]
    if (row >= r.sr && row <= r.er && col >= r.sc && col <= r.ec) {
      const ranges = sel.ranges.slice()
      ranges.splice(i, 1)
      if (ranges.length === 0) return singleCell(row, col)
      const last = ranges[ranges.length - 1]
      const ac = (row >= last.sr && row <= last.er && col >= last.sc && col <= last.ec)
        ? sel.activeCell
        : { row: last.sr, col: last.sc }
      return { ranges, activeCell: ac }
    }
  }
  return sel
}

// Shift+click/Arrow/drag 扩展活动区：由活动区域中 activeCell 的对角格（固定锚点）→ focus 框出 bbox，
// 替换 ranges[last] 边界（新造数组）；activeCell=focus（移动端，F2 编辑光标原 sel.focus 语义）。
// 锚点 = ranges[last] 中 activeCell 的对角格：activeCell 是移动端（扩展路径下恒为角点），
// 对角格即生长起始格（固定）→ 连续扩展生长不滑动；activeCell 非角点（结构级联等）按其落在
// 区域四象限取对角角兜底，保证不变式 activeCell∈ranges[last] 成立。
export function extendActiveRange(sel: Selection, focus: CellAddr): Selection {
  const last = sel.ranges[sel.ranges.length - 1]
  const cR = (last.sr + last.er) / 2
  const cC = (last.sc + last.ec) / 2
  const anchor: CellAddr = {
    row: sel.activeCell.row <= cR ? last.er : last.sr,
    col: sel.activeCell.col <= cC ? last.ec : last.sc,
  }
  const n = normalizeRange({ sr: anchor.row, sc: anchor.col, er: focus.row, ec: focus.col })
  const ranges = sel.ranges.slice(0, -1)
  ranges.push(n)
  return { ranges, activeCell: { row: focus.row, col: focus.col } }
}
