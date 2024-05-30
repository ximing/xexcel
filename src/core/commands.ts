import { rangesIntersect } from './addr'
import { CellStyle } from './model'
import { forEachSelectionRange, rangeSelection, selectionRange, singleCell } from './selection'
import type { SheetState } from './state'
import type { Transaction } from './transaction'

export type Command = (state: SheetState, dispatch?: (tr: Transaction) => void) => boolean

// 清空所有选区内已有内容的单元格；无内容 → false
export const clearSelection: Command = (state, dispatch) => {
  const data = state.activeSheet
  let hasCell = false
  forEachSelectionRange(state.selection, r => { data.forEachInRange(r, c => { if (c) hasCell = true }) })
  if (!hasCell) return false
  if (dispatch) {
    const tr = state.tr
    forEachSelectionRange(state.selection, r => tr.clearRange(r))
    dispatch(tr)
  }
  return true
}

// 全选：ranges=[全表]，activeCell={0,0}
export const selectAll: Command = (state, dispatch) => {
  const sel = rangeSelection(
    { sr: 0, sc: 0, er: state.activeSheet.rowCount - 1, ec: state.activeSheet.colCount - 1 },
    { row: 0, col: 0 },
  )
  if (dispatch) dispatch(state.tr.setSelection(sel))
  return true
}

export function setCellText(row: number, col: number, text: string): Command {
  return (state, dispatch) => {
    if (dispatch) dispatch(state.tr.setCell(row, col, text))
    return true
  }
}

// 对所有选区打样式补丁（undefined 删除对应样式键）
export function applyStylePatch(patch: Partial<CellStyle>): Command {
  return (state, dispatch) => {
    if (!dispatch) return true
    const tr = state.tr
    forEachSelectionRange(state.selection, r => tr.patchStyle(r, patch))
    dispatch(tr)
    return true
  }
}

// 合并所有选区：相交旧 merge 先移除；每区域保留左上锚点值，清其余格
export const mergeSelection: Command = (state, dispatch) => {
  const data = state.activeSheet
  const ranges = state.selection.ranges
  if (ranges.every(r => r.sr === r.er && r.sc === r.ec)) return false
  if (!dispatch) return true
  const kept = data.merges.filter(m => !ranges.some(r => rangesIntersect(m, r)))
  const newMerges = [...kept, ...ranges.filter(r => !(r.sr === r.er && r.sc === r.ec))]
  const tr = state.tr.setMerges(newMerges)
  const entries: { row: number; col: number; cell: null }[] = []
  for (const r of ranges) {
    if (r.sr === r.er && r.sc === r.ec) continue
    for (let row = r.sr; row <= r.er; row++)
      for (let col = r.sc; col <= r.ec; col++) {
        if (row === r.sr && col === r.sc) continue
        if (data.getCell(row, col)) entries.push({ row, col, cell: null })
      }
  }
  if (entries.length) tr.setCells(state.doc.active, entries)
  dispatch(tr)
  return true
}

// 拆分：移除与任一选区相交的所有 merge
export const unmergeSelection: Command = (state, dispatch) => {
  const ranges = state.selection.ranges
  const data = state.activeSheet
  const kept = data.merges.filter(m => !ranges.some(r => rangesIntersect(m, r)))
  if (kept.length === data.merges.length) return false
  if (dispatch) dispatch(state.tr.setMerges(kept))
  return true
}

// 保留 selectionRange 再导出供 CAT-B/C 未迁移站点引用（迁移完成后可移除）
export { selectionRange, singleCell }
