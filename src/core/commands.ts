import { rangesIntersect } from './addr'
import { CellStyle } from './model'
import { Selection, selectionRange } from './selection'
import type { SheetState } from './state'
import type { Transaction } from './transaction'

export type Command = (state: SheetState, dispatch?: (tr: Transaction) => void) => boolean

// 清空当前选区内所有已有内容的单元格；选区内无内容 → false
export const clearSelection: Command = (state, dispatch) => {
  const r = selectionRange(state.selection)
  const data = state.activeSheet
  let hasCell = false
  data.forEachInRange(r, cell => {
    if (cell) hasCell = true
  })
  if (!hasCell) return false
  if (dispatch) dispatch(state.tr.clearRange(r))
  return true
}

// 全选：anchor 左上、focus 右下
export const selectAll: Command = (state, dispatch) => {
  const sel: Selection = {
    anchor: { row: 0, col: 0 },
    focus: { row: state.activeSheet.rowCount - 1, col: state.activeSheet.colCount - 1 },
  }
  if (dispatch) dispatch(state.tr.setSelection(sel))
  return true
}

export function setCellText(row: number, col: number, text: string): Command {
  return (state, dispatch) => {
    if (dispatch) dispatch(state.tr.setCell(row, col, text))
    return true
  }
}

// 对当前选区打样式补丁（undefined 值删除对应样式键）
export function applyStylePatch(patch: Partial<CellStyle>): Command {
  return (state, dispatch) => {
    if (dispatch) dispatch(state.tr.patchStyle(selectionRange(state.selection), patch))
    return true
  }
}

// 合并当前选区：相交旧 merge 先移除；保留左上锚点值，清其余格
export const mergeSelection: Command = (state, dispatch) => {
  const r = selectionRange(state.selection)
  if (r.sr === r.er && r.sc === r.ec) return false // 单格无可合并
  const data = state.activeSheet
  const kept = data.merges.filter((m) => !rangesIntersect(m, r))
  const tr = state.tr.setMerges([...kept, r])
  const entries: { row: number; col: number; cell: null }[] = []
  for (let row = r.sr; row <= r.er; row++) {
    for (let col = r.sc; col <= r.ec; col++) {
      if (row === r.sr && col === r.sc) continue
      if (data.getCell(row, col)) entries.push({ row, col, cell: null })
    }
  }
  if (entries.length) tr.setCells(state.doc.active, entries)
  if (dispatch) dispatch(tr)
  return true
}

// 拆分：移除与选区相交的所有 merge（不恢复合并时清除的值）
export const unmergeSelection: Command = (state, dispatch) => {
  const r = selectionRange(state.selection)
  const data = state.activeSheet
  const kept = data.merges.filter((m) => !rangesIntersect(m, r))
  if (kept.length === data.merges.length) return false
  if (dispatch) dispatch(state.tr.setMerges(kept))
  return true
}
