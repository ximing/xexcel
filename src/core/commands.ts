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
