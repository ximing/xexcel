import { CellAddr, CellRange, normalizeRange } from './addr'

export interface Selection { anchor: CellAddr; focus: CellAddr } // focus=活动单元格

export function selectionRange(sel: Selection): CellRange {
  return normalizeRange({
    sr: sel.anchor.row,
    sc: sel.anchor.col,
    er: sel.focus.row,
    ec: sel.focus.col,
  })
}

export function singleCell(row: number, col: number): Selection {
  return { anchor: { row, col }, focus: { row, col } }
}
