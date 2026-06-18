// src/react/csvImport.ts
// CSV 导入组合：grid → 新建 sheet + 单事务（undo 一步复原）。纯逻辑，不 import view/DOM。
import { Cell, nextSheetId } from '@xexcel/core'
import { singleCell } from '@xexcel/core'
import { dedupeSheetName } from '@xexcel/core'
import type { SheetState } from '@xexcel/core'
import type { Transaction } from '@xexcel/core'
import { normalizedCell } from '@xexcel/core'

export const IMPORT_MIN_ROWS = 100
export const IMPORT_MIN_COLS = 26

// 工作表名去重（大小写不敏感，同 renameSheet 规则）；去重逻辑已提升 core/sheetname，此处保留导出兼容
export function uniqueSheetName(existing: Iterable<string>, base: string): string {
  return dedupeSheetName(existing, base.trim() || 'CSV')
}

export function buildImportTr(state: SheetState, grid: string[][], baseName: string): Transaction {
  const rowCount = Math.max(grid.length, IMPORT_MIN_ROWS)
  const colCount = Math.max(grid.reduce((m, r) => Math.max(m, r.length), 0), IMPORT_MIN_COLS)
  const id = nextSheetId(state.doc)
  const name = uniqueSheetName(state.doc.names.values(), baseName)

  const entries: { row: number; col: number; cell: Cell }[] = []
  grid.forEach((row, r) => {
    row.forEach((text, c) => {
      if (text === '') return // 空格跳过，保持稀疏
      entries.push({ row: r, col: c, cell: normalizedCell(text, undefined) })
    })
  })

  const tr = state.tr.insertSheet(id, name, { rowCount, colCount })
  if (entries.length) tr.setCells(id, entries)
  tr.setSelection(singleCell(0, 0))
  return tr
}
