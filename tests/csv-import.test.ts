// tests/csv-import.test.ts
import { describe, expect, it } from 'vitest'
import { history, undo, undoDepth } from '@gmi/excel-core'
import { Workbook } from '@gmi/excel-core'
import { SheetState } from '@gmi/excel-core'
import { buildImportTr, IMPORT_MIN_COLS, IMPORT_MIN_ROWS, uniqueSheetName } from '../src/react/csvImport'

const mk = () =>
  SheetState.create({ doc: Workbook.create({ rowCount: 10, colCount: 5 }), plugins: [history()] })

describe('uniqueSheetName', () => {
  it('无冲突直接用；空 base 用 CSV', () => {
    expect(uniqueSheetName(['Sheet1'], 'data')).toBe('data')
    expect(uniqueSheetName([], '  ')).toBe('CSV')
  })

  it('大小写不敏感去重，追加序号', () => {
    expect(uniqueSheetName(['Data', 'data (2)'], 'DATA')).toBe('DATA (3)')
  })
})

describe('buildImportTr', () => {
  it('新建 sheet 并激活；公式/数字/日期经 normalizedCell 解析；空格跳过', () => {
    let s = mk()
    const grid = [
      ['姓名', '数量'],
      ['苹果', '3'],
      ['=B2*2', '2026/1/5'],
    ]
    s = s.applyTransaction(buildImportTr(s, grid, '导入表')).state
    expect(s.doc.order).toEqual(['s1', 's2'])
    expect(s.doc.active).toBe('s2')
    expect(s.doc.names.get('s2')).toBe('导入表')

    const sh = s.doc.sheet('s2')
    expect(sh.rowCount).toBe(IMPORT_MIN_ROWS)
    expect(sh.colCount).toBe(IMPORT_MIN_COLS)
    expect(sh.getCell(1, 0)).toEqual({ raw: '苹果' })
    expect(sh.getCell(2, 0)).toEqual({ raw: '=B2*2' })
    // 日期 → serial + numFmt（serial 具体值由 dateSerial 决定，只断言 numFmt 与 raw 非原文）
    const d = sh.getCell(2, 1)!
    expect(d.style?.numFmt).toBe('yyyy/m/d')
    expect(d.raw).not.toBe('2026/1/5')
    expect(s.selection.activeCell).toEqual({ row: 0, col: 0 })
  })

  it('undo 一步复原（insertSheet+setCells 同组）', () => {
    let s = mk()
    s = s.applyTransaction(buildImportTr(s, [['a']], 'x')).state
    expect(undoDepth(s)).toBe(1)
    expect(undo(s, (tr) => { s = s.applyTransaction(tr).state })).toBe(true)
    expect(s.doc.order).toEqual(['s1'])
    expect(s.doc.active).toBe('s1')
  })

  it('行/列数超出下限值时按内容扩容', () => {
    let s = mk()
    const grid = Array.from({ length: 150 }, () => ['x'])
    s = s.applyTransaction(buildImportTr(s, grid, 'big')).state
    expect(s.doc.sheet('s2').rowCount).toBe(150)
  })
})
