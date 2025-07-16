import { describe, it, expect } from 'vitest'
import { Workbook } from '@gmi/excel-core'
import { evaluatorFor } from '@gmi/excel-core'
import { computeFillEntries } from '../src/plugins/fillhandle'

const mk = (cells: Array<[number, number, string]>) => {
  let sheet = Workbook.create({ rowCount: 20, colCount: 10 }).activeSheet
  for (const [r, c, raw] of cells) sheet = sheet.setCell(r, c, { raw })
  return sheet
}

describe('computeFillEntries', () => {
  it('公式向下填充：相对引用逐行偏移', () => {
    const sheet = mk([[0, 0, '1'], [0, 1, '=A1*2']])
    const entries = computeFillEntries(sheet, () => '', { sr: 0, sc: 1, er: 0, ec: 1 }, { sr: 0, sc: 1, er: 2, ec: 1 })
    expect(entries).toEqual([
      { row: 1, col: 1, cell: { raw: '=A2*2' } },
      { row: 2, col: 1, cell: { raw: '=A3*2' } },
    ])
  })
  it('公式向右填充 + $ 锁定', () => {
    const sheet = mk([[0, 0, '=$A1*2']])
    const entries = computeFillEntries(sheet, () => '', { sr: 0, sc: 0, er: 0, ec: 0 }, { sr: 0, sc: 0, er: 0, ec: 2 })
    expect(entries.map((e) => e.cell?.raw)).toEqual(['=$A1*2', '=$A1*2'])
  })
  it('多格源区域平铺时按各自偏移量调整', () => {
    const sheet = mk([[0, 0, '=A1'], [1, 0, '=A2']])
    const entries = computeFillEntries(
      sheet,
      () => '',
      { sr: 0, sc: 0, er: 1, ec: 0 },
      { sr: 0, sc: 0, er: 3, ec: 0 },
    )
    expect(entries.map((e) => e.cell?.raw)).toEqual(['=A3', '=A4'])
  })
  it('数字序列仍为等差（不偏移）', () => {
    const wb = Workbook.create({ rowCount: 20, colCount: 10 })
    const sheet = wb.activeSheet.setCell(0, 0, { raw: '1' }).setCell(1, 0, { raw: '3' })
    const doc = wb.setSheet('s1', sheet)
    const ev = evaluatorFor(doc)
    const entries = computeFillEntries(sheet, (r, c) => ev.get('s1', r, c), { sr: 0, sc: 0, er: 1, ec: 0 }, { sr: 0, sc: 0, er: 3, ec: 0 })
    expect(entries.map((e) => e.cell?.raw)).toEqual(['5', '7'])
  })
  it('文本源原样平铺；源空格清目标格；样式随拷贝', () => {
    const sheet = mk([[0, 0, 'x']])
    const styled = sheet.setCell(0, 1, { raw: 'y', style: { bold: true } })
    const entries = computeFillEntries(styled, () => '', { sr: 0, sc: 0, er: 0, ec: 2 }, { sr: 0, sc: 0, er: 2, ec: 2 })
    const at = (r: number, c: number) => entries.find((e) => e.row === r && e.col === c)?.cell
    expect(at(1, 0)?.raw).toBe('x')
    expect(at(1, 1)).toEqual({ raw: 'y', style: { bold: true } })
    expect(at(1, 2)).toBeNull()
  })
})
