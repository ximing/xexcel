import { describe, it, expect } from 'vitest'
import { SheetData, Workbook, DEFAULT_ROW_HEIGHT, DEFAULT_COL_WIDTH } from '../src/core/model'

describe('SheetData', () => {
  it('不可变 setCell：原对象不变，未受影响行结构共享', () => {
    const d0 = SheetData.create({ rowCount: 100, colCount: 26 })
    const d1 = d0.setCell(0, 0, { raw: 'a' })
    const d2 = d1.setCell(5, 3, { raw: 'b' })
    expect(d0.getCell(0, 0)).toBeUndefined()
    expect(d1.getCell(0, 0)).toEqual({ raw: 'a' })
    expect(d2.getCell(0, 0)).toEqual({ raw: 'a' })
    expect(d2.getCell(5, 3)).toEqual({ raw: 'b' })
    expect(d1.getCell(5, 3)).toBeUndefined()
  })
  it('raw==="" 且无 style 视为删除', () => {
    const d = SheetData.create({ rowCount: 10, colCount: 5 }).setCell(1, 1, { raw: 'x' })
    expect(d.setCell(1, 1, { raw: '' }).getCell(1, 1)).toBeUndefined()
    expect(d.setCell(1, 1, { raw: '', style: { bold: true } }).getCell(1, 1)).toEqual({ raw: '', style: { bold: true } })
    expect(d.setCell(1, 1, null).getCell(1, 1)).toBeUndefined()
  })
  it('行高列宽：默认值与自定义', () => {
    const d = SheetData.create({ rowCount: 10, colCount: 5 })
    expect(d.rowHeight(3)).toBe(DEFAULT_ROW_HEIGHT)
    expect(d.colWidth(2)).toBe(DEFAULT_COL_WIDTH)
    const d2 = d.setRowHeight(3, 40).setColWidth(2, 120)
    expect(d2.rowHeight(3)).toBe(40)
    expect(d2.colWidth(2)).toBe(120)
    expect(d2.customRowHeights.get(3)).toBe(40)
    expect(d2.setRowHeight(3, null).rowHeight(3)).toBe(DEFAULT_ROW_HEIGHT)
    expect(d2.setRowHeight(3, null).customRowHeights.has(3)).toBe(false)
  })
  it('usedRange / forEachInRange / toJSON 往返', () => {
    const d = SheetData.create({ rowCount: 100, colCount: 26 })
      .setCell(2, 1, { raw: 'a' }).setCell(7, 5, { raw: 'b' })
    expect(d.usedRange()).toEqual({ sr: 2, sc: 1, er: 7, ec: 5 })
    expect(SheetData.create({ rowCount: 10, colCount: 5 }).usedRange()).toEqual({ sr: 0, sc: 0, er: 0, ec: 0 })
    const seen: string[] = []
    d.forEachInRange({ sr: 2, sc: 1, er: 3, ec: 2 }, (c, r, col) => seen.push(`${r},${col}:${c?.raw ?? '-'}`))
    expect(seen).toEqual(['2,1:a', '2,2:-', '3,1:-', '3,2:-'])
    const back = SheetData.fromJSON(d.toJSON())
    expect(back.getCell(7, 5)).toEqual({ raw: 'b' })
    expect(back.rowCount).toBe(100)
  })
})

describe('Workbook', () => {
  it('create / sheet / setSheet / setActive / addSheet / removeSheet', () => {
    const wb = Workbook.create({ rowCount: 10, colCount: 5 })
    expect(wb.order).toEqual(['s1'])
    expect(wb.active).toBe('s1')
    expect(wb.names.get('s1')).toBe('Sheet1')
    const s2 = SheetData.create({ rowCount: 10, colCount: 5 })
    const wb2 = wb.addSheet('s2', s2)
    expect(wb2.order).toEqual(['s1', 's2'])
    expect(wb2.active).toBe('s1')
    expect(wb2.setActive('s2').active).toBe('s2')
    const wb3 = wb2.setSheet('s1', s2.setCell(0, 0, { raw: 'z' }))
    expect(wb3.sheet('s1').getCell(0, 0)).toEqual({ raw: 'z' })
    expect(wb2.sheet('s1').getCell(0, 0)).toBeUndefined() // 不可变
    const wb4 = wb2.setActive('s2').removeSheet('s2')
    expect(wb4.order).toEqual(['s1'])
    expect(wb4.active).toBe('s1')
    expect(() => wb4.removeSheet('s1')).toThrow() // 不允许删最后一个
    expect(wb2.renameSheet('s1', '数据').names.get('s1')).toBe('数据')
  })
  it('toJSON 往返', () => {
    const wb = Workbook.create({ rowCount: 10, colCount: 5 }).setSheet('s1', SheetData.create({ rowCount: 10, colCount: 5 }).setCell(1, 1, { raw: 'q' }))
    const back = Workbook.fromJSON(wb.toJSON())
    expect(back.sheet('s1').getCell(1, 1)).toEqual({ raw: 'q' })
    expect(back.active).toBe('s1')
  })
  it('新样式字段 JSON 往返', () => {
    const sheet = Workbook.create({ rowCount: 5, colCount: 5 }).activeSheet.setCell(1, 1, {
      raw: 'x',
      style: { numFmt: '0%', fontFamily: 'monospace', fontSize: 18, underline: true, strikethrough: true },
    })
    const back = SheetData.fromJSON(JSON.parse(JSON.stringify(sheet.toJSON())))
    expect(back.getCell(1, 1)).toEqual(sheet.getCell(1, 1))
  })
})
