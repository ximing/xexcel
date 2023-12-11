import { describe, it, expect } from 'vitest'
import { Workbook } from '../src/core/model'
import { evaluatorFor } from '../src/formula/engine'
import { adjustDecimals, formatValue } from '../src/formula/format'

const mk = (raw: string, numFmt?: string) => {
  let sheet = Workbook.create({ rowCount: 10, colCount: 5 }).activeSheet
  sheet = sheet.setCell(0, 0, numFmt ? { raw, style: { numFmt } } : { raw })
  return Workbook.create({ rowCount: 10, colCount: 5 }).setSheet('s1', sheet)
}
const txt = (wb: Workbook) => evaluatorFor(wb).displayText('s1', 0, 0)

describe('formatValue / adjustDecimals', () => {
  it('formatValue 常见格式', () => {
    expect(formatValue('#,##0.00', 1234.5)).toBe('1,234.50')
    expect(formatValue('0%', 0.5)).toBe('50%')
    expect(formatValue('¥#,##0.00', 1234.5)).toBe('¥1,234.50')
    expect(formatValue('0.00E+00', 12345)).toBe('1.23E+04')
    expect(formatValue('yyyy/m/d', 45292)).toBe('2024/1/1') // serial 45292 = 2024-01-01
  })
  it('formatValue 非法格式串 → null', () => {
    expect(formatValue('"unclosed', 1)).toBe(null)
  })
  it('adjustDecimals', () => {
    expect(adjustDecimals(undefined, 1)).toBe('#,##0.000')
    expect(adjustDecimals('#,##0.00', 1)).toBe('#,##0.000')
    expect(adjustDecimals('#,##0.00', -1)).toBe('#,##0.0')
    expect(adjustDecimals('#,##0.0', -1)).toBe('#,##0')
    expect(adjustDecimals('#,##0', -1)).toBe('#,##0')
    expect(adjustDecimals('0%', 1)).toBe('0.0%')
  })
})

describe('displayText format-aware', () => {
  it('numFmt 格式化数字显示，raw 不变', () => {
    const wb = mk('0.5', '0%')
    expect(txt(wb)).toBe('50%')
    expect(wb.sheet('s1').getCell(0, 0)!.raw).toBe('0.5')
  })
  it('无 numFmt → 原行为', () => {
    expect(txt(mk('0.1'))).toBe('0.1')
    expect(txt(mk('=1/3'))).toBe('0.3333333333')
  })
  it('numFmt 不影响非数字值', () => {
    expect(txt(mk('abc', '#,##0.00'))).toBe('abc')
    expect(txt(mk('=1/0', '#,##0.00'))).toBe('#DIV/0!')
  })
  it('非法 numFmt 回退 formatNumber', () => {
    expect(txt(mk('1234.5', '"unclosed'))).toBe('1234.5')
  })
})
