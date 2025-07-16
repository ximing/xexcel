import { describe, it, expect } from 'vitest'
import { Workbook } from '../src/core/model'
import { evaluatorFor } from '../src/formula/engine'
import { parseFormula } from '../src/formula/parser'
import { shiftFormula } from '../src/formula/transform'

describe('shiftFormula', () => {
  it('相对引用整体偏移', () => {
    expect(shiftFormula('=A1*2', 1, 0)).toBe('=A2*2')
    expect(shiftFormula('=A1+B2', 2, 3)).toBe('=D3+E4')
    expect(shiftFormula('=SUM(A1:A3)', 1, 1)).toBe('=SUM(B2:B4)')
  })
  it('$ 维度锁定', () => {
    expect(shiftFormula('=$A$1*2', 1, 1)).toBe('=$A$1*2')
    expect(shiftFormula('=$A1*2', 0, 5)).toBe('=$A1*2')
    expect(shiftFormula('=$A1*2', 3, 0)).toBe('=$A4*2')
    expect(shiftFormula('=A$1*2', 5, 2)).toBe('=C$1*2')
  })
  it('向上/向左越界 → #REF!', () => {
    expect(shiftFormula('=A2*2', -5, 0)).toBe('=#REF!*2')
    expect(shiftFormula('=SUM(B2:C3)', 0, -2)).toBe('=SUM(#REF!)')
  })
  it('跨表引用同样偏移本表维度', () => {
    expect(shiftFormula('=Sheet2!A1+1', 1, 1)).toBe('=Sheet2!B2+1')
    expect(shiftFormula('=Sheet2!$A$1+1', 1, 1)).toBe('=Sheet2!$A$1+1')
  })
  it('结构与字符串字面量不受影响', () => {
    expect(shiftFormula('="A1"&A1', 1, 1)).toBe('="A1"&B2')
    expect(shiftFormula('=IF(A1>0,"y","n")', 0, 1)).toBe('=IF(B1>0,"y","n")')
  })
  it('非公式与非法公式原文返回', () => {
    expect(shiftFormula('hello', 1, 1)).toBe('hello')
    expect(shiftFormula('123', 1, 1)).toBe('123')
    expect(shiftFormula('=A1+', 1, 1)).toBe('=A1+')
  })
  it('零偏移等价恒等', () => {
    expect(shiftFormula('=A1+B$2', 0, 0)).toBe('=A1+B$2')
  })
  it('含未知裸名（#NAME? err 节点）：改写后可再 parse 且语义不变', () => {
    const shifted = shiftFormula('=IFERROR(NOPE,A1)', 1, 0)
    expect(shifted).toBe('=IFERROR(#NAME?,A2)')
    expect(() => parseFormula(shifted.slice(1))).not.toThrow()
    // 语义不变：原公式与改写后公式求值一致（IFERROR 兜住 #NAME? 取第二参）
    let wb = Workbook.create({ rowCount: 5, colCount: 5 })
    let data = wb.activeSheet
    data = data.setCell(0, 0, { raw: '42' })
    data = data.setCell(1, 0, { raw: '42' })
    data = data.setCell(0, 1, { raw: '=IFERROR(NOPE,A1)' })
    data = data.setCell(0, 2, { raw: shifted })
    wb = wb.setSheet(wb.active, data)
    const ev = evaluatorFor(wb)
    expect(ev.get(wb.active, 0, 1)).toBe(42)
    expect(ev.get(wb.active, 0, 2)).toBe(42)
  })
})
