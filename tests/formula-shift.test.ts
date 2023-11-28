import { describe, it, expect } from 'vitest'
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
})
