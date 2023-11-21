import { describe, it, expect } from 'vitest'
import { parseFormula } from '../src/formula/parser'

describe('formula refs: $ 与表名', () => {
  it('$ 各形态解析', () => {
    expect(parseFormula('$A$1')).toEqual({
      type: 'ref',
      ref: { row: 0, col: 0, rowAbs: true, colAbs: true },
    })
    expect(parseFormula('$A1')).toEqual({
      type: 'ref',
      ref: { row: 0, col: 0, rowAbs: false, colAbs: true },
    })
    expect(parseFormula('A$1')).toEqual({
      type: 'ref',
      ref: { row: 0, col: 0, rowAbs: true, colAbs: false },
    })
    expect(parseFormula('A1')).toEqual({
      type: 'ref',
      ref: { row: 0, col: 0, rowAbs: false, colAbs: false },
    })
  })
  it('表名前缀：裸名与引号名', () => {
    expect(parseFormula('Sheet2!A1')).toEqual({
      type: 'ref',
      ref: { sheet: 'Sheet2', row: 0, col: 0, rowAbs: false, colAbs: false },
    })
    expect(parseFormula("'My Sheet'!$A$1")).toEqual({
      type: 'ref',
      ref: { sheet: 'My Sheet', row: 0, col: 0, rowAbs: true, colAbs: true },
    })
    // 引号内 '' 转义为单引号
    expect(parseFormula("'It''s'!A1")).toEqual({
      type: 'ref',
      ref: { sheet: "It's", row: 0, col: 0, rowAbs: false, colAbs: false },
    })
  })
  it('区域带表名；异表区域报语法错', () => {
    expect(parseFormula('Sheet2!A1:B2')).toEqual({
      type: 'range',
      a: { sheet: 'Sheet2', row: 0, col: 0, rowAbs: false, colAbs: false },
      b: { sheet: 'Sheet2', row: 1, col: 1, rowAbs: false, colAbs: false },
    })
    // 区域两端可带 $，且不归一（保留书写方向）
    expect(parseFormula('$B2:A$1')).toEqual({
      type: 'range',
      a: { row: 1, col: 1, rowAbs: false, colAbs: true },
      b: { row: 0, col: 0, rowAbs: true, colAbs: false },
    })
    expect(() => parseFormula('S1!A1:S2!B2')).toThrow()
  })
  it('#REF! 字面量解析为 err 节点', () => {
    expect(parseFormula('#REF!')).toEqual({ type: 'err', error: '#REF!' })
    expect(parseFormula('=#REF!+1'.slice(1))).toEqual({
      type: 'binary',
      op: '+',
      left: { type: 'err', error: '#REF!' },
      right: { type: 'num', value: 1 },
    })
  })
  it('既有语法不回归：函数/优先级/paren', () => {
    expect(parseFormula('SUM(A1:B2,5)').type).toBe('call')
    expect(parseFormula('-(1+2)')).toEqual({
      type: 'unary',
      op: '-',
      expr: {
        type: 'paren',
        expr: {
          type: 'binary',
          op: '+',
          left: { type: 'num', value: 1 },
          right: { type: 'num', value: 2 },
        },
      },
    })
  })
})
