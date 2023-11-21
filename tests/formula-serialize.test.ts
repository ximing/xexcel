import { describe, it, expect } from 'vitest'
import { parseFormula } from '../src/formula/parser'
import { serialize, serializeSheetName } from '../src/formula/serialize'

// round-trip：serialize(parse(x)) 再 parse 应与 parse(x) 结构等价
const rt = (x: string) => expect(parseFormula(serialize(parseFormula(x)))).toEqual(parseFormula(x))

describe('serialize', () => {
  it('原子与运算', () => {
    expect(serialize(parseFormula('1+2*3'))).toBe('1+2*3')
    expect(serialize(parseFormula('(1+2)*3'))).toBe('(1+2)*3')
    expect(serialize(parseFormula('"a"&"b"'))).toBe('"a"&"b"')
    expect(serialize(parseFormula('50%'))).toBe('50%')
    expect(serialize(parseFormula('-A1^2'))).toBe('-A1^2')
  })
  it('字符串转义', () => {
    expect(serialize(parseFormula('"a""b"'))).toBe('"a""b"')
  })
  it('$ 引用与表名', () => {
    expect(serialize(parseFormula('$A$1'))).toBe('$A$1')
    expect(serialize(parseFormula('$A1'))).toBe('$A1')
    expect(serialize(parseFormula('A$1'))).toBe('A$1')
    expect(serialize(parseFormula('Sheet2!A1'))).toBe('Sheet2!A1')
    expect(serialize(parseFormula("'My Sheet'!A1"))).toBe("'My Sheet'!A1")
    expect(serialize(parseFormula('Sheet2!A1:B2'))).toBe('Sheet2!A1:B2')
  })
  it('err 节点', () => {
    expect(serialize(parseFormula('#REF!'))).toBe('#REF!')
  })
  it('serializeSheetName：安全名不引号，其余引号并转义', () => {
    expect(serializeSheetName('Sheet2')).toBe('Sheet2')
    expect(serializeSheetName('My Sheet')).toBe("'My Sheet'")
    expect(serializeSheetName("It's")).toBe("'It''s'")
    expect(serializeSheetName('2024')).toBe("'2024'")
  })
  it('round-trip 结构等价', () => {
    for (const x of [
      '1+2*3',
      '(1+2)*3',
      'SUM(A1:B2,5)',
      'IF(A1>0,"y","n")',
      '$A$1+Sheet2!$B2',
      "'My Sheet'!A1:B2",
      '-(1+2)',
      '2^3^2',
      'A1<=B2',
      '#REF!',
    ]) {
      rt(x)
    }
  })
})
