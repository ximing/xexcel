import { describe, it, expect } from 'vitest'
import { Workbook } from '@gmi/excel-core'
import { evaluatorFor, isFormula } from '@gmi/excel-core'

const mk = (cells: Array<[number, number, string]>) => {
  let sheet = Workbook.create({ rowCount: 50, colCount: 26 }).activeSheet
  for (const [r, c, raw] of cells) sheet = sheet.setCell(r, c, { raw })
  return Workbook.create({ rowCount: 50, colCount: 26 }).setSheet('s1', sheet)
}
const val = (wb: Workbook, r: number, c: number) => evaluatorFor(wb).get('s1', r, c)
const txt = (wb: Workbook, r: number, c: number) => evaluatorFor(wb).displayText('s1', r, c)

describe('formula', () => {
  it('isFormula', () => {
    expect(isFormula('=1+1')).toBe(true)
    expect(isFormula('1+1')).toBe(false)
  })
  it('算术与优先级', () => {
    const wb = mk([[0, 0, '=1+2*3'], [0, 1, '=(1+2)*3'], [0, 2, '=2^3^2'], [0, 3, '=10/4'], [0, 4, '=-1^2'], [0, 5, '=50%+1'], [1, 0, '=-(1+2)']])
    expect(val(wb, 0, 0)).toBe(7)
    expect(val(wb, 0, 1)).toBe(9)
    expect(val(wb, 0, 2)).toBe(512) // ^ 右结合
    expect(val(wb, 0, 3)).toBe(2.5)
    expect(val(wb, 0, 4)).toBe(1)   // Excel 语义：(-1)^2
    expect(val(wb, 0, 5)).toBe(1.5)
    expect(val(wb, 1, 0)).toBe(-3)
  })
  it('拼接与比较', () => {
    const wb = mk([[0, 0, '="a"&"b"&1'], [0, 1, '=1<2'], [0, 2, '="a"="A"'], [0, 3, '=TRUE>""'], [0, 4, '=1<>1']])
    expect(val(wb, 0, 0)).toBe('ab1')
    expect(val(wb, 0, 1)).toBe(true)
    expect(val(wb, 0, 2)).toBe(true)  // 字符串比较不区分大小写
    expect(val(wb, 0, 3)).toBe(true)  // Excel 序：布尔>字符串
    expect(val(wb, 0, 4)).toBe(false)
  })
  it('单元格引用与区域函数', () => {
    const wb = mk([[0, 0, '10'], [1, 0, '20'], [2, 0, 'abc'], [3, 0, '=SUM(A1:A3)'], [3, 1, '=AVERAGE(A1:A3)'], [3, 2, '=COUNT(A1:A3)'], [3, 3, '=MAX(A1:A3)'], [3, 4, '=MIN(A1:A3)'], [4, 0, '=A1*2'], [4, 1, '=SUM(A1,5,A2)']])
    expect(val(wb, 3, 0)).toBe(30)
    expect(val(wb, 3, 1)).toBe(15)
    expect(val(wb, 3, 2)).toBe(2)
    expect(val(wb, 3, 3)).toBe(20)
    expect(val(wb, 3, 4)).toBe(10)
    expect(val(wb, 4, 0)).toBe(20)
    expect(val(wb, 4, 1)).toBe(35)
  })
  it('ABS / ROUND / IF', () => {
    const wb = mk([[0, 0, '=ABS(-3)'], [0, 1, '=ROUND(3.14159,2)'], [0, 2, '=ROUND(2.5)'], [0, 3, '=IF(1>0,"yes","no")'], [0, 4, '=IF(FALSE,1,2)']])
    expect(val(wb, 0, 0)).toBe(3)
    expect(val(wb, 0, 1)).toBe(3.14)
    expect(val(wb, 0, 2)).toBe(3)
    expect(val(wb, 0, 3)).toBe('yes')
    expect(val(wb, 0, 4)).toBe(2)
  })
  it('错误：DIV/0、NAME?、VALUE!、REF!、CYCLE!', () => {
    const wb = mk([
      [0, 0, '=1/0'],
      [0, 1, '=NOSUCH(1)'],
      [0, 2, '=1+"abc"'],
      [0, 3, '=A100'],          // row 99 越界（50 行表）
      [0, 4, '=E2'],            // 与 E2 互相引用
      [1, 4, '=E1'],
      [1, 0, '=A1+1'],          // 错误传播：A1 是 #DIV/0!
    ])
    expect(val(wb, 0, 0)).toEqual({ error: '#DIV/0!' })
    expect(val(wb, 0, 1)).toEqual({ error: '#NAME?' })
    expect(val(wb, 0, 2)).toEqual({ error: '#VALUE!' })
    expect(val(wb, 0, 3)).toEqual({ error: '#REF!' })
    expect(val(wb, 0, 4)).toEqual({ error: '#CYCLE!' })
    expect(val(wb, 1, 4)).toEqual({ error: '#CYCLE!' })
    expect(val(wb, 1, 0)).toEqual({ error: '#DIV/0!' })
  })
  it('非公式单元格类型转换与空单元格', () => {
    const wb = mk([[0, 0, '3.5'], [0, 1, '50%'], [0, 2, 'abc'], [0, 3, '=-1e3']])
    expect(val(wb, 0, 0)).toBe(3.5)
    expect(val(wb, 0, 1)).toBe(0.5)
    expect(val(wb, 0, 2)).toBe('abc')
    expect(val(wb, 5, 5)).toBe('')   // 空
    expect(val(wb, 0, 3)).toBe(-1000)
  })
  it('空单元格在算术中按 0', () => {
    const wb = mk([[0, 0, '=B5+1']])
    expect(val(wb, 0, 0)).toBe(1)
  })
  it('空单元格在比较中按 0', () => {
    const wb = mk([[0, 0, '=B5=0'], [0, 1, '=B5<1'], [0, 2, '=B5>1'], [0, 3, '=B5&""']])
    expect(val(wb, 0, 0)).toBe(true)
    expect(val(wb, 0, 1)).toBe(true)
    expect(val(wb, 0, 2)).toBe(false)
    expect(val(wb, 0, 3)).toBe('') // 空格拼接为空串，不受影响
  })
  it('displayText', () => {
    const wb = mk([[0, 0, '=1/3'], [0, 1, '=TRUE'], [0, 2, '=1/0'], [0, 3, '42'], [1, 0, '=0.1+0.2']])
    expect(txt(wb, 0, 1)).toBe('TRUE')
    expect(txt(wb, 0, 2)).toBe('#DIV/0!')
    expect(txt(wb, 0, 3)).toBe('42')
    expect(txt(wb, 1, 0)).toBe('0.3') // 长度>12 → toPrecision(10) 去尾零
  })
  it('evaluatorFor 按 doc 缓存', () => {
    const wb = mk([[0, 0, '=1+1']])
    expect(evaluatorFor(wb)).toBe(evaluatorFor(wb))
  })
})
