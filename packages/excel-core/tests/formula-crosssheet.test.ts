import { describe, it, expect } from 'vitest'
import { SheetData, Workbook } from '../src/core/model'
import { evaluatorFor } from '../src/formula/engine'

// 两表工作簿：s1=Sheet1（活动），s2 改名 Data；Data!A1=10
const mk = () => {
  let wb = Workbook.create({ rowCount: 10, colCount: 5 })
  wb = wb.addSheet('s2', SheetData.create({ rowCount: 10, colCount: 5 }))
  wb = wb.renameSheet('s2', 'Data')
  wb = wb.setSheet('s2', wb.sheet('s2').setCell(0, 0, { raw: '10' }))
  return wb
}
const val = (wb: Workbook, r: number, c: number) => evaluatorFor(wb).get('s1', r, c)

describe('cross-sheet eval', () => {
  it('跨表引用求值', () => {
    let wb = mk()
    wb = wb.setSheet('s1', wb.sheet('s1').setCell(0, 0, { raw: '=Data!A1*2' }))
    expect(val(wb, 0, 0)).toBe(20)
  })
  it('表名不区分大小写', () => {
    let wb = mk()
    wb = wb.setSheet('s1', wb.sheet('s1').setCell(0, 0, { raw: '=data!a1*2' }))
    expect(val(wb, 0, 0)).toBe(20)
  })
  it('跨表区域聚合', () => {
    let wb = mk()
    wb = wb.setSheet('s2', wb.sheet('s2').setCell(1, 0, { raw: '20' }))
    wb = wb.setSheet('s1', wb.sheet('s1').setCell(0, 0, { raw: '=SUM(Data!A1:A2)' }))
    expect(val(wb, 0, 0)).toBe(30)
  })
  it('未知表名 → #REF!', () => {
    let wb = mk()
    wb = wb.setSheet('s1', wb.sheet('s1').setCell(0, 0, { raw: '=Nope!A1' }))
    expect(val(wb, 0, 0)).toEqual({ error: '#REF!' })
  })
  it('表改名后旧名引用 → #REF!（已知限制：不改写公式）', () => {
    let wb = mk()
    wb = wb.setSheet('s1', wb.sheet('s1').setCell(0, 0, { raw: '=Data!A1' }))
    wb = wb.renameSheet('s2', 'Other')
    expect(val(wb, 0, 0)).toEqual({ error: '#REF!' })
  })
  it('跨表循环引用 → 链上全部 #CYCLE!', () => {
    let wb = mk()
    wb = wb.setSheet('s1', wb.sheet('s1').setCell(0, 0, { raw: '=Data!B1+1' }))
    wb = wb.setSheet('s2', wb.sheet('s2').setCell(0, 1, { raw: '=Sheet1!A1+1' }))
    expect(val(wb, 0, 0)).toEqual({ error: '#CYCLE!' })
    expect(evaluatorFor(wb).get('s2', 0, 1)).toEqual({ error: '#CYCLE!' })
  })
  it('#REF! 字面量参与运算 → 错误传播', () => {
    let wb = mk()
    wb = wb.setSheet('s1', wb.sheet('s1').setCell(0, 0, { raw: '=#REF!+1' }))
    expect(val(wb, 0, 0)).toEqual({ error: '#REF!' })
  })
  it('带引号表名求值', () => {
    let wb = mk()
    wb = wb.renameSheet('s2', 'My Data')
    wb = wb.setSheet('s1', wb.sheet('s1').setCell(0, 0, { raw: "='My Data'!A1+1" }))
    expect(val(wb, 0, 0)).toBe(11)
  })
})
