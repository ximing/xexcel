import { describe, it, expect } from 'vitest'
import { Workbook } from '../src/core/model'
import { evaluatorFor } from '../src/formula/engine'
import { parseFormula } from '../src/formula/parser'
import { adjustFormulaForStructure, StructureSpec } from '../src/formula/transform'

const ins = (index: number, count = 1): StructureSpec => ({ sheet: 'Sheet1', axis: 'row', index, count, mode: 'insert' })
const del = (index: number, count = 1): StructureSpec => ({ sheet: 'Sheet1', axis: 'row', index, count, mode: 'delete' })
const H = 'Sheet1' // 公式所在表

describe('adjustFormulaForStructure', () => {
  it('插入行：本表引用 ≥ index 平移，< index 不动', () => {
    expect(adjustFormulaForStructure('=A5*2', ins(4), H)).toBe('=A6*2')
    expect(adjustFormulaForStructure('=A4*2', ins(4), H)).toBe('=A4*2')
    expect(adjustFormulaForStructure('=SUM(A1:A10)', ins(2, 3), H)).toBe('=SUM(A1:A13)')
  })
  it('插入列：列维度平移', () => {
    const spec: StructureSpec = { sheet: 'Sheet1', axis: 'col', index: 1, count: 1, mode: 'insert' }
    expect(adjustFormulaForStructure('=B2*2', spec, H)).toBe('=C2*2')
    expect(adjustFormulaForStructure('=A2*2', spec, H)).toBe('=A2*2')
  })
  it('$ 锁定维度同样平移（Excel 同款）', () => {
    expect(adjustFormulaForStructure('=$A$5*2', ins(4), H)).toBe('=$A$6*2')
    expect(adjustFormulaForStructure('=A$5*2', ins(4), H)).toBe('=A$6*2')
  })
  it('删除行：删除区内 → #REF!，之后 → 前移', () => {
    expect(adjustFormulaForStructure('=A5*2', del(4), H)).toBe('=#REF!*2')
    expect(adjustFormulaForStructure('=A6*2', del(4), H)).toBe('=A5*2')
    expect(adjustFormulaForStructure('=A4*2', del(4), H)).toBe('=A4*2')
    expect(adjustFormulaForStructure('=SUM(A3:A6)', del(4, 2), H)).toBe('=SUM(#REF!)')
  })
  it('range 坍塌为局部 #REF!（不外溢到整个公式）', () => {
    expect(adjustFormulaForStructure('=SUM(A3:A6)+B1', del(4, 2), H)).toBe('=SUM(#REF!)+B1')
  })
  it('跨表引用：指向被改表才调整（按表名，不区分大小写）', () => {
    expect(adjustFormulaForStructure('=Sheet1!A5*2', ins(4), H)).toBe('=Sheet1!A6*2')
    expect(adjustFormulaForStructure('=Other!A5*2', ins(4), H)).toBe('=Other!A5*2')
    // 公式在别的表上，本表引用（无表名）不动
    expect(adjustFormulaForStructure('=A5*2', ins(4), 'Other')).toBe('=A5*2')
    expect(adjustFormulaForStructure('=sheet1!A5*2', ins(4), 'Other')).toBe('=sheet1!A6*2')
  })
  it('非公式与解析失败原文返回', () => {
    expect(adjustFormulaForStructure('hello', ins(0), H)).toBe('hello')
    expect(adjustFormulaForStructure('=A1+', ins(0), H)).toBe('=A1+')
  })
  it('含未知裸名（#NAME? err 节点）：改写后可再 parse 且语义不变', () => {
    const out = adjustFormulaForStructure('=IFERROR(NOPE,A5)', ins(4), H)
    expect(out).toBe('=IFERROR(#NAME?,A6)')
    expect(() => parseFormula(out.slice(1))).not.toThrow()
    // 语义不变：求值取平移后的 A6
    let wb = Workbook.create({ rowCount: 10, colCount: 3 })
    let data = wb.activeSheet
    data = data.setCell(5, 0, { raw: '7' })
    data = data.setCell(0, 1, { raw: out })
    wb = wb.setSheet(wb.active, data)
    expect(evaluatorFor(wb).get(wb.active, 0, 1)).toBe(7)
  })
})
