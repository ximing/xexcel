// tests/formula-intersection.test.ts
// 隐式交集：标量位置的 range 按公式自身坐标取交点（Excel 经典语义）
import { describe, expect, it } from 'vitest'
import { fromA1 } from '../src/core/addr'
import { Workbook } from '../src/core/model'
import { evaluatorFor } from '../src/formula/engine'

function wbWith(cells: Record<string, string>): Workbook {
  let wb = Workbook.create({ rowCount: 20, colCount: 10 })
  let data = wb.activeSheet
  for (const [a1, raw] of Object.entries(cells)) {
    const { row, col } = fromA1(a1)!
    data = data.setCell(row, col, { raw })
  }
  return wb.setSheet(wb.active, data)
}

function get(wb: Workbook, a1: string): unknown {
  const { row, col } = fromA1(a1)!
  return evaluatorFor(wb).get(wb.active, row, col)
}

// 双表工作簿：cells1 在 s1（active），cells2 在 Sheet2（id s2）
function wbTwo(cells1: Record<string, string>, cells2: Record<string, string>): Workbook {
  let wb = wbWith(cells1)
  wb = wb.addSheet('s2', Workbook.create({ rowCount: 20, colCount: 10 }).activeSheet, 1, 'Sheet2')
  let d2 = wb.sheet('s2')
  for (const [a1, raw] of Object.entries(cells2)) {
    const { row, col } = fromA1(a1)!
    d2 = d2.setCell(row, col, { raw })
  }
  return wb.setSheet('s2', d2)
}

describe('隐式交集', () => {
  it('单列区域：公式行相交取交点', () => {
    const wb = wbWith({ A1: '1', A2: '2', A3: '3', B2: '=A1:A3' })
    expect(get(wb, 'B2')).toBe(2)
  })

  it('单列区域：无交集 → #VALUE!', () => {
    const wb = wbWith({ A1: '1', A2: '2', A3: '3', B5: '=A1:A3' })
    expect(get(wb, 'B5')).toEqual({ error: '#VALUE!' })
  })

  it('单行区域：公式列相交取交点', () => {
    const wb = wbWith({ A1: '1', B1: '2', C1: '3', C4: '=A1:C1' })
    expect(get(wb, 'C4')).toBe(3)
  })

  it('单行区域：列不相交 → #VALUE!', () => {
    const wb = wbWith({ A1: '1', B1: '2', C1: '3', E4: '=A1:C1' })
    expect(get(wb, 'E4')).toEqual({ error: '#VALUE!' })
  })

  it('二维区域：公式在区域外 → #VALUE!', () => {
    const wb = wbWith({ A1: '1', B2: '2', D5: '=A1:B2' })
    expect(get(wb, 'D5')).toEqual({ error: '#VALUE!' })
  })

  it('二维区域：行列均在区域内 → 取 (公式行, 公式列) 格值（跨表，避免自引用）', () => {
    const wb = wbTwo({ B2: '=Sheet2!A1:B2' }, { A1: '1', B1: '2', A2: '3', B2: '4' })
    expect(get(wb, 'B2')).toBe(4) // 公式在 (1,1) → 取 Sheet2!B2
  })

  it('二维区域同表：交点即公式格自身 → #CYCLE!', () => {
    const wb = wbWith({ A1: '1', B2: '=A1:B2' })
    expect(get(wb, 'B2')).toEqual({ error: '#CYCLE!' })
  })

  it('跨表单列区域隐式交集（=Sheet2!C1:C3 按公式行取）', () => {
    const wb = wbTwo({ A2: '=Sheet2!C1:C3' }, { C1: '10', C2: '20', C3: '30' })
    expect(get(wb, 'A2')).toBe(20) // 公式行 2 → Sheet2!C2
  })

  it('交点为空格：BLANK 语义（算术按 0）', () => {
    const wb = wbWith({ A1: '1', A3: '3', B2: '=A1:A3+1' })
    expect(get(wb, 'B2')).toBe(1)
  })

  it('交点为空格：单独引用显示空串', () => {
    const wb = wbWith({ A1: '1', A3: '3', B2: '=A1:A3' })
    expect(get(wb, 'B2')).toBe('')
  })

  it('聚合参数中的区域不受影响', () => {
    const wb = wbWith({ A1: '1', A2: '2', A3: '3', B1: '=SUM(A1:A3)' })
    expect(get(wb, 'B1')).toBe(6)
  })

  it('区域可参与算术与比较', () => {
    const wb = wbWith({ A1: '5', A2: '10', B2: '=A1:A3*2', C2: '=IF(A1:A3>5,"big","small")' })
    expect(get(wb, 'B2')).toBe(20)
    expect(get(wb, 'C2')).toBe('big')
  })
})
