import { describe, it, expect } from 'vitest'
import { Workbook } from '../src/core/model'
import { SetFreezeStep, stepFromJSON } from '../src/core/steps'

describe('freeze', () => {
  it('setFrozen 不可变 + JSON 往返', () => {
    const s0 = Workbook.create({ rowCount: 10, colCount: 5 }).activeSheet
    const s1 = s0.setFrozen(2, 1)
    expect(s0.frozenRows).toBe(0)
    expect(s1.frozenRows).toBe(2)
    expect(s1.frozenCols).toBe(1)
    const wb = Workbook.create({ rowCount: 10, colCount: 5 }).setSheet('s1', s1)
    const back = Workbook.fromJSON(JSON.parse(JSON.stringify(wb.toJSON())))
    expect(back.activeSheet.frozenRows).toBe(2)
    expect(back.activeSheet.frozenCols).toBe(1)
  })
  it('SetFreezeStep apply/invert/toJSON 往返', () => {
    const d0 = Workbook.create({ rowCount: 10, colCount: 5 })
    const step = new SetFreezeStep('s1', 1, 2)
    const r = step.apply(d0)
    expect(r.doc!.activeSheet.frozenRows).toBe(1)
    expect(r.doc!.activeSheet.frozenCols).toBe(2)
    const r2 = step.invert(d0).apply(r.doc!)
    expect(r2.doc!.activeSheet.frozenRows).toBe(0)
    const back = stepFromJSON(JSON.parse(JSON.stringify(step.toJSON())))
    expect(back.toJSON()).toEqual(step.toJSON())
  })
  it('SetFreezeStep 越界 → failed', () => {
    const d0 = Workbook.create({ rowCount: 10, colCount: 5 })
    expect(new SetFreezeStep('s1', 10, 0).apply(d0).ok).toBe(false)
    expect(new SetFreezeStep('s1', -1, 0).apply(d0).ok).toBe(false)
  })
  it('delete 行列时冻结数钳到新尺寸；insert 不动', () => {
    const s0 = Workbook.create({ rowCount: 10, colCount: 5 }).activeSheet.setFrozen(9, 4)
    // 删行：9 → min(9, 10-6)=4；列方向不受影响
    const s1 = s0.deleteRows(4, 6)
    expect(s1.rowCount).toBe(4)
    expect(s1.frozenRows).toBe(4)
    expect(s1.frozenCols).toBe(4)
    // 删列：4 → min(4, 5-3)=2；行方向不受影响
    const s2 = s1.deleteCols(2, 3)
    expect(s2.colCount).toBe(2)
    expect(s2.frozenRows).toBe(4)
    expect(s2.frozenCols).toBe(2)
    // insert 不动冻结
    const s3 = s2.insertRows(1, 3).insertCols(0, 2)
    expect(s3.frozenRows).toBe(4)
    expect(s3.frozenCols).toBe(2)
    // 冻结数未越界时 delete 保持不变
    const s4 = Workbook.create({ rowCount: 10, colCount: 5 }).activeSheet.setFrozen(2, 1)
    const s5 = s4.deleteRows(4, 3).deleteCols(2, 2)
    expect(s5.frozenRows).toBe(2)
    expect(s5.frozenCols).toBe(1)
  })
})
