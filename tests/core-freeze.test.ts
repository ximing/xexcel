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
})
