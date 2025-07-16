import { describe, expect, it } from 'vitest'
import { Workbook } from '../src/core/model'
import { MoveSheetStep, stepFromJSON } from '../src/core/steps'

const wb3 = (): Workbook => {
  let wb = Workbook.create({ rowCount: 5, colCount: 5 })
  wb = wb.addSheet('s2', wb.sheet('s1'))
  wb = wb.addSheet('s3', wb.sheet('s1'))
  return wb
}

describe('MoveSheetStep', () => {
  it('apply：前移/后移', () => {
    expect(new MoveSheetStep('s3', 0).apply(wb3()).doc!.order).toEqual(['s3', 's1', 's2'])
    expect(new MoveSheetStep('s1', 2).apply(wb3()).doc!.order).toEqual(['s2', 's3', 's1'])
  })
  it('invert 恢复原位；active 不变', () => {
    const wb = wb3()
    const step = new MoveSheetStep('s1', 2)
    const after = step.apply(wb).doc!
    expect(step.invert(wb).apply(after).doc!.order).toEqual(['s1', 's2', 's3'])
    expect(after.active).toBe('s1')
  })
  it('越界/缺表失败', () => {
    expect(new MoveSheetStep('s1', 3).apply(wb3()).ok).toBe(false)
    expect(new MoveSheetStep('s9', 0).apply(wb3()).ok).toBe(false)
  })
  it('toJSON → stepFromJSON 往返', () => {
    const s = new MoveSheetStep('s2', 0)
    const back = stepFromJSON(JSON.parse(JSON.stringify(s.toJSON())))
    expect(back).toBeInstanceOf(MoveSheetStep)
    expect(back.apply(wb3()).doc!.order).toEqual(['s2', 's1', 's3'])
  })
})
