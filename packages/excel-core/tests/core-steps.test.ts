import { describe, it, expect } from 'vitest'
import { Workbook } from '../src/core/model'
import { SetCellsStep, PatchStyleStep, ResizeStep, stepFromJSON } from '../src/core/steps'

const wb = () => Workbook.create({ rowCount: 10, colCount: 5 })

describe('steps', () => {
  it('SetCellsStep apply/invert 往返', () => {
    const d0 = wb().setSheet('s1', wb().activeSheet.setCell(0, 0, { raw: 'old' }))
    const step = new SetCellsStep('s1', [
      { row: 0, col: 0, cell: { raw: 'new' } },
      { row: 1, col: 1, cell: { raw: 'x' } },
    ])
    const r = step.apply(d0)
    expect(r.ok).toBe(true)
    expect(r.doc!.sheet('s1').getCell(0, 0)).toEqual({ raw: 'new' })
    const inv = step.invert(d0)
    const r2 = inv.apply(r.doc!)
    expect(r2.ok).toBe(true)
    expect(r2.doc!.sheet('s1').getCell(0, 0)).toEqual({ raw: 'old' })
    expect(r2.doc!.sheet('s1').getCell(1, 1)).toBeUndefined()
  })
  it('SetCellsStep 越界 → failed', () => {
    const r = new SetCellsStep('s1', [{ row: 99, col: 0, cell: { raw: 'a' } }]).apply(wb())
    expect(r.ok).toBe(false)
    expect(r.failed).toBeTruthy()
  })
  it('PatchStyleStep 合并与删除键', () => {
    const d0 = wb().setSheet('s1', wb().activeSheet.setCell(0, 0, { raw: 'a', style: { bold: true, color: '#f00' } }))
    const step = new PatchStyleStep('s1', { sr: 0, sc: 0, er: 0, ec: 0 }, { color: undefined, italic: true })
    const r = step.apply(d0)
    expect(r.doc!.sheet('s1').getCell(0, 0)!.style).toEqual({ bold: true, italic: true })
    // invert 恢复
    const r2 = step.invert(d0).apply(r.doc!)
    expect(r2.doc!.sheet('s1').getCell(0, 0)!.style).toEqual({ bold: true, color: '#f00' })
  })
  it('PatchStyleStep 对空格打样式 → 建格；样式清空且 raw 空 → 删格', () => {
    const step = new PatchStyleStep('s1', { sr: 2, sc: 2, er: 2, ec: 2 }, { bg: '#ff0' })
    const r = step.apply(wb())
    expect(r.doc!.sheet('s1').getCell(2, 2)).toEqual({ raw: '', style: { bg: '#ff0' } })
    const r2 = new PatchStyleStep('s1', { sr: 2, sc: 2, er: 2, ec: 2 }, { bg: undefined }).apply(r.doc!)
    expect(r2.doc!.sheet('s1').getCell(2, 2)).toBeUndefined()
  })
  it('PatchStyleStep 撤销可恢复纯样式格（格被删后重建）', () => {
    // 空格 → patch bg（建格 {raw:'',style:{bg}}）→ patch bg:undefined（删格）→ undo 第二步 → 格恢复
    const step1 = new PatchStyleStep('s1', { sr: 2, sc: 2, er: 2, ec: 2 }, { bg: '#ff0' })
    const d1 = step1.apply(wb()).doc!
    const step2 = new PatchStyleStep('s1', { sr: 2, sc: 2, er: 2, ec: 2 }, { bg: undefined })
    const d2 = step2.apply(d1).doc!
    expect(d2.sheet('s1').getCell(2, 2)).toBeUndefined()
    const d3 = step2.invert(d1).apply(d2).doc!
    expect(d3.sheet('s1').getCell(2, 2)).toEqual({ raw: '', style: { bg: '#ff0' } })
  })
  it('ResizeStep 往返', () => {
    const d0 = wb()
    const step = new ResizeStep('s1', 'col', 1, 150)
    const r = step.apply(d0)
    expect(r.doc!.sheet('s1').colWidth(1)).toBe(150)
    const r2 = step.invert(d0).apply(r.doc!)
    expect(r2.doc!.sheet('s1').customColWidths.has(1)).toBe(false)
  })
  it('stepFromJSON 往返', () => {
    const steps = [
      new SetCellsStep('s1', [{ row: 0, col: 0, cell: { raw: 'a', style: { bold: true } } }]),
      new PatchStyleStep('s1', { sr: 0, sc: 0, er: 1, ec: 1 }, { italic: true }),
      new ResizeStep('s1', 'row', 3, null),
    ]
    for (const s of steps) {
      const back = stepFromJSON(JSON.parse(JSON.stringify(s.toJSON())))
      expect(back.toJSON()).toEqual(s.toJSON())
    }
  })
})
