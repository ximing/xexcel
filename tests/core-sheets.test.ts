import { describe, it, expect } from 'vitest'
import { nextSheetId, nextSheetName, SheetData, Workbook } from '../src/core/model'
import {
  InsertSheetStep,
  RemoveSheetStep,
  RenameSheetStep,
  SetActiveSheetStep,
  stepFromJSON,
} from '../src/core/steps'

const wb = () => Workbook.create({ rowCount: 10, colCount: 5 })

describe('sheet steps', () => {
  it('InsertSheetStep：插入并设为 active；undo 移除且恢复 active', () => {
    const d0 = wb()
    const step = new InsertSheetStep('s2', 'Data', SheetData.create({ rowCount: 10, colCount: 5 }), null, 's2')
    const r = step.apply(d0)
    expect(r.ok).toBe(true)
    expect(r.doc!.order).toEqual(['s1', 's2'])
    expect(r.doc!.active).toBe('s2')
    expect(r.doc!.names.get('s2')).toBe('Data')
    const r2 = step.invert(d0).apply(r.doc!)
    expect(r2.ok).toBe(true)
    expect(r2.doc!.order).toEqual(['s1'])
    expect(r2.doc!.active).toBe('s1')
  })
  it('InsertSheetStep：id 已存在 → failed', () => {
    const r = new InsertSheetStep('s1', 'X', SheetData.create({ rowCount: 1, colCount: 1 }), null, null).apply(wb())
    expect(r.ok).toBe(false)
  })
  it('RemoveSheetStep：删除 active 表 → active 移到相邻；undo 恢复数据/名称/位置/active', () => {
    let d0 = wb()
    d0 = d0.setSheet('s1', d0.sheet('s1').setCell(2, 2, { raw: 'keep' }))
    const ins = new InsertSheetStep('s2', 'Data', SheetData.create({ rowCount: 10, colCount: 5 }), null, 's2')
    d0 = ins.apply(d0).doc!
    const step = new RemoveSheetStep('s2')
    const r = step.apply(d0)
    expect(r.ok).toBe(true)
    expect(r.doc!.order).toEqual(['s1'])
    expect(r.doc!.active).toBe('s1')
    // 先在 s2 写点数据再删，验证快照恢复
    let d1 = ins.apply(wb()).doc!
    d1 = d1.setSheet('s2', d1.sheet('s2').setCell(1, 1, { raw: '42' }))
    const del = new RemoveSheetStep('s2')
    const d2 = del.apply(d1).doc!
    const d3 = del.invert(d1).apply(d2).doc!
    expect(d3.sheet('s2').getCell(1, 1)).toEqual({ raw: '42' })
    expect(d3.names.get('s2')).toBe('Data')
    expect(d3.order).toEqual(['s1', 's2'])
    expect(d3.active).toBe('s2')
  })
  it('RemoveSheetStep：删除最后一张表 → failed', () => {
    const r = new RemoveSheetStep('s1').apply(wb())
    expect(r.ok).toBe(false)
    expect(r.failed).toContain('last sheet')
  })
  it('RenameSheetStep：改名/undo；空名与重名（不区分大小写）→ failed', () => {
    const d0 = wb()
    const step = new RenameSheetStep('s1', 'Budget')
    const r = step.apply(d0)
    expect(r.ok).toBe(true)
    expect(r.doc!.names.get('s1')).toBe('Budget')
    const r2 = step.invert(d0).apply(r.doc!)
    expect(r2.doc!.names.get('s1')).toBe('Sheet1')
    expect(new RenameSheetStep('s1', '  ').apply(d0).ok).toBe(false)
    const ins = new InsertSheetStep('s2', 'Data', SheetData.create({ rowCount: 10, colCount: 5 }), null, null)
    const d1 = ins.apply(d0).doc!
    expect(new RenameSheetStep('s2', 'sheet1').apply(d1).ok).toBe(false)
  })
  it('SetActiveSheetStep：切换与 undo；未知表 → failed', () => {
    const ins = new InsertSheetStep('s2', 'Data', SheetData.create({ rowCount: 10, colCount: 5 }), null, null)
    const d0 = ins.apply(wb()).doc! // active 仍为 s1
    const step = new SetActiveSheetStep('s2')
    const r = step.apply(d0)
    expect(r.doc!.active).toBe('s2')
    expect(step.invert(d0).apply(r.doc!).doc!.active).toBe('s1')
    expect(new SetActiveSheetStep('nope').apply(d0).ok).toBe(false)
  })
  it('stepFromJSON 往返', () => {
    const steps = [
      new InsertSheetStep('s2', 'Data', SheetData.create({ rowCount: 3, colCount: 2 }).setCell(0, 0, { raw: 'x' }), 1, 's2'),
      new RemoveSheetStep('s2', 's1'),
      new RenameSheetStep('s1', 'Budget'),
      new SetActiveSheetStep('s2'),
    ]
    for (const s of steps) {
      const back = stepFromJSON(JSON.parse(JSON.stringify(s.toJSON())))
      expect(back.toJSON()).toEqual(s.toJSON())
    }
  })
  it('nextSheetId / nextSheetName', () => {
    const d0 = wb()
    expect(nextSheetId(d0)).toBe('s2')
    expect(nextSheetName(d0)).toBe('Sheet2')
    const d1 = new InsertSheetStep('s2', 'Sheet2', SheetData.create({ rowCount: 1, colCount: 1 }), null, null).apply(d0).doc!
    const d2 = new RemoveSheetStep('s1').apply(d1).doc! // 只剩 s2(Sheet2)
    expect(nextSheetId(d2)).toBe('s3')
    expect(nextSheetName(d2)).toBe('Sheet3')
  })
})
