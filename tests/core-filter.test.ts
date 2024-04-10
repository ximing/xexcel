import { describe, expect, it } from 'vitest'
import { FilterState, SheetData, Workbook } from '../src/core/model'
import { SetFilterStep, stepFromJSON } from '../src/core/steps'

const filter: FilterState = {
  range: { sr: 0, sc: 0, er: 9, ec: 2 },
  criteria: { 1: { type: 'values', excluded: ['x'] } },
}

function wbWith(f?: FilterState): Workbook {
  const wb = Workbook.create({ rowCount: 20, colCount: 10 })
  return wb.setSheet(wb.active, wb.activeSheet.setFilter(f))
}

describe('SheetData.filter', () => {
  it('setFilter 设置与清除', () => {
    const d = wbWith(filter).activeSheet
    expect(d.filter).toEqual(filter)
    expect(d.setFilter(undefined).filter).toBeUndefined()
  })

  it('toJSON/fromJSON 往返', () => {
    const d = wbWith(filter).activeSheet
    const back = SheetData.fromJSON(JSON.parse(JSON.stringify(d.toJSON())))
    expect(back.filter).toEqual(filter)
    expect(SheetData.fromJSON(JSON.parse(JSON.stringify(wbWith().activeSheet.toJSON()))).filter).toBeUndefined()
  })

  it('结构操作：行插入平移 range', () => {
    const d = wbWith(filter).activeSheet.insertRows(2, 2)
    expect(d.filter!.range).toEqual({ sr: 0, sc: 0, er: 11, ec: 2 })
    expect(d.filter!.criteria).toEqual(filter.criteria)
  })

  it('结构操作：删除表头行 → 筛选移除', () => {
    const d = wbWith(filter).activeSheet.deleteRows(0, 1)
    expect(d.filter).toBeUndefined()
  })

  it('结构操作：删除数据行 → range 裁剪', () => {
    const d = wbWith(filter).activeSheet.deleteRows(8, 2)
    expect(d.filter!.range.er).toBe(7)
  })

  it('结构操作：列删除重映射 criteria 键并裁剪', () => {
    const f2: FilterState = { range: { sr: 0, sc: 0, er: 9, ec: 3 }, criteria: { 1: { type: 'values', excluded: ['x'] }, 3: { type: 'condition', field: 'num', op: 'gt', v1: '5' } } }
    const d = wbWith(f2).activeSheet.deleteCols(1, 1)
    expect(d.filter!.range.ec).toBe(2)
    expect(Object.keys(d.filter!.criteria)).toEqual(['2']) // 键 1 被删，3 → 2
  })

  it('结构操作：列插入平移 criteria 键', () => {
    const d = wbWith(filter).activeSheet.insertCols(0, 1)
    expect(d.filter!.range).toEqual({ sr: 0, sc: 1, er: 9, ec: 3 })
    expect(Object.keys(d.filter!.criteria)).toEqual(['2'])
  })
})

describe('SetFilterStep', () => {
  it('apply + invert 恢复', () => {
    const doc0 = wbWith()
    const step = new SetFilterStep('s1', filter)
    const r = step.apply(doc0)
    expect(r.ok).toBe(true)
    expect(r.doc!.sheet('s1').filter).toEqual(filter)
    const back = step.invert(doc0).apply(r.doc!)
    expect(back.doc!.sheet('s1').filter).toBeUndefined()
  })

  it('toJSON → stepFromJSON 往返', () => {
    const step = new SetFilterStep('s1', filter)
    const back = stepFromJSON(JSON.parse(JSON.stringify(step.toJSON()))) as SetFilterStep
    expect(back.filter).toEqual(filter)
    const doc0 = wbWith()
    expect(back.apply(doc0).doc!.sheet('s1').filter).toEqual(filter)
  })
})
