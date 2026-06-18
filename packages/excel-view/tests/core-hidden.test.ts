import { describe, expect, it } from 'vitest'
import { SheetData, Workbook } from '@xexcel/core'
import { SetHiddenStep } from '@xexcel/core'
import { stepFromJSON } from '@xexcel/core'
import { StructureStep } from '@xexcel/core'
import { GridGeometry } from '../src/view/geometry'

function wb(): Workbook {
  return Workbook.create({ rowCount: 10, colCount: 10 })
}

describe('SheetData 隐藏行列', () => {
  it('setHidden 隐藏/恢复，数组有序去重', () => {
    let d = wb().activeSheet
    d = d.setHidden('row', [3, 1, 3], true)
    expect(d.hiddenRows).toEqual([1, 3])
    d = d.setHidden('row', [1], false)
    expect(d.hiddenRows).toEqual([3])
  })

  it('隐藏行高为 0、隐藏列宽为 0，自定义尺寸不丢失', () => {
    let d = wb().activeSheet.setRowHeight(2, 40)
    d = d.setHidden('row', [2], true)
    expect(d.rowHeight(2)).toBe(0)
    d = d.setHidden('row', [2], false)
    expect(d.rowHeight(2)).toBe(40)
    const d2 = wb().activeSheet.setHidden('col', [0], true)
    expect(d2.colWidth(0)).toBe(0)
  })

  it('toJSON/fromJSON 往返', () => {
    const d = wb().activeSheet.setHidden('row', [1, 2], true).setHidden('col', [0], true)
    const json = JSON.parse(JSON.stringify(d.toJSON()))
    const back = SheetData.fromJSON(json)
    expect(back.hiddenRows).toEqual([1, 2])
    expect(back.hiddenCols).toEqual([0])
  })

  it('结构操作重映射：插入平移、删除裁剪', () => {
    let d = wb().activeSheet.setHidden('row', [1, 5], true)
    d = d.insertRows(2, 2)
    expect(d.hiddenRows).toEqual([1, 7])
    d = wb().activeSheet.setHidden('row', [1, 2, 5], true).deleteRows(2, 2)
    expect(d.hiddenRows).toEqual([1, 3]) // 2 被删，5 → 3
    const c = wb().activeSheet.setHidden('col', [0, 4], true).deleteCols(0, 1)
    expect(c.hiddenCols).toEqual([3])
  })

  it('merges 之外字段不受影响（冻结保持）', () => {
    let d = wb().activeSheet.setFrozen(1, 1).setHidden('row', [3], true)
    d = d.deleteRows(0, 1)
    expect(d.hiddenRows).toEqual([2])
    expect(d.frozenRows).toBe(0) // 冻结边界随内容走：删除区内的冻结行被裁掉
  })

  it('delete rows 的 undo 整体恢复隐藏标记（wholesale）', () => {
    // 隐藏行 1（删除区外）与 5（删除区内）：delete 后 5 物理丢失，undo 须完整恢复
    const doc0 = wb().setSheet('s1', wb().activeSheet.setHidden('row', [1, 5], true))
    const step = new StructureStep({ sheet: 's1', axis: 'row', index: 4, count: 2, mode: 'delete' }, null)
    const r = step.apply(doc0)
    expect(r.doc!.sheet('s1').hiddenRows).toEqual([1])
    const back = step.invert(doc0).apply(r.doc!)
    expect(back.ok).toBe(true)
    expect(back.doc!.sheet('s1').hiddenRows).toEqual([1, 5])
  })

  it('insert rows 的 undo 不受影响（remap 自身可逆）', () => {
    const doc0 = wb().setSheet('s1', wb().activeSheet.setHidden('row', [1, 5], true))
    const step = new StructureStep({ sheet: 's1', axis: 'row', index: 2, count: 2, mode: 'insert' }, null)
    const r = step.apply(doc0)
    expect(r.doc!.sheet('s1').hiddenRows).toEqual([1, 7])
    const back = step.invert(doc0).apply(r.doc!)
    expect(back.doc!.sheet('s1').hiddenRows).toEqual([1, 5])
  })
})

describe('SetHiddenStep', () => {
  it('apply + invert 恒等（混合前置状态）', () => {
    const doc0 = wb().setSheet('s1', wb().activeSheet.setHidden('row', [2], true))
    const step = new SetHiddenStep('s1', 'row', [2, 5], true)
    const r = step.apply(doc0)
    expect(r.ok).toBe(true)
    expect(r.doc!.sheet('s1').hiddenRows).toEqual([2, 5])
    const back = step.invert(doc0).apply(r.doc!)
    expect(back.ok).toBe(true)
    expect(back.doc!.sheet('s1').hiddenRows).toEqual([2]) // 2 原本就隐藏，undo 不误恢复
  })

  it('取消隐藏的 undo 恢复隐藏', () => {
    const doc0 = wb().setSheet('s1', wb().activeSheet.setHidden('col', [1], true))
    const step = new SetHiddenStep('s1', 'col', [1], false)
    const r = step.apply(doc0)
    expect(r.doc!.sheet('s1').hiddenCols).toEqual([])
    const back = step.invert(doc0).apply(r.doc!)
    expect(back.doc!.sheet('s1').hiddenCols).toEqual([1])
  })

  it('toJSON → stepFromJSON 往返', () => {
    const step = new SetHiddenStep('s1', 'row', [1, 2], true)
    const json = JSON.parse(JSON.stringify(step.toJSON()))
    const back = stepFromJSON(json) as SetHiddenStep
    expect(back.sheet).toBe('s1')
    expect(back.indices).toEqual([1, 2])
    expect(back.hidden).toBe(true)
    const doc0 = wb()
    expect(back.apply(doc0).doc!.sheet('s1').hiddenRows).toEqual([1, 2])
  })

  it('越界 index → failed', () => {
    const r = new SetHiddenStep('s1', 'row', [99], true).apply(wb())
    expect(r.ok).toBe(false)
  })
})

describe('geometry 隐藏', () => {
  it('rowHeight/colWidth 访问器；隐藏行不可点击命中', () => {
    const d = wb().activeSheet.setHidden('row', [1], true)
    const g = new GridGeometry(d, 0, 0)
    expect(g.rowHeight(1)).toBe(0)
    expect(g.colWidth(0)).toBe(96)
    // 行 1 高度 0：rowAt(24) 落在行 2（隐藏行被跳过）
    expect(g.rowAt(24)).toBe(2)
    expect(g.rowAt(23)).toBe(0)
  })
})
