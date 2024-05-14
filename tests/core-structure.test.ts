import { describe, it, expect } from 'vitest'
import '../src/formula/transform' // 触发 cascade 注入（registerStructureCascade）
import { FilterState, Workbook } from '../src/core/model'
import { StructureStep, stepFromJSON } from '../src/core/steps'

const mk = () => {
  let wb = Workbook.create({ rowCount: 10, colCount: 5 })
  let s = wb.activeSheet
  s = s.setCell(4, 0, { raw: '50' }) // A5
  s = s.setCell(0, 0, { raw: '=A5*2' }) // A1
  s = s.setCell(1, 0, { raw: '=$A$5+1' }) // A2
  s = s.setRowHeight(7, 40)
  return wb.setSheet('s1', s)
}

describe('SheetData 结构操作', () => {
  it('insertRows：cells/rowHeights/merges 平移，rowCount 增加', () => {
    const d = mk().activeSheet.insertRows(4, 2)
    expect(d.rowCount).toBe(12)
    expect(d.getCell(6, 0)?.raw).toBe('50')
    expect(d.customRowHeights.get(9)).toBe(40)
    expect(d.customRowHeights.has(7)).toBe(false)
  })
  it('deleteRows：删除区移除，其后前移', () => {
    const d = mk().activeSheet.deleteRows(4, 1)
    expect(d.rowCount).toBe(9)
    expect(d.getCell(4, 0)).toBeUndefined()
    expect(d.customRowHeights.get(6)).toBe(40)
  })
  it('deleteRows：merges 裁剪/移除', () => {
    const s = mk().activeSheet.setMerges([
      { sr: 3, sc: 0, er: 6, ec: 0 }, // 与删除区 [4,6) 相交
      { sr: 4, sc: 1, er: 5, ec: 1 }, // 完全在删除区内
      { sr: 0, sc: 2, er: 1, ec: 2 }, // 不受影响
    ])
    const d = s.deleteRows(4, 2)
    expect(d.merges).toEqual([
      { sr: 3, sc: 0, er: 4, ec: 0 },
      { sr: 0, sc: 2, er: 1, ec: 2 },
    ])
  })
  it('insertCols/deleteCols 对称', () => {
    const d = mk().activeSheet.insertCols(0, 1)
    expect(d.colCount).toBe(6)
    expect(d.getCell(4, 1)?.raw).toBe('50')
    const d2 = d.deleteCols(0, 1)
    expect(d2.colCount).toBe(5)
    expect(d2.getCell(4, 0)?.raw).toBe('50')
  })
})

describe('StructureStep', () => {
  it('插入行：公式级联平移（含 $），undo 恢复', () => {
    const d0 = mk()
    const step = new StructureStep({ sheet: 's1', axis: 'row', index: 4, count: 1, mode: 'insert' }, null)
    const r = step.apply(d0)
    expect(r.ok).toBe(true)
    expect(r.doc!.activeSheet.getCell(0, 0)?.raw).toBe('=A6*2')
    expect(r.doc!.activeSheet.getCell(1, 0)?.raw).toBe('=$A$6+1')
    expect(r.doc!.activeSheet.getCell(5, 0)?.raw).toBe('50')
    const r2 = step.invert(d0).apply(r.doc!)
    expect(r2.doc!.activeSheet.getCell(0, 0)?.raw).toBe('=A5*2')
    expect(r2.doc!.activeSheet.getCell(4, 0)?.raw).toBe('50')
    expect(r2.doc!.activeSheet.rowCount).toBe(10)
  })
  it('删除行：引用变 #REF!，undo 恢复原文', () => {
    const d0 = mk()
    const step = new StructureStep({ sheet: 's1', axis: 'row', index: 4, count: 1, mode: 'delete' }, null)
    const r = step.apply(d0)
    expect(r.doc!.activeSheet.getCell(0, 0)?.raw).toBe('=#REF!*2')
    expect(r.doc!.activeSheet.getCell(1, 0)?.raw).toBe('=#REF!+1')
    const r2 = step.invert(d0).apply(r.doc!)
    expect(r2.doc!.activeSheet.getCell(0, 0)?.raw).toBe('=A5*2')
    expect(r2.doc!.activeSheet.getCell(1, 0)?.raw).toBe('=$A$5+1')
    expect(r2.doc!.activeSheet.getCell(4, 0)?.raw).toBe('50')
  })
  it('删除行的 undo：恢复删除区行高与合并区（含完全在删除区内的 merge）', () => {
    let wb = Workbook.create({ rowCount: 10, colCount: 5 })
    let s = wb.activeSheet
    s = s.setCell(4, 0, { raw: 'x' })
    s = s.setRowHeight(4, 40)
    s = s.setRowHeight(5, 30)
    s = s.setMerges([
      { sr: 4, sc: 1, er: 5, ec: 1 }, // 完全在删除区 [4,6) 内
      { sr: 3, sc: 2, er: 5, ec: 2 }, // 部分重叠（尾部入删除区）
      { sr: 0, sc: 3, er: 1, ec: 3 }, // 不受影响
    ])
    wb = wb.setSheet('s1', s)
    const step = new StructureStep({ sheet: 's1', axis: 'row', index: 4, count: 2, mode: 'delete' }, null)
    const r = step.apply(wb)
    expect(r.ok).toBe(true)
    // 正向：删除区行高丢失、merge 裁剪/移除
    expect(r.doc!.activeSheet.customRowHeights.has(4)).toBe(false)
    expect(r.doc!.activeSheet.merges).toEqual([
      { sr: 3, sc: 2, er: 3, ec: 2 },
      { sr: 0, sc: 3, er: 1, ec: 3 },
    ])
    const r2 = step.invert(wb).apply(r.doc!)
    const d = r2.doc!.activeSheet
    expect(d.rowCount).toBe(10)
    expect(d.getCell(4, 0)?.raw).toBe('x')
    expect(d.customRowHeights.get(4)).toBe(40)
    expect(d.customRowHeights.get(5)).toBe(30)
    expect(d.merges).toEqual([
      { sr: 4, sc: 1, er: 5, ec: 1 },
      { sr: 3, sc: 2, er: 5, ec: 2 },
      { sr: 0, sc: 3, er: 1, ec: 3 },
    ])
  })
  it('stepFromJSON 往返：含 sizes/merges 恢复负载的逆操作实例', () => {
    let wb = Workbook.create({ rowCount: 10, colCount: 5 })
    let s = wb.activeSheet
    s = s.setCell(4, 0, { raw: 'x' })
    s = s.setCell(0, 0, { raw: '=A5*2' })
    s = s.setRowHeight(4, 40)
    s = s.setMerges([{ sr: 4, sc: 1, er: 5, ec: 1 }])
    wb = wb.setSheet('s1', s)
    const step = new StructureStep({ sheet: 's1', axis: 'row', index: 4, count: 2, mode: 'delete' }, null)
    const inv = step.invert(wb)
    const back = stepFromJSON(JSON.parse(JSON.stringify(inv.toJSON())))
    expect(back.toJSON()).toEqual(inv.toJSON())
    // 反序列化实例 apply 行为一致：行高/合并区/公式原文全部恢复
    const r = step.apply(wb)
    const d = back.apply(r.doc!).doc!.activeSheet
    expect(d.customRowHeights.get(4)).toBe(40)
    expect(d.merges).toEqual([{ sr: 4, sc: 1, er: 5, ec: 1 }])
    expect(d.getCell(0, 0)?.raw).toBe('=A5*2')
    expect(d.getCell(4, 0)?.raw).toBe('x')
  })
  it('逆操作实例再 invert → 正向实例（redo 对称）', () => {
    const d0 = mk()
    const step = new StructureStep({ sheet: 's1', axis: 'row', index: 4, count: 1, mode: 'delete' }, null)
    const inv = step.invert(d0) as StructureStep
    const r = step.apply(d0)
    const back = inv.invert(r.doc!) as StructureStep
    const r2 = back.apply(r.doc!)
    expect(r2.doc!.activeSheet.getCell(0, 0)?.raw).toBe('=#REF!*2')
  })
  it('越界参数 → failed', () => {
    const d0 = mk()
    expect(new StructureStep({ sheet: 's1', axis: 'row', index: 11, count: 1, mode: 'insert' }, null).apply(d0).ok).toBe(false)
    expect(new StructureStep({ sheet: 's1', axis: 'row', index: 8, count: 5, mode: 'delete' }, null).apply(d0).ok).toBe(false)
    expect(new StructureStep({ sheet: 's1', axis: 'row', index: 0, count: 0, mode: 'insert' }, null).apply(d0).ok).toBe(false)
  })
  it('stepFromJSON 往返', () => {
    const d0 = mk()
    const step = new StructureStep({ sheet: 's1', axis: 'row', index: 4, count: 1, mode: 'delete' }, null)
    const inv = step.invert(d0)
    for (const s of [step, inv]) {
      const back = stepFromJSON(JSON.parse(JSON.stringify(s.toJSON())))
      expect(back.toJSON()).toEqual(s.toJSON())
    }
  })
  it('restore 负载越界 → failed（restore cell out of bounds）', () => {
    const d0 = mk() // 10 行 5 列
    const base = { sizes: [], merges: [], hiddenRows: [], hiddenCols: [], filter: undefined, condFormats: [] }
    const spec = { sheet: 's1', axis: 'row', index: 2, count: 1, mode: 'delete' } as const
    // 逆操作实例（restore 非 null）：执行方向翻转为 insert（10 → 11 行），row 99 仍越界
    const r = new StructureStep(spec, { ...base, cells: [{ sheet: 's1', row: 99, col: 0, cell: { raw: 'x' } }] }).apply(d0)
    expect(r.ok).toBe(false)
    expect(r.failed).toBe('restore cell out of bounds')
    // 列越界（colCount 5，col 5 越界）与未知表同样拦截
    expect(new StructureStep(spec, { ...base, cells: [{ sheet: 's1', row: 0, col: 5, cell: null }] }).apply(d0).ok).toBe(false)
    expect(new StructureStep(spec, { ...base, cells: [{ sheet: 'nope', row: 0, col: 0, cell: null }] }).apply(d0).ok).toBe(false)
  })

  describe('delete 的 undo 恢复 filter', () => {
    const filter: FilterState = {
      range: { sr: 0, sc: 0, er: 9, ec: 3 },
      criteria: {
        1: { type: 'values', excluded: ['x'] },
        3: { type: 'condition', field: 'num', op: 'gt', v1: '5' },
      },
    }
    const mkFilter = () => {
      const wb = Workbook.create({ rowCount: 10, colCount: 5 })
      return wb.setSheet('s1', wb.activeSheet.setFilter(filter))
    }
    it('删数据行（range 被裁剪）→ undo → filter.range 完整恢复', () => {
      const wb = mkFilter()
      const step = new StructureStep({ sheet: 's1', axis: 'row', index: 8, count: 2, mode: 'delete' }, null)
      const r = step.apply(wb)
      expect(r.doc!.activeSheet.filter!.range.er).toBe(7) // 正向：下缘裁剪
      const r2 = step.invert(wb).apply(r.doc!)
      expect(r2.doc!.activeSheet.filter).toEqual(filter)
    })
    it('删表头行（filter 被移除）→ undo → filter 完整恢复（含 criteria）', () => {
      const wb = mkFilter()
      const step = new StructureStep({ sheet: 's1', axis: 'row', index: 0, count: 1, mode: 'delete' }, null)
      const r = step.apply(wb)
      expect(r.doc!.activeSheet.filter).toBeUndefined() // 正向：整体移除
      const r2 = step.invert(wb).apply(r.doc!)
      expect(r2.doc!.activeSheet.filter).toEqual(filter)
    })
    it('删 criteria 所在列（键被重映射）→ undo → 键与值完整恢复', () => {
      const wb = mkFilter()
      const step = new StructureStep({ sheet: 's1', axis: 'col', index: 1, count: 1, mode: 'delete' }, null)
      const r = step.apply(wb)
      expect(Object.keys(r.doc!.activeSheet.filter!.criteria)).toEqual(['2']) // 正向：1 被删，3 → 2
      const r2 = step.invert(wb).apply(r.doc!)
      expect(r2.doc!.activeSheet.filter).toEqual(filter)
    })
  })
})
