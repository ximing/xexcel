import { describe, expect, it } from 'vitest'
import { SheetData, Workbook } from '../src/core/model'
import { SetValidationsStep, stepFromJSON } from '../src/core/steps'

const numRule = { id: 'v1', range: { sr: 0, sc: 0, er: 9, ec: 0 }, type: 'numRange' as const, op: 'between' as const, v1: '1', v2: '9' }
const listRule = { id: 'v2', range: { sr: 0, sc: 1, er: 5, ec: 1 }, type: 'list' as const, items: ['a', 'b'] }

describe('SheetData.validations 模型面', () => {
  it('setValidations + toJSON/fromJSON 往返；旧 JSON 无字段默认 []', () => {
    const s = SheetData.create({ rowCount: 100, colCount: 26 }).setValidations([numRule, listRule])
    const back = SheetData.fromJSON(s.toJSON())
    expect(back.validations).toEqual([numRule, listRule])
    const legacy = SheetData.fromJSON({ rowCount: 100, colCount: 26 })
    expect(legacy.validations).toEqual([])
  })

  it('结构插删 remap：范围内平移，起点被删丢规则', () => {
    const s = SheetData.create({ rowCount: 100, colCount: 26 }).setValidations([numRule])
    // 在第 0 行前插 2 行 → range 下移
    expect(s.insertRows(0, 2).validations[0].range).toEqual({ sr: 2, sc: 0, er: 11, ec: 0 })
    // 删除 range 起点所在行 → 丢规则
    expect(s.deleteRows(0, 1).validations).toEqual([])
    // 删除 range 之后的行 → 不动
    expect(s.deleteRows(50, 1).validations[0].range).toEqual(numRule.range)
  })
})

describe('SetValidationsStep', () => {
  it('apply/invert/toJSON 往返', () => {
    let wb = Workbook.create({ rowCount: 100, colCount: 26 })
    const step = new SetValidationsStep('s1', [numRule])
    const r = step.apply(wb)
    expect(r.ok).toBe(true)
    wb = r.doc!
    expect(wb.sheet('s1').validations).toEqual([numRule])
    const inv = step.invert(Workbook.create({ rowCount: 100, colCount: 26 }))
    expect(inv.apply(wb).doc!.sheet('s1').validations).toEqual([])
    expect(stepFromJSON(JSON.parse(JSON.stringify(step.toJSON()))).toJSON()).toEqual(step.toJSON())
  })
})

describe('StructureStep delete-undo 恢复 validations', () => {
  it('删行 undo 后规则原文恢复', async () => {
    const { SheetState } = await import('../src/core/state')
    const { history, undo, undoDepth } = await import('../src/core/history')
    const rule = { ...numRule, range: { sr: 20, sc: 0, er: 29, ec: 0 } }
    let wb = Workbook.create({ rowCount: 100, colCount: 26 })
    wb = wb.setSheet('s1', wb.sheet('s1').setValidations([rule]))
    let s = SheetState.create({ doc: wb, plugins: [history()] })
    s = s.applyTransaction(s.tr.structure('row', 10, 5, 'delete')).state
    // range 起点 20 ≥ 删除区(10..14) → 平移到 15
    expect(s.activeSheet.validations[0].range.sr).toBe(15)
    expect(undoDepth(s)).toBeGreaterThan(0)
    undo(s, (tr) => { s = s.applyTransaction(tr).state })
    // invert 应用后 validations 与原文完全相等
    expect(s.activeSheet.validations).toEqual([rule])
  })
})
