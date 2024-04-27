import { describe, expect, it } from 'vitest'
import { parseA1, parseRangeA1 } from '../src/core/addr'
import { CondFormatRule, SheetData } from '../src/core/model'
import { SetCondFormatsStep } from '../src/core/steps'
import { Workbook } from '../src/core/model'

const rule = (over: Partial<CondFormatRule> = {}): CondFormatRule => ({
  id: 'cf1',
  range: { sr: 0, sc: 0, er: 3, ec: 1 },
  type: 'value',
  op: 'gt',
  v1: '10',
  style: { bg: '#ffc7ce' },
  ...over,
} as CondFormatRule)

describe('parseA1/parseRangeA1', () => {
  it('单格与区域', () => {
    expect(parseA1('B3')).toEqual({ row: 2, col: 1 })
    expect(parseA1('AA10')).toEqual({ row: 9, col: 26 })
    expect(parseRangeA1('B2:C3')).toEqual({ sr: 1, sc: 1, er: 2, ec: 2 })
    expect(parseRangeA1('C3:B2')).toEqual({ sr: 1, sc: 1, er: 2, ec: 2 }) // normalize
    expect(parseA1('1A')).toBeNull()
    expect(parseRangeA1('A1:B2:C3')).toBeNull()
  })
})

describe('condFormats 模型', () => {
  it('setCondFormats + JSON 往返', () => {
    const s = SheetData.create({ rowCount: 10, colCount: 10 }).setCondFormats([rule()])
    expect(s.condFormats).toHaveLength(1)
    const back = SheetData.fromJSON(JSON.parse(JSON.stringify(s.toJSON())))
    expect(back.condFormats).toEqual(s.condFormats)
  })
  it('空表缺省 []', () => {
    expect(SheetData.create({ rowCount: 2, colCount: 2 }).condFormats).toEqual([])
  })
  it('结构 remap：删行平移 rule.range；表头行删空丢规则', () => {
    const s = SheetData.create({ rowCount: 10, colCount: 10 }).setCondFormats([rule()])
    const del = s.deleteRows(1, 2)
    expect(del.condFormats[0].range).toEqual({ sr: 0, sc: 0, er: 1, ec: 1 })
    const gone = s.deleteRows(0, 4)
    expect(gone.condFormats).toEqual([])
  })
})

describe('SetCondFormatsStep', () => {
  it('apply/invert/toJSON 往返', () => {
    const wb = Workbook.create({ rowCount: 10, colCount: 10 })
    const step = new SetCondFormatsStep('s1', [rule()])
    const r = step.apply(wb)
    expect(r.ok).toBe(true)
    expect(r.doc!.sheet('s1').condFormats).toHaveLength(1)
    const inv = step.invert(wb)
    const r2 = inv.apply(r.doc!)
    expect(r2.doc!.sheet('s1').condFormats).toEqual([])
    expect(JSON.parse(JSON.stringify(step.toJSON()))).toEqual(step.toJSON())
  })
})

describe('结构 undo', () => {
  it('删行 undo 恢复 condFormats', async () => {
    const { SheetState } = await import('../src/core/state')
    const { history, undo, undoDepth } = await import('../src/core/history')
    let s = SheetState.create({ doc: Workbook.create({ rowCount: 10, colCount: 10 }), plugins: [history()] })
    s = s.applyTransaction(s.tr.setCondFormats([rule()])).state
    s = s.applyTransaction(s.tr.structure('row', 0, 4, 'delete')).state
    expect(s.activeSheet.condFormats).toEqual([])
    expect(undoDepth(s)).toBeGreaterThan(0)
    undo(s, (tr) => { s = s.applyTransaction(tr).state })
    expect(s.activeSheet.condFormats).toEqual([rule()])
  })
})
