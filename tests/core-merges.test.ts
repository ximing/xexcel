import { describe, it, expect } from 'vitest'
import { Workbook } from '../src/core/model'
import { SheetState } from '../src/core/state'
import { mergeSelection, unmergeSelection } from '../src/core/commands'
import { SetMergesStep, stepFromJSON } from '../src/core/steps'
import type { Transaction } from '../src/core/transaction'

const mk = (cells: Array<[number, number, string]>) => {
  let wb = Workbook.create({ rowCount: 10, colCount: 10 })
  let sheet = wb.activeSheet
  for (const [r, c, raw] of cells) sheet = sheet.setCell(r, c, { raw })
  return wb.setSheet('s1', sheet)
}
const stateOf = (doc: Workbook, sel = { anchor: { row: 1, col: 1 }, focus: { row: 2, col: 2 } }) =>
  SheetState.create({ doc, selection: sel })
const run = (state: SheetState, cmd: (s: SheetState, d?: (tr: Transaction) => void) => boolean) => {
  let tr: Transaction | null = null
  cmd(state, (t) => (tr = t))
  return tr! ? state.apply(tr!) : state
}

describe('merges', () => {
  it('mergeSelection：写入 merge 并清非锚点格', () => {
    const st = run(stateOf(mk([[1, 1, 'a'], [1, 2, 'b'], [2, 2, 'c']])), mergeSelection)
    expect(st.activeSheet.merges).toEqual([{ sr: 1, sc: 1, er: 2, ec: 2 }])
    expect(st.activeSheet.getCell(1, 1)?.raw).toBe('a')
    expect(st.activeSheet.getCell(1, 2)).toBeUndefined()
    expect(st.activeSheet.getCell(2, 2)).toBeUndefined()
  })
  it('mergeSelection：与旧 merge 相交 → 旧移除新并入', () => {
    let st = run(stateOf(mk([])), mergeSelection) // merge B2:C3
    st = run(
      SheetState.create({ doc: st.doc, selection: { anchor: { row: 2, col: 2 }, focus: { row: 3, col: 3 } } }),
      mergeSelection,
    ) // merge C3:D4（与 B2:C3 交于 C3）
    expect(st.activeSheet.merges).toEqual([{ sr: 2, sc: 2, er: 3, ec: 3 }])
  })
  it('unmergeSelection：移除相交 merge，不恢复被清的值', () => {
    let st = run(stateOf(mk([[1, 1, 'a'], [1, 2, 'b']])), mergeSelection)
    st = run(stateOf(st.doc), unmergeSelection)
    expect(st.activeSheet.merges).toEqual([])
    expect(st.activeSheet.getCell(1, 2)).toBeUndefined() // 值已在合并时清除，拆分不恢复
  })
  it('undo 合并 → 值与 merges 恢复', () => {
    const before = stateOf(mk([[1, 1, 'a'], [1, 2, 'b']]))
    const st = run(before, mergeSelection)
    const trs: Transaction[] = []
    mergeSelection(before, (t) => trs.push(t))
    // 手动构造 undo：invert 每个 step
    let doc = st.doc
    for (let i = trs[0].steps.length - 1; i >= 0; i--) {
      const beforeI = i === 0 ? before.doc : trs[0].docs[i - 1]
      const inv = trs[0].steps[i].invert(beforeI)
      const r = inv.apply(doc)
      doc = r.doc!
    }
    expect(doc.activeSheet.merges).toEqual([])
    expect(doc.activeSheet.getCell(1, 2)?.raw).toBe('b')
  })
  it('mergeAt：命中合并区任意位置返回区域，未命中 null', () => {
    const st = run(stateOf(mk([])), mergeSelection)
    expect(st.activeSheet.mergeAt(2, 2)).toEqual({ sr: 1, sc: 1, er: 2, ec: 2 })
    expect(st.activeSheet.mergeAt(1, 1)).toEqual({ sr: 1, sc: 1, er: 2, ec: 2 })
    expect(st.activeSheet.mergeAt(0, 0)).toBeNull()
  })
  it('SetMergesStep apply/invert/toJSON 往返', () => {
    const d0 = mk([])
    const step = new SetMergesStep('s1', [{ sr: 0, sc: 0, er: 1, ec: 1 }])
    const r = step.apply(d0)
    expect(r.doc!.activeSheet.merges).toHaveLength(1)
    const r2 = step.invert(d0).apply(r.doc!)
    expect(r2.doc!.activeSheet.merges).toEqual([])
    const back = stepFromJSON(JSON.parse(JSON.stringify(step.toJSON())))
    expect(back.toJSON()).toEqual(step.toJSON())
  })
})
