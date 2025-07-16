import { describe, it, expect } from 'vitest'
import { Workbook } from '../src/core/model'
import { SheetState } from '../src/core/state'
import { history, undo, redo, undoDepth, redoDepth } from '../src/core/history'
import { singleCell } from '../src/core/selection'

const mk = () => SheetState.create({ doc: Workbook.create({ rowCount: 10, colCount: 5 }), plugins: [history()] })
const apply = (s: SheetState, build: (s: SheetState) => any) => s.applyTransaction(build(s)).state

describe('history', () => {
  it('undo/redo 单组往返（含 selection 恢复）', () => {
    let s = mk()
    s = apply(s, st => st.tr.setSelection(singleCell(3, 3)))
    s = apply(s, st => st.tr.setCell(0, 0, 'a').setCell(1, 1, 'b').setSelection(singleCell(0, 0)))
    expect(undoDepth(s)).toBe(1) // 纯 selection 事务不入栈
    expect(undo(s, tr => { s = s.applyTransaction(tr).state })).toBe(true)
    expect(s.doc.sheet('s1').getCell(0, 0)).toBeUndefined()
    expect(s.selection).toEqual({ ranges: [{ sr: 3, sc: 3, er: 3, ec: 3 }], activeCell: { row: 3, col: 3 } })
    expect(redoDepth(s)).toBe(1)
    expect(redo(s, tr => { s = s.applyTransaction(tr).state })).toBe(true)
    expect(s.doc.sheet('s1').getCell(0, 0)).toEqual({ raw: 'a' })
    expect(s.doc.sheet('s1').getCell(1, 1)).toEqual({ raw: 'b' })
  })
  it('多组依次撤销；新事务清空 redo 栈', () => {
    let s = mk()
    s = apply(s, st => st.tr.setCell(0, 0, 'a'))
    s = apply(s, st => st.tr.setCell(0, 1, 'b'))
    expect(undoDepth(s)).toBe(2)
    const doUndo = () => undo(s, tr => { s = s.applyTransaction(tr).state })
    const doRedo = () => redo(s, tr => { s = s.applyTransaction(tr).state })
    expect(doUndo()).toBe(true)
    expect(doUndo()).toBe(true)
    expect(s.doc.sheet('s1').getCell(0, 0)).toBeUndefined()
    expect(s.doc.sheet('s1').getCell(0, 1)).toBeUndefined()
    expect(doRedo()).toBe(true)
    expect(s.doc.sheet('s1').getCell(0, 0)).toEqual({ raw: 'a' })
    expect(s.doc.sheet('s1').getCell(0, 1)).toBeUndefined()
    expect(doUndo()).toBe(true)
    s = apply(s, st => st.tr.setCell(2, 2, 'c'))
    expect(redoDepth(s)).toBe(0)
  })
  it('addToHistory:false 不入栈', () => {
    let s = mk()
    s = apply(s, st => st.tr.setCell(0, 0, 'temp').setMeta('addToHistory', false))
    expect(undoDepth(s)).toBe(0)
  })
  it('空栈 undo 返回 false', () => {
    expect(undo(mk())).toBe(false)
  })
  it('done 上限 200 组', () => {
    let s = mk()
    for (let i = 0; i < 210; i++) s = apply(s, st => st.tr.setCell(0, 0, String(i)))
    expect(undoDepth(s)).toBe(200)
  })
})
