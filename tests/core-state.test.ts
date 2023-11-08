import { describe, it, expect } from 'vitest'
import { Workbook } from '../src/core/model'
import { SheetState } from '../src/core/state'
import { Plugin, PluginKey } from '../src/core/plugin'
import { singleCell } from '../src/core/selection'

const mkState = (plugins: Plugin[] = []) =>
  SheetState.create({ doc: Workbook.create({ rowCount: 10, colCount: 5 }), plugins })

describe('state/transaction', () => {
  it('tr builder 多步 doc 链', () => {
    const s = mkState()
    const tr = s.tr.setCell(0, 0, 'a').setCell(1, 1, 'b').patchStyle({ sr: 0, sc: 0, er: 0, ec: 0 }, { bold: true })
    expect(tr.steps.length).toBe(3)
    expect(tr.doc.sheet('s1').getCell(1, 1)).toEqual({ raw: 'b' })
    expect(tr.doc.sheet('s1').getCell(0, 0)!.style).toEqual({ bold: true })
    const s2 = s.apply(tr)
    expect(s2.doc.sheet('s1').getCell(0, 0)!.style).toEqual({ bold: true })
    expect(s.doc.sheet('s1').getCell(0, 0)).toBeUndefined() // 旧 state 不变
  })
  it('tr.setCell raw==="" → 删除；setSelection/scrollIntoView/meta', () => {
    const s = mkState().apply(mkState().tr.setCell(0, 0, 'x'))
    const tr = s.tr.setCell(0, 0, '').setSelection(singleCell(2, 2)).scrollIntoView().setMeta('k', 42)
    expect(tr.scrolledIntoView).toBe(true)
    expect(tr.getMeta('k')).toBe(42)
    const s2 = s.apply(tr)
    expect(s2.doc.sheet('s1').getCell(0, 0)).toBeUndefined()
    expect(s2.selection).toEqual({ anchor: { row: 2, col: 2 }, focus: { row: 2, col: 2 } })
  })
  it('插件 state field init/apply 次序与 getField', () => {
    const calls: string[] = []
    const key = new PluginKey('counter')
    const plugin = new Plugin({
      key,
      state: {
        init: () => { calls.push('init'); return 0 },
        apply: (tr, v) => { calls.push('apply'); return v + tr.steps.length },
      },
    })
    const s = mkState([plugin])
    expect(calls).toEqual(['init'])
    expect(s.getField(key)).toBe(0)
    const s2 = s.apply(s.tr.setCell(0, 0, 'a').setCell(1, 1, 'b'))
    expect(s2.getField(key)).toBe(2)
  })
  it('appendTransaction 展开且防死循环', () => {
    let fired = 0
    const p = new Plugin({
      appendTransaction: (trs, oldS, newS) => {
        fired++
        // 只在 A1 非空且 B1 空时补一格，自然终止
        if (newS.doc.sheet('s1').getCell(0, 0) && !newS.doc.sheet('s1').getCell(0, 1)) {
          return newS.tr.setCell(0, 1, 'auto')
        }
        return null
      },
    })
    const s = mkState([p])
    const { state, trs } = s.applyTransaction(s.tr.setCell(0, 0, 'go'))
    expect(state.doc.sheet('s1').getCell(0, 1)).toEqual({ raw: 'auto' })
    expect(trs.length).toBe(2)
    expect(fired).toBeGreaterThanOrEqual(2)
  })
})
