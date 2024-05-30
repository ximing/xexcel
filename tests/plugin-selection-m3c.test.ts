// selection 插件 Ctrl/Shift 交互 + keymap Shift+Arrow 单测。
// 用 SheetState.create + 最小 fake view（state getter + dispatch 捕获 + focus/geometry no-op），
// 直接调 plugin.spec.props.handleMouseDown / handleKeyDown，断言 state.selection 形态。
import { describe, expect, it } from 'vitest'
import { Workbook } from '../src/core/model'
import { SheetState } from '../src/core/state'
import { appendRange, rangeSelection, singleCell } from '../src/core/selection'
import type { Transaction } from '../src/core/transaction'
import { selection } from '../src/plugins/selection'
import { keymap } from '../src/plugins/keymap'
import type { EditorView } from '../src/view/editorview'
import type { HitResult } from '../src/core/plugin'

// 最小 fake view：selection 插件 mousedown 只用 state/dispatch/focus；
// keymap move 用 geometry() 判隐藏行列，mock 成全可见。
interface FakeViewOpts { selection?: ReturnType<typeof singleCell> }
const mkView = (opts: FakeViewOpts = {}) => {
  let state = SheetState.create({
    doc: Workbook.create({ rowCount: 10, colCount: 10 }),
    selection: opts.selection ?? singleCell(0, 0),
    plugins: [selection(), keymap()],
  })
  const view = {
    get state() { return state },
    dispatch: (tr: Transaction) => { state = state.applyTransaction(tr).state },
    focus: () => {},
    geometry: () => ({ rowHeight: () => 24, colWidth: () => 96, frozenRows: 0, frozenCols: 0, frozenWidth: 0, frozenHeight: 0 }),
  } as unknown as EditorView
  return { view, getState: () => state }
}

const cellHit = (row: number, col: number): HitResult => ({ region: 'cell', row, col })
const mkEvent = (mods: { ctrl?: boolean; shift?: boolean } = {}): MouseEvent =>
  ({ ctrlKey: !!mods.ctrl, shiftKey: !!mods.shift, metaKey: !!mods.ctrl, button: 0, clientX: 0, clientY: 0 } as unknown as MouseEvent)

describe('selection 插件 Ctrl/Shift 交互', () => {
  // Ctrl+click：格不在选区 → 追加单格区域
  it('Ctrl+click 追加单格区域', () => {
    const { view, getState } = mkView({ selection: singleCell(0, 0) })
    const handled = selection().spec.props!.handleMouseDown!(view, mkEvent({ ctrl: true }), cellHit(2, 2))
    expect(handled).toBe(true)
    const sel = getState().selection
    expect(sel.ranges.length).toBe(2)
    expect(sel.ranges[1]).toEqual({ sr: 2, sc: 2, er: 2, ec: 2 })
    expect(sel.activeCell).toEqual({ row: 2, col: 2 })
  })

  // Ctrl+click 已选格 → 反选 LIFO 移除最后加入者
  it('Ctrl+click 已选格 → 反选 LIFO 移除最后加入者', () => {
    // 两重叠 range 都含 (0,0)
    const sel0 = appendRange(appendRange(singleCell(0, 0), { sr: 0, sc: 0, er: 1, ec: 1 }), { sr: 0, sc: 0, er: 2, ec: 2 })
    const { view, getState } = mkView({ selection: sel0 })
    expect(getState().selection.ranges.length).toBe(3)
    selection().spec.props!.handleMouseDown!(view, mkEvent({ ctrl: true }), cellHit(0, 0))
    const sel = getState().selection
    expect(sel.ranges.length).toBe(2) // 移除最后加入者 ranges[2]
    expect(sel.ranges[1]).toEqual({ sr: 0, sc: 0, er: 1, ec: 1 })
  })

  // Shift+click：扩展活动区域到该格
  it('Shift+click 扩展活动区域到该格', () => {
    const { view, getState } = mkView({ selection: singleCell(0, 0) })
    selection().spec.props!.handleMouseDown!(view, mkEvent({ shift: true }), cellHit(3, 4))
    const sel = getState().selection
    expect(sel.ranges).toEqual([{ sr: 0, sc: 0, er: 3, ec: 4 }])
    expect(sel.activeCell).toEqual({ row: 3, col: 4 })
  })

  // 无修饰 click：singleCell 重置
  it('无修饰 click 重置为单格选区', () => {
    const { view, getState } = mkView({ selection: rangeSelection({ sr: 0, sc: 0, er: 5, ec: 5 }) })
    selection().spec.props!.handleMouseDown!(view, mkEvent(), cellHit(2, 2))
    const sel = getState().selection
    expect(sel.ranges).toEqual([{ sr: 2, sc: 2, er: 2, ec: 2 }])
    expect(sel.activeCell).toEqual({ row: 2, col: 2 })
  })
})

describe('keymap Shift+Arrow', () => {
  it('Shift+Right 扩展活动区域右扩一列', () => {
    const { view, getState } = mkView({ selection: singleCell(0, 0) })
    const handled = keymap().spec.props!.handleKeyDown!(view, { key: 'ArrowRight', shiftKey: true, ctrlKey: false, metaKey: false, altKey: false, preventDefault: () => {} } as KeyboardEvent)
    expect(handled).toBe(true)
    const sel = getState().selection
    expect(sel.ranges).toEqual([{ sr: 0, sc: 0, er: 0, ec: 1 }])
    expect(sel.activeCell).toEqual({ row: 0, col: 1 })
  })

  it('ArrowRight 无 Shift 塌缩为单格', () => {
    const { view, getState } = mkView({ selection: rangeSelection({ sr: 0, sc: 0, er: 3, ec: 4 }) })
    keymap().spec.props!.handleKeyDown!(view, { key: 'ArrowRight', shiftKey: false, ctrlKey: false, metaKey: false, altKey: false, preventDefault: () => {} } as KeyboardEvent)
    const sel = getState().selection
    // navigateFocus 从 activeCell(0,0) 右移 → (0,1)；无 shift → singleCell
    expect(sel.ranges).toEqual([{ sr: 0, sc: 1, er: 0, ec: 1 }])
    expect(sel.activeCell).toEqual({ row: 0, col: 1 })
  })
})
