// selection 插件 Ctrl/Shift 交互 + keymap Shift+Arrow 单测。
// 用 SheetState.create + 最小 fake view（state getter + dispatch 捕获 + focus/geometry no-op），
// 直接调 plugin.spec.props.handleMouseDown / handleKeyDown，断言 state.selection 形态。
import { describe, expect, it } from 'vitest'
import { Workbook } from '@xexcel/core'
import { SheetState } from '@xexcel/core'
import { appendRange, rangeSelection, singleCell } from '@xexcel/core'
import type { Transaction } from '@xexcel/core'
import { selection } from '../src/plugins/selection'
import { keymap } from '../src/plugins/keymap'
import type { EditorView } from '../src/view/editorview'
import type { HitResult } from '@xexcel/core'

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
    // 拖拽测试需要：dom/stage 供 edgeDelta 判定自动滚动（大视口不触发），pointerToCell 供 addrAt 映射
    dom: { getBoundingClientRect: () => ({ left: 0, top: 0, right: 1000, bottom: 1000, width: 1000, height: 1000 }) },
    stage: { width: () => 1000, height: () => 1000 },
    pointerToCell: (x: number, y: number) => ({ row: Math.max(0, Math.min(9, Math.floor((y - 30) / 50))), col: Math.max(0, Math.min(9, Math.floor((x - 30) / 50))) }),
  } as unknown as EditorView
  return { view, getState: () => state }
}

const cellHit = (row: number, col: number): HitResult => ({ region: 'cell', row, col })
const mkEvent = (mods: { ctrl?: boolean; shift?: boolean } = {}): MouseEvent =>
  ({ ctrlKey: !!mods.ctrl, shiftKey: !!mods.shift, metaKey: !!mods.ctrl, button: 0, clientX: 0, clientY: 0 } as unknown as MouseEvent)
const keyDown = (key: string, shift = false): KeyboardEvent =>
  ({ key, shiftKey: shift, ctrlKey: false, metaKey: false, altKey: false, preventDefault: () => {} } as unknown as KeyboardEvent)
const move = (x: number, y: number): MouseEvent =>
  ({ clientX: x, clientY: y, button: 0 } as unknown as MouseEvent)

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

// 审查员 repro：旧 extendActiveRange 以 sel.activeCell 作锚点 → 连续 Shift 扩展选区滑动丢起始格。
// 新实现以活动区域中 activeCell 的对角格为固定锚点 → 生长。下列三例覆盖 keymap / Shift+click / drag 多步路径。
describe('Shift 扩展多步生长（锚点固定，不滑动）', () => {
  it('keymap 连续两次 Shift+Right 生长', () => {
    const { view, getState } = mkView({ selection: singleCell(0, 0) })
    const km = keymap()
    km.spec.props!.handleKeyDown!(view, keyDown('ArrowRight', true))
    expect(getState().selection.ranges).toEqual([{ sr: 0, sc: 0, er: 0, ec: 1 }])
    km.spec.props!.handleKeyDown!(view, keyDown('ArrowRight', true))
    expect(getState().selection.ranges).toEqual([{ sr: 0, sc: 0, er: 0, ec: 2 }]) // 生长，非滑动到 {0,1..0,2}
    expect(getState().selection.activeCell).toEqual({ row: 0, col: 2 })
  })

  it('Shift+click 后 Shift+Arrow 继续生长（锚点固定在起始格）', () => {
    const { view, getState } = mkView({ selection: singleCell(0, 0) })
    const sel = selection()
    sel.spec.props!.handleMouseDown!(view, mkEvent({ shift: true }), cellHit(0, 3)) // Shift+click (0,3)
    expect(getState().selection.ranges).toEqual([{ sr: 0, sc: 0, er: 0, ec: 3 }])
    keymap().spec.props!.handleKeyDown!(view, keyDown('ArrowRight', true)) // Shift+Right
    expect(getState().selection.ranges).toEqual([{ sr: 0, sc: 0, er: 0, ec: 4 }]) // 从 (0,0) 继续生长
    expect(getState().selection.activeCell).toEqual({ row: 0, col: 4 })
  })

  it('drag 连续 mousemove 生长（锚点固定在 mousedown 格）', () => {
    const { view, getState } = mkView({ selection: singleCell(0, 0) })
    const sel = selection()
    sel.spec.props!.handleMouseDown!(view, mkEvent(), cellHit(0, 0)) // mousedown (0,0) → singleCell + 进入拖拽态
    sel.spec.props!.handleMouseMove!(view, move(130, 30), cellHit(0, 2)) // pointerToCell → (0,2)
    expect(getState().selection.ranges).toEqual([{ sr: 0, sc: 0, er: 0, ec: 2 }])
    sel.spec.props!.handleMouseMove!(view, move(230, 30), cellHit(0, 4)) // → (0,4)，应从 (0,0) 生长，非滑动到 {0,2..0,4}
    expect(getState().selection.ranges).toEqual([{ sr: 0, sc: 0, er: 0, ec: 4 }])
    expect(getState().selection.activeCell).toEqual({ row: 0, col: 4 })
  })

  // issue #1：右键 mousedown 不得进入拖拽态（右键菜单会拦截 mouseup 冒泡，drag 若残留则 mousemove 持续改选区）
  it('右键 mousedown 不进入拖拽态，后续 mousemove 不改选区', () => {
    const { view, getState } = mkView({ selection: rangeSelection({ sr: 1, sc: 1, er: 3, ec: 3 }) })
    const sel = selection()
    const rightDown = { ...mkEvent(), button: 2 } as MouseEvent
    const handled = sel.spec.props!.handleMouseDown!(view, rightDown, cellHit(2, 2))
    expect(handled).toBe(false)
    expect(getState().selection.ranges).toEqual([{ sr: 1, sc: 1, er: 3, ec: 3 }]) // 选区不被右键改动
    sel.spec.props!.handleMouseMove!(view, move(330, 180), cellHit(5, 5)) // 菜单打开期间移动鼠标
    expect(getState().selection.ranges).toEqual([{ sr: 1, sc: 1, er: 3, ec: 3 }]) // 选区不跟随
  })

  it('右键点行列边框不进入 resize 拖拽态', () => {
    const { view, getState } = mkView()
    const sel = selection()
    const rightDown = { ...mkEvent(), button: 2 } as MouseEvent
    const before = getState().activeSheet.colWidth(2)
    const handled = sel.spec.props!.handleMouseDown!(view, rightDown, { region: 'colborder', row: -1, col: 2 })
    expect(handled).toBe(false)
    expect(getState().activeSheet.colWidth(2)).toBe(before) // 未拖拽，列宽不变
  })
})
