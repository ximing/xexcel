import { describe, expect, it } from 'vitest'
import { SheetState } from '../src/core/state'
import { CellStyle, Workbook } from '../src/core/model'
import { Selection, singleCell } from '../src/core/selection'
import { Transaction } from '../src/core/transaction'
import { FormatPainterState, contextMenuKey, formatPainterKey } from '../src/view/types'
import type { EditorView } from '../src/view/editorview'
import { metaField, painter } from '../src/plugins/uistate'

// metaField 经 tr meta 透传：激活/锁定/解除（显式注册 metaField，getField 默认 null）
const mkState = () =>
  SheetState.create({
    doc: Workbook.create({ rowCount: 10, colCount: 10 }),
    plugins: [metaField(formatPainterKey, null)],
  })

describe('formatPainter state field', () => {
  it('默认 null；setMeta 设置；不带 meta 的事务保持', () => {
    let s = mkState()
    expect(s.getField(formatPainterKey)).toBeNull()
    const fp = { style: { bold: true }, locked: false }
    s = s.applyTransaction(s.tr.setMeta(formatPainterKey, fp).setMeta('addToHistory', false)).state
    expect(s.getField(formatPainterKey)).toEqual(fp)
    s = s.applyTransaction(s.tr.setCell(0, 0, 'x')).state
    expect(s.getField(formatPainterKey)).toEqual(fp)
  })
})

describe('格式刷应用', () => {
  it('setCellStyles 整体替换目标 style（含 border/numFmt），非合并', () => {
    let s = mkState()
    s = s.applyTransaction(
      s.tr.setCell(0, 0, 'a', { bold: true, numFmt: '0%' }).setCell(1, 1, 'b', { italic: true, bg: '#fff' }),
    ).state
    const src = s.activeSheet.getCell(0, 0)!.style!
    s = s.applyTransaction(s.tr.setCellStyles([{ row: 1, col: 1, style: { ...src } }])).state
    expect(s.activeSheet.getCell(1, 1)!.style).toEqual({ bold: true, numFmt: '0%' })
    // undo 恢复
    s = s.applyTransaction(s.tr.setCellStyles([{ row: 1, col: 1, style: { italic: true, bg: '#fff' } }])).state
    expect(s.activeSheet.getCell(1, 1)!.style).toEqual({ italic: true, bg: '#fff' })
  })
})

// painter 插件本体：最小 fake view（state + dispatch 捕获 tr），直接调 props.handleMouseUp
describe('painter 插件 handleMouseUp', () => {
  const mkView = (sel: Selection, fp: FormatPainterState | null, menuOpen = false) => {
    let state = SheetState.create({
      doc: Workbook.create({ rowCount: 10, colCount: 10 }),
      selection: sel,
      plugins: [metaField(formatPainterKey, null), metaField(contextMenuKey, null), painter()],
    })
    if (fp) {
      state = state.applyTransaction(
        state.tr.setMeta(formatPainterKey, fp).setMeta('addToHistory', false),
      ).state
    }
    if (menuOpen) {
      state = state.applyTransaction(
        state.tr.setMeta(contextMenuKey, { x: 0, y: 0 }).setMeta('addToHistory', false),
      ).state
    }
    const trs: Transaction[] = []
    const view = {
      get state() {
        return state
      },
      dispatch: (tr: Transaction) => {
        trs.push(tr)
        state = state.applyTransaction(tr).state
      },
    } as unknown as EditorView
    return { view, trs, getState: () => state }
  }
  const handleMouseUp = painter().spec.props!.handleMouseUp!
  const evt = { button: 0 } as MouseEvent
  const stylesIn = (s: SheetState, sr: number, er: number, sc: number, ec: number): (CellStyle | undefined)[] => {
    const out: (CellStyle | undefined)[] = []
    for (let r = sr; r <= er; r++) for (let c = sc; c <= ec; c++) out.push(s.activeSheet.getCell(r, c)?.style)
    return out
  }

  it('非锁定：刷选区各格整体 style，meta 置 null 自动解除，且事务可入 history', () => {
    const fp: FormatPainterState = { style: { bold: true, numFmt: '0%' }, locked: false }
    const sel: Selection = { anchor: { row: 1, col: 1 }, focus: { row: 2, col: 2 } }
    const { view, trs, getState } = mkView(sel, fp)
    const handled = handleMouseUp(view, evt, { region: 'cell', row: 2, col: 2 })
    expect(handled).toBe(true)
    expect(trs).toHaveLength(1)
    const tr = trs[0]
    expect(tr.getMeta(formatPainterKey)).toBeNull() // 非锁定刷一次解除
    expect(tr.getMeta('addToHistory')).toBeUndefined() // 样式变更必须可 undo
    const applied = getState()
    expect(stylesIn(applied, 1, 2, 1, 2)).toEqual([fp.style, fp.style, fp.style, fp.style])
    expect(applied.activeSheet.getCell(0, 0)?.style).toBeUndefined() // 选区外不受影响
    expect(applied.getField(formatPainterKey)).toBeNull()
  })

  it('锁定：meta 保持 fp，可连刷', () => {
    const fp: FormatPainterState = { style: { italic: true }, locked: true }
    const { view, trs, getState } = mkView(singleCell(0, 0), fp)
    expect(handleMouseUp(view, evt, { region: 'cell', row: 0, col: 0 })).toBe(true)
    expect(trs[0].getMeta(formatPainterKey)).toEqual(fp)
    expect(getState().getField(formatPainterKey)).toEqual(fp)
    expect(getState().activeSheet.getCell(0, 0)?.style).toEqual(fp.style)
  })

  it('region 非 cell：不 dispatch，返回 false', () => {
    const fp: FormatPainterState = { style: { bold: true }, locked: false }
    const { view, trs } = mkView(singleCell(0, 0), fp)
    expect(handleMouseUp(view, evt, { region: 'rowheader', row: 0, col: 0 })).toBe(false)
    expect(trs).toHaveLength(0)
  })

  it('未激活：不 dispatch，返回 false', () => {
    const { view, trs } = mkView(singleCell(0, 0), null)
    expect(handleMouseUp(view, evt, { region: 'cell', row: 0, col: 0 })).toBe(false)
    expect(trs).toHaveLength(0)
  })

  it('右键 mouseup（button=2）：不 dispatch，返回 false', () => {
    const fp: FormatPainterState = { style: { bold: true }, locked: false }
    const { view, trs, getState } = mkView(singleCell(0, 0), fp)
    expect(handleMouseUp(view, { button: 2 } as MouseEvent, { region: 'cell', row: 0, col: 0 })).toBe(false)
    expect(trs).toHaveLength(0)
    expect(getState().getField(formatPainterKey)).toEqual(fp) // 右键不消耗格式刷
  })

  it('右键菜单打开中：左键 mouseup 不 dispatch，返回 false', () => {
    const fp: FormatPainterState = { style: { bold: true }, locked: false }
    const { view, trs, getState } = mkView(singleCell(0, 0), fp, true)
    expect(handleMouseUp(view, evt, { region: 'cell', row: 0, col: 0 })).toBe(false)
    expect(trs).toHaveLength(0)
    expect(getState().getField(formatPainterKey)).toEqual(fp)
  })
})
