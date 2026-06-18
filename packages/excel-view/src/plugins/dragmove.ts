// F3 边框拖动移动：selborder mousedown 进入拖拽态；mousemove 画目标虚线框（dragPreviewKey）；
// mouseup → cut 移动（SetCellsStep 写目标 raw+style + clearRange 清源，一个 Transaction 一次 undo）；
// 源与目标相交 / 目标落 merge → 拒绝 + pluginNotice 提示（plugins 层不经 window.alert）；
// 多区域仅移动活动区域（selectionRange(sel) = ranges[last]）。
import { CellRange, clampRange, normalizeRange, rangesIntersect } from '@gmi/excel-core'
import { Cell, SheetData } from '@gmi/excel-core'
import { EditorViewLike, HitResult, Plugin } from '@gmi/excel-core'
import { rangeSelection, selectionRange } from '@gmi/excel-core'
import type { EditorView } from '../view/editorview'
import { dragPreviewKey } from '../view/types'
import { pluginNotice } from './notify'

// 纯函数：构造移动 entries（供单测）。源 raw+style 整体搬到目标偏移；公式不 shift（cut 语义）。
export function buildMove(
  sheet: SheetData, src: CellRange, dst: CellRange,
): { entries: { row: number; col: number; cell: Cell | null }[]; clearSource: boolean; reject: boolean } {
  const s = normalizeRange(src), d = normalizeRange(dst)
  const rows = s.er - s.sr + 1, cols = s.ec - s.sc + 1
  const clamped = clampRange(d, sheet.rowCount, sheet.colCount)
  // 相交（剪裁后仍与源相交）或目标落 merge → 拒绝
  if (rangesIntersect(s, clamped)) return { entries: [], clearSource: false, reject: true }
  if (sheet.merges.some(m => rangesIntersect(m, clamped) && !rangesIntersect(m, s)))
    return { entries: [], clearSource: false, reject: true }
  // 目标被 clampRange 收缩（src 装不下 clamped 目标）→ 拒绝，避免越界写 / 数据丢失
  const fitRows = clamped.er - clamped.sr + 1
  const fitCols = clamped.ec - clamped.sc + 1
  if (fitRows < rows || fitCols < cols) return { entries: [], clearSource: false, reject: true }
  const entries: { row: number; col: number; cell: Cell | null }[] = []
  for (let i = 0; i < rows; i++)
    for (let j = 0; j < cols; j++) {
      const cell = sheet.getCell(s.sr + i, s.sc + j) ?? null
      entries.push({ row: clamped.sr + i, col: clamped.sc + j, cell: cell ? { ...cell } : null })
    }
  return { entries, clearSource: true, reject: false }
}

export function dragmove(): Plugin {
  let dragging = false
  let src: CellRange | null = null
  let startCell: { row: number; col: number } | null = null

  // 拖拽中预览更新（不入 undo）；与 zoom/contextMenu 的非文档态 meta 同模式
  const setPreview = (view: EditorView, r: CellRange | null): void => {
    view.dispatch(view.state.tr.setMeta(dragPreviewKey, r).setMeta('addToHistory', false))
  }

  return new Plugin({
    key: dragPreviewKey,
    state: {
      init: (): CellRange | null => null,
      apply: (tr, prev: CellRange | null): CellRange | null => {
        const m = tr.getMeta(dragPreviewKey)
        return m === undefined ? prev : (m as CellRange | null)
      },
    },
    props: {
      handleMouseDown(view: EditorViewLike, e: MouseEvent, hit: HitResult): boolean {
        if (e.button !== 0) return false // 仅左键拖拽（同 selection，防右键菜单吞 mouseup）
        if (hit.region !== 'selborder') return false
        const v = view as EditorView
        dragging = true
        src = selectionRange(v.state.selection)
        startCell = v.pointerToCell(e.clientX, e.clientY)
        v.focus()
        return true
      },
      handleMouseMove(view: EditorViewLike, e: MouseEvent): boolean {
        if (!dragging || !src || !startCell) return false
        const v = view as EditorView
        const cur = v.pointerToCell(e.clientX, e.clientY)
        const dR = cur.row - startCell.row, dC = cur.col - startCell.col
        setPreview(v, normalizeRange({ sr: src.sr + dR, sc: src.sc + dC, er: src.er + dR, ec: src.ec + dC }))
        return true
      },
      handleMouseUp(view: EditorViewLike): boolean {
        if (!dragging || !src || !startCell) return false
        const v = view as EditorView
        const S = src
        dragging = false; src = null; startCell = null
        // 目标取上次 mousemove 写入的预览（避免依赖 event 坐标时序）
        const preview = v.state.getField(dragPreviewKey) as CellRange | null
        if (!preview) return true
        const sheet = v.state.activeSheet
        const { entries, clearSource, reject } = buildMove(sheet, S, preview)
        if (reject) {
          pluginNotice('目标区域与源相交、落在合并区或超出表格边界，无法移动')
          setPreview(v, null)
          return true
        }
        // 清预览 + 写目标 + 清源 + 选区：一个事务一次 undo（fillhandle 同模式）
        const tr = v.state.tr.setMeta(dragPreviewKey, null)
        if (entries.length) tr.setCells(v.state.doc.active, entries)
        if (clearSource) tr.clearRange(S)
        const n = normalizeRange(preview)
        tr.setSelection(rangeSelection(n, { row: n.sr, col: n.sc })).scrollIntoView()
        v.dispatch(tr)
        return true
      },
    },
  })
}
