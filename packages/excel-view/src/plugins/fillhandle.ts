// 填充手柄插件。职责：
// - mousedown（region==='fillhandle'）→ 记当前选区为源 range 进入拖拽
// - mousemove 计算目标 range：源 + 主轴方向延伸（|dr| vs |dc| 取大者，只沿该轴），
//   存 plugin state field（key=fillPreviewKey，值 CellRange|null），layers 读非空画 1px 虚线框
// - mouseup → SetCellsStep：源区域沿轴重复平铺；源为单行/单列且全部数字（raw 非 = 开头、
//   evaluator.get 为 number）→ 等差序列（步长 =(末值-首值)/(n-1)）；dispatch 并 setSelection
//   到新区域、清 preview
// 全部经 dispatch transaction，不直接改 doc。拖拽态存插件闭包变量。
import { CellRange, rangesEqual } from '@xexcel/core'
import { Cell, SheetData } from '@xexcel/core'
import { EditorViewLike, Plugin } from '@xexcel/core'
import { rangeSelection, selectionRange } from '@xexcel/core'
import type { Transaction } from '@xexcel/core'
import type { FormulaValue } from '@xexcel/core'
import { evaluatorFor } from '@xexcel/core'
import { shiftFormula } from '@xexcel/core'
import type { EditorView } from '../view/editorview'
import { fillPreviewKey } from '../view/types'

export function fillhandle(): Plugin {
  let source: CellRange | null = null
  let target: CellRange | null = null

  // 指针位置 → 延伸后的目标 range（指针在源内 → null 不预览；冻结感知走 view.pointerToCell）
  const targetAt = (view: EditorView, clientX: number, clientY: number): CellRange | null => {
    if (!source) return null
    const a = view.pointerToCell(clientX, clientY)
    const row = a.row
    const col = a.col
    const src = source
    const dr = row < src.sr ? row - src.sr : row > src.er ? row - src.er : 0
    const dc = col < src.sc ? col - src.sc : col > src.ec ? col - src.ec : 0
    if (Math.abs(dr) >= Math.abs(dc)) {
      if (dr > 0) return { ...src, er: row }
      if (dr < 0) return { ...src, sr: row }
    } else {
      if (dc > 0) return { ...src, ec: col }
      if (dc < 0) return { ...src, sc: col }
    }
    return null
  }

  const setPreview = (view: EditorView, next: CellRange | null): void => {
    const same =
      (target === null && next === null) || (target !== null && next !== null && rangesEqual(target, next))
    if (same) return
    target = next
    view.dispatch(view.state.tr.setMeta(fillPreviewKey, target))
  }

  return new Plugin({
    key: fillPreviewKey,
    state: {
      init: (): CellRange | null => null,
      apply: (tr, value: CellRange | null): CellRange | null => {
        const v = tr.getMeta(fillPreviewKey)
        return v === undefined ? value : (v as CellRange | null)
      },
    },
    props: {
      handleMouseDown(view: EditorViewLike, _e: MouseEvent, hit): boolean {
        if (_e.button !== 0) return false // 仅左键拖拽（同 selection，防右键菜单吞 mouseup）
        if (hit.region !== 'fillhandle') return false
        const v = view as EditorView
        source = selectionRange(v.state.selection)
        target = null
        v.focus()
        return true
      },
      handleMouseMove(view: EditorViewLike, e: MouseEvent): boolean {
        if (!source) return false
        setPreview(view as EditorView, targetAt(view as EditorView, e.clientX, e.clientY))
        return true
      },
      handleMouseUp(view: EditorViewLike): boolean {
        if (!source) return false
        const v = view as EditorView
        const src = source
        const dst = target
        source = null
        target = null
        const tr = v.state.tr.setMeta(fillPreviewKey, null)
        if (dst) buildFill(v, tr, src, dst)
        v.dispatch(tr)
        return true
      },
    },
  })
}

// 生成填充内容：延伸区平铺源区域（公式按各自偏移量调整引用）；满足等差条件时写序列值。
// 纯函数提取供单测；getValue 通常为 (r,c) => evaluatorFor(doc).get(sheetId, r, c)。
export function computeFillEntries(
  sheet: SheetData,
  getValue: (row: number, col: number) => FormulaValue,
  src: CellRange,
  dst: CellRange,
): { row: number; col: number; cell: Cell | null }[] {
  const srcRows = src.er - src.sr + 1
  const srcCols = src.ec - src.sc + 1
  const vertical = dst.er > src.er || dst.sr < src.sr // 主轴方向（targetAt 保证只沿一轴延伸）
  const entries: { row: number; col: number; cell: Cell | null }[] = []

  // 等差判定：源单列纵向延伸 / 源单行横向延伸，且每格 raw 非公式、求值为 number
  let series: number[] | null = null
  if ((vertical && srcCols === 1) || (!vertical && srcRows === 1)) {
    const nums: number[] = []
    let ok = true
    const n = vertical ? srcRows : srcCols
    for (let i = 0; i < n; i++) {
      const r = vertical ? src.sr + i : src.sr
      const c = vertical ? src.sc : src.sc + i
      const cell = sheet.getCell(r, c)
      if (!cell || cell.raw.startsWith('=')) {
        ok = false
        break
      }
      const val = getValue(r, c)
      if (typeof val !== 'number') {
        ok = false
        break
      }
      nums.push(val)
    }
    if (ok && nums.length >= 2) series = nums
  }

  if (series) {
    const step = (series[series.length - 1] - series[0]) / (series.length - 1)
    const last = series[series.length - 1]
    const first = series[0]
    const numCell = (v: number): Cell => ({ raw: String(v) })
    if (vertical) {
      for (let r = src.er + 1; r <= dst.er; r++) entries.push({ row: r, col: src.sc, cell: numCell(last + step * (r - src.er)) })
      for (let r = dst.sr; r < src.sr; r++) entries.push({ row: r, col: src.sc, cell: numCell(first - step * (src.sr - r)) })
    } else {
      for (let c = src.ec + 1; c <= dst.ec; c++) entries.push({ row: src.sr, col: c, cell: numCell(last + step * (c - src.ec)) })
      for (let c = dst.sc; c < src.sc; c++) entries.push({ row: src.sr, col: c, cell: numCell(first - step * (src.sc - c)) })
    }
    return entries
  }

  // 平铺：源区域按周期重复（公式 shift、样式拷贝；源空格 → 清目标格）
  for (let r = dst.sr; r <= dst.er; r++) {
    for (let c = dst.sc; c <= dst.ec; c++) {
      if (r >= src.sr && r <= src.er && c >= src.sc && c <= src.ec) continue
      const srow = src.sr + (((r - src.sr) % srcRows) + srcRows) % srcRows
      const scol = src.sc + (((c - src.sc) % srcCols) + srcCols) % srcCols
      const s = sheet.getCell(srow, scol)
      if (!s) {
        entries.push({ row: r, col: c, cell: null })
        continue
      }
      const raw = s.raw.startsWith('=') ? shiftFormula(s.raw, r - srow, c - scol) : s.raw
      entries.push({
        row: r,
        col: c,
        cell: { raw, ...(s.style ? { style: { ...s.style } } : {}) },
      })
    }
  }
  return entries
}

// 填充后选区切到整个目标区域。
function buildFill(view: EditorView, tr: Transaction, src: CellRange, dst: CellRange): void {
  const state = view.state
  const sheetId = state.doc.active
  const ev = evaluatorFor(state.doc)
  const entries = computeFillEntries(state.activeSheet, (r, c) => ev.get(sheetId, r, c), src, dst)
  if (entries.length) tr.setCells(sheetId, entries)
  tr.setSelection(rangeSelection(dst, { row: dst.sr, col: dst.sc })).scrollIntoView()
}
