// 选区与行列调宽（resize）插件。职责：
// - 单元格 mousedown → 记 anchor 进入拖拽；mousemove 更新 focus；shift+mousedown 扩展 focus
// - 行头 mousedown → 选整行（可拖多行）；列头同理；corner → selectAll
// - 拖拽指针距视口边缘 <24px 时 rAF 自动滚动并重算 focus，mouseup 停止
// - 列头右缘/行头下缘（colborder/rowborder）拖拽调宽：参考线存 resizeGuideKey state field
//   （layers 读取画虚线），mouseup dispatch tr.resize（最小 20px）；双击边框恢复默认尺寸
// 全部经 props 拦截 + dispatch transaction，不直接改 doc。拖拽态存插件闭包变量。
import { CellAddr } from '../core/addr'
import { selectAll } from '../core/commands'
import { COL_HEADER_HEIGHT, ROW_HEADER_WIDTH } from '../core/model'
import { EditorViewLike, HitResult, Plugin } from '../core/plugin'
import { singleCell } from '../core/selection'
import type { EditorView } from '../view/editorview'
import { ResizeGuide, resizeGuideKey } from '../view/types'

const EDGE = 24 // 距视口边缘不足 24px 触发自动滚动
const MIN_SIZE = 20 // 行列最小尺寸

type Drag =
  | { kind: 'select'; mode: 'cell' | 'row' | 'col' }
  | { kind: 'resize'; axis: 'row' | 'col'; index: number; startSize: number; size: number }

export function selection(): Plugin {
  let drag: Drag | null = null
  let activeView: EditorView | null = null
  let lastClient = { x: 0, y: 0 }
  let rafId = 0

  const stopAutoScroll = (): void => {
    if (rafId) cancelAnimationFrame(rafId)
    rafId = 0
  }

  // 指针 client 坐标 → clamp 到表内的单元格地址（越界钳到边缘行/列；冻结感知走 view.pointerToCell）
  const addrAt = (view: EditorView, clientX: number, clientY: number): CellAddr =>
    view.pointerToCell(clientX, clientY)

  // 按拖拽模式重算 focus 并 dispatch（focus 未变则仅重绘，避免空事务刷屏）
  const applyDragFocus = (view: EditorView): void => {
    if (!drag || drag.kind !== 'select') return
    const sheet = view.state.activeSheet
    const addr = addrAt(view, lastClient.x, lastClient.y)
    const focus: CellAddr =
      drag.mode === 'cell'
        ? addr
        : drag.mode === 'row'
          ? { row: addr.row, col: sheet.colCount - 1 }
          : { row: sheet.rowCount - 1, col: addr.col }
    const sel = view.state.selection
    if (sel.focus.row === focus.row && sel.focus.col === focus.col) {
      view.render() // 仅滚动发生变化
      return
    }
    view.dispatch(view.state.tr.setSelection({ anchor: sel.anchor, focus }))
  }

  // 指针相对视口四边的滚动增量（0 表示该轴不滚）
  const edgeDelta = (view: EditorView): { dx: number; dy: number } => {
    const rect = view.dom.getBoundingClientRect()
    const x = lastClient.x - rect.left
    const y = lastClient.y - rect.top
    const w = view.stage.width()
    const h = view.stage.height()
    let dx = 0
    let dy = 0
    if (x >= 0 && x < EDGE) dx = -Math.max(2, Math.round((EDGE - x) / 2))
    else if (x > w - EDGE && x <= w) dx = Math.max(2, Math.round((x - (w - EDGE)) / 2))
    if (y >= 0 && y < EDGE) dy = -Math.max(2, Math.round((EDGE - y) / 2))
    else if (y > h - EDGE && y <= h) dy = Math.max(2, Math.round((y - (h - EDGE)) / 2))
    return { dx, dy }
  }

  const tick = (): void => {
    rafId = 0
    const view = activeView
    if (!view || !drag || drag.kind !== 'select') return
    const { dx, dy } = edgeDelta(view)
    if (!dx && !dy) return
    view.scrollX += dx
    view.scrollY += dy
    view.clampScroll()
    applyDragFocus(view)
    rafId = requestAnimationFrame(tick)
  }

  // resize 拖拽：按指针内容坐标实时计算尺寸并更新参考线
  const updateResize = (view: EditorView): void => {
    if (!drag || drag.kind !== 'resize') return
    const rect = view.dom.getBoundingClientRect()
    const geom = view.geometry()
    const pos =
      drag.axis === 'col'
        ? lastClient.x - rect.left - ROW_HEADER_WIDTH + view.scrollX
        : lastClient.y - rect.top - COL_HEADER_HEIGHT + view.scrollY
    const left = drag.axis === 'col' ? geom.colLeft(drag.index) : geom.rowTop(drag.index)
    drag.size = Math.max(MIN_SIZE, Math.round(pos - left))
    const guide: ResizeGuide = { axis: drag.axis, pos: left + drag.size }
    view.dispatch(view.state.tr.setMeta(resizeGuideKey, guide))
  }

  const startSelectDrag = (v: EditorView, e: MouseEvent, mode: 'cell' | 'row' | 'col'): void => {
    drag = { kind: 'select', mode }
    activeView = v
    lastClient = { x: e.clientX, y: e.clientY }
  }

  return new Plugin({
    key: resizeGuideKey,
    state: {
      init: (): ResizeGuide | null => null,
      apply: (tr, value: ResizeGuide | null): ResizeGuide | null => {
        const v = tr.getMeta(resizeGuideKey)
        return v === undefined ? value : (v as ResizeGuide | null)
      },
    },
    props: {
      handleMouseDown(view: EditorViewLike, e: MouseEvent, hit: HitResult): boolean {
        const v = view as EditorView
        const sheet = v.state.activeSheet
        switch (hit.region) {
          case 'cell': {
            const m = sheet.mergeAt(hit.row, hit.col)
            if (e.shiftKey) {
              v.dispatch(
                v.state.tr.setSelection({ anchor: v.state.selection.anchor, focus: { row: hit.row, col: hit.col } }),
              )
            } else if (m) {
              // 点击合并区任意位置 → 选中整个区
              v.dispatch(
                v.state.tr.setSelection({ anchor: { row: m.sr, col: m.sc }, focus: { row: m.er, col: m.ec } }),
              )
              startSelectDrag(v, e, 'cell')
            } else {
              v.dispatch(v.state.tr.setSelection(singleCell(hit.row, hit.col)))
              startSelectDrag(v, e, 'cell')
            }
            v.focus()
            return true
          }
          case 'rowheader': {
            const focus = { row: hit.row, col: sheet.colCount - 1 }
            const anchor = e.shiftKey ? v.state.selection.anchor : { row: hit.row, col: 0 }
            v.dispatch(v.state.tr.setSelection({ anchor, focus }))
            if (!e.shiftKey) startSelectDrag(v, e, 'row')
            v.focus()
            return true
          }
          case 'colheader': {
            const focus = { row: sheet.rowCount - 1, col: hit.col }
            const anchor = e.shiftKey ? v.state.selection.anchor : { row: 0, col: hit.col }
            v.dispatch(v.state.tr.setSelection({ anchor, focus }))
            if (!e.shiftKey) startSelectDrag(v, e, 'col')
            v.focus()
            return true
          }
          case 'corner': {
            selectAll(v.state, (tr) => v.dispatch(tr))
            v.focus()
            return true
          }
          case 'colborder':
          case 'rowborder': {
            const axis = hit.region === 'colborder' ? 'col' : 'row'
            const index = axis === 'col' ? hit.col : hit.row
            const startSize = axis === 'col' ? sheet.colWidth(index) : sheet.rowHeight(index)
            drag = { kind: 'resize', axis, index, startSize, size: startSize }
            activeView = v
            lastClient = { x: e.clientX, y: e.clientY }
            v.focus()
            return true
          }
          default:
            return false // fillhandle / outside 交给后续插件
        }
      },
      handleMouseMove(view: EditorViewLike, e: MouseEvent): boolean {
        if (!drag) return false
        const v = view as EditorView
        lastClient = { x: e.clientX, y: e.clientY }
        if (drag.kind === 'resize') {
          updateResize(v)
          return true
        }
        applyDragFocus(v)
        const { dx, dy } = edgeDelta(v)
        if (dx || dy) {
          if (!rafId) rafId = requestAnimationFrame(tick)
        } else {
          stopAutoScroll()
        }
        return true
      },
      handleMouseUp(view: EditorViewLike): boolean {
        if (!drag) return false
        const v = view as EditorView
        stopAutoScroll()
        if (drag.kind === 'resize') {
          const tr = v.state.tr.setMeta(resizeGuideKey, null)
          if (drag.size !== drag.startSize) tr.resize(drag.axis, drag.index, drag.size)
          v.dispatch(tr)
        }
        drag = null
        activeView = null
        return true
      },
      handleDoubleClick(view: EditorViewLike, _e: MouseEvent, hit: HitResult): boolean {
        if (hit.region !== 'colborder' && hit.region !== 'rowborder') return false
        const v = view as EditorView
        const axis = hit.region === 'colborder' ? 'col' : 'row'
        v.dispatch(v.state.tr.resize(axis, axis === 'col' ? hit.col : hit.row, null))
        return true
      },
    },
  })
}
