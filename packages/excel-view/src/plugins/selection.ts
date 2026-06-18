// 选区与行列调宽（resize）插件。职责：
// - 单元格 mousedown → 进入拖拽态；mousemove 经 extendActiveRange 更新活动区域边界（固定锚点生长）；Shift+mousedown 扩展活动区域
// - 行头 mousedown → 选整行（可拖多行）；列头同理；corner → selectAll
// - 拖拽指针距视口边缘 <24px 时 rAF 自动滚动并重算 focus，mouseup 停止
// - 列头右缘/行头下缘（colborder/rowborder）拖拽调宽：参考线存 resizeGuideKey state field
//   （layers 读取画虚线），mouseup dispatch tr.resize（最小 20px）；双击边框自适应内容尺寸
//   （整行/整列选区包含目标时批量；空行/列经 resize null 恢复默认）
// 全部经 props 拦截 + dispatch transaction，不直接改 doc。拖拽态存插件闭包变量。
import { CellAddr, rangeContains } from '@gmi/excel-core'
import { selectAll } from '@gmi/excel-core'
import { CellStyle, COL_HEADER_HEIGHT, ROW_HEADER_WIDTH } from '@gmi/excel-core'
import { EditorViewLike, HitResult, Plugin } from '@gmi/excel-core'
import { appendRange, extendActiveRange, rangeSelection, selectionRange, singleCell, toggleRange } from '@gmi/excel-core'
import { evaluatorFor } from '@gmi/excel-core'
import type { EditorView } from '../view/editorview'
import { measureTextWidth, optimalColWidth, optimalRowHeight } from '../view/measure'
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

  // 拖拽更新活动区域边界（Shift+drag 与无 Ctrl 的 drag 都走此路径）
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
    if (sel.activeCell.row === focus.row && sel.activeCell.col === focus.col) {
      view.render() // 仅滚动发生变化
      return
    }
    view.dispatch(view.state.tr.setSelection(extendActiveRange(sel, focus)))
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

  // resize 拖拽：按指针内容坐标实时计算尺寸并更新参考线。
  // geom 已为缩放后坐标，drag.size 记缩放后像素（guide 直接用）；提交时 /zoom 折算回基准单位
  const updateResize = (view: EditorView): void => {
    if (!drag || drag.kind !== 'resize') return
    const rect = view.dom.getBoundingClientRect()
    const geom = view.geometry()
    const z = view.zoom()
    // 冻结感知：滚动区内容坐标 = 屏幕坐标 + scroll（frozenWidth 抵消），
    // 冻结区内容坐标 = 屏幕坐标（不含 scroll），否则拖冻结行列边框时尺寸跳变
    const pos =
      drag.axis === 'col'
        ? (() => {
            const cx = lastClient.x - rect.left - ROW_HEADER_WIDTH * z
            return cx < geom.frozenWidth ? cx : cx + view.scrollX
          })()
        : (() => {
            const cy = lastClient.y - rect.top - COL_HEADER_HEIGHT * z
            return cy < geom.frozenHeight ? cy : cy + view.scrollY
          })()
    const left = drag.axis === 'col' ? geom.colLeft(drag.index) : geom.rowTop(drag.index)
    drag.size = Math.max(MIN_SIZE * z, Math.round(pos - left))
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
        // 仅左键发起选择/调宽拖拽：右键 mousedown 后 contextmenu 菜单会拦截 mouseup 冒泡，
        // 若允许右键进入拖拽态，drag 永不清除，mousemove 会持续改选区（issue #1）
        if (e.button !== 0) return false
        const v = view as EditorView
        const sheet = v.state.activeSheet
        switch (hit.region) {
          case 'cell': {
            const me = e
            const sel = v.state.selection
            const m = sheet.mergeAt(hit.row, hit.col)
            // 点在合并区内 → 整个合并区作单区域（Ctrl 追加）
            if (m) {
              if (me.ctrlKey || me.metaKey) {
                v.dispatch(v.state.tr.setSelection(appendRange(sel, m, { row: m.sr, col: m.sc })))
              } else {
                v.dispatch(v.state.tr.setSelection(rangeSelection(m, { row: m.sr, col: m.sc })))
              }
              v.focus()
              return true
            }
            // Ctrl+click：格已在选区 → 反选（LIFO 移除最后含该格的 range）；否则追加单格区域
            // Shift+click：扩展活动区域到该格
            // 无修饰：singleCell 重置
            if (me.shiftKey) {
              v.dispatch(v.state.tr.setSelection(extendActiveRange(sel, { row: hit.row, col: hit.col })).scrollIntoView())
            } else if (me.ctrlKey || me.metaKey) {
              const inSel = sel.ranges.some(r => rangeContains(r, hit.row, hit.col))
              v.dispatch(v.state.tr.setSelection(
                inSel ? toggleRange(sel, hit.row, hit.col) : appendRange(sel, { sr: hit.row, sc: hit.col, er: hit.row, ec: hit.col }, { row: hit.row, col: hit.col }),
              ))
            } else {
              v.dispatch(v.state.tr.setSelection(singleCell(hit.row, hit.col)))
            }
            startSelectDrag(v, e, 'cell')
            v.focus()
            return true
          }
          case 'rowheader': {
            // 行头：整行 range（Ctrl 追加，否则替换）
            const sel = v.state.selection
            const focus: CellAddr = { row: hit.row, col: sheet.colCount - 1 }
            const fullRow = { sr: hit.row, sc: 0, er: hit.row, ec: sheet.colCount - 1 }
            if (e.shiftKey) {
              v.dispatch(v.state.tr.setSelection(extendActiveRange(sel, focus)).scrollIntoView())
            } else if (e.ctrlKey || e.metaKey) {
              v.dispatch(v.state.tr.setSelection(appendRange(sel, fullRow, focus)))
            } else {
              v.dispatch(v.state.tr.setSelection(rangeSelection(fullRow, focus)))
              startSelectDrag(v, e, 'row')
            }
            v.focus()
            return true
          }
          case 'colheader': {
            // 列头：整列 range（Ctrl 追加，否则替换）
            const sel = v.state.selection
            const focus: CellAddr = { row: sheet.rowCount - 1, col: hit.col }
            const fullCol = { sr: 0, sc: hit.col, er: sheet.rowCount - 1, ec: hit.col }
            if (e.shiftKey) {
              v.dispatch(v.state.tr.setSelection(extendActiveRange(sel, focus)).scrollIntoView())
            } else if (e.ctrlKey || e.metaKey) {
              v.dispatch(v.state.tr.setSelection(appendRange(sel, fullCol, focus)))
            } else {
              v.dispatch(v.state.tr.setSelection(rangeSelection(fullCol, focus)))
              startSelectDrag(v, e, 'col')
            }
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
          // drag.size 为缩放后像素，折算回基准单位再写模型（startSize 也是基准单位）
          const size = Math.round(drag.size / v.zoom())
          if (size !== drag.startSize) tr.resize(drag.axis, drag.index, size)
          v.dispatch(tr)
          drag = null
          activeView = null
          return true
        }
        // select 拖拽的 mouseup 只做清理、无 doc 变更：不拦截，放行给后续插件（如格式刷 painter）
        drag = null
        activeView = null
        return false
      },
      handleDoubleClick(view: EditorViewLike, _e: MouseEvent, hit: HitResult): boolean {
        if (hit.region !== 'colborder' && hit.region !== 'rowborder') return false
        const v = view as EditorView
        const state = v.state
        const sheet = state.activeSheet
        const sheetId = state.doc.active
        const ev = evaluatorFor(state.doc)
        const axis = hit.region === 'colborder' ? 'col' : 'row'
        const index = axis === 'col' ? hit.col : hit.row
        // 选区覆盖整列/行且包含双击目标 → 批量；否则只作用目标
        const sel = selectionRange(state.selection)
        let indices = [index]
        if (axis === 'col' && sel.sr === 0 && sel.er === sheet.rowCount - 1 && index >= sel.sc && index <= sel.ec) {
          indices = []
          for (let c = sel.sc; c <= sel.ec; c++) indices.push(c)
        }
        if (axis === 'row' && sel.sc === 0 && sel.ec === sheet.colCount - 1 && index >= sel.sr && index <= sel.er) {
          indices = []
          for (let r = sel.sr; r <= sel.er; r++) indices.push(r)
        }
        const tr = state.tr
        for (const i of indices) {
          if (axis === 'col') {
            const items: { text: string; style?: CellStyle }[] = []
            for (let r = 0; r < sheet.rowCount; r++) {
              const cell = sheet.getCell(r, i)
              if (cell && cell.raw !== '') items.push({ text: ev.displayText(sheetId, r, i), style: cell.style })
            }
            tr.resize('col', i, optimalColWidth(items, measureTextWidth))
          } else {
            const items: { style?: CellStyle }[] = []
            for (let c = 0; c < sheet.colCount; c++) {
              const cell = sheet.getCell(i, c)
              if (cell && cell.raw !== '') items.push({ style: cell.style })
            }
            tr.resize('row', i, optimalRowHeight(items))
          }
        }
        v.dispatch(tr)
        return true
      },
    },
  })
}
