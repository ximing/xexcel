// EditorView：命令式 Konva 视图。事件流：DOM 事件 → hitTest → 插件 someProp 链 →
// 未拦截走默认行为 → dispatch(tr) → applyTransaction → updateState → rAF 合帧重绘。
import Konva from 'konva'
import { CellAddr } from '../core/addr'
import { COL_HEADER_HEIGHT, ROW_HEADER_WIDTH, SheetData } from '../core/model'
import { EditorViewLike, PluginProps, PluginView } from '../core/plugin'
import { selectionRange, singleCell } from '../core/selection'
import type { SheetState } from '../core/state'
import type { Transaction } from '../core/transaction'
import { evaluatorFor } from '../formula/engine'
import { filterHiddenRows } from '../formula/filter'
import { openEditor } from './editbox'
import { GridGeometry } from './geometry'
import { FILL_HANDLE_SIZE, renderAll } from './layers'
import { contentViewport, hScrollbar, thumbHit, vScrollbar } from './scrollbar'
import { HitResult, Rect } from './types'

export interface DirectEditorProps {
  state: SheetState
}

const BORDER_TOLERANCE = 3 // 行列头边缘 ±3px 判定为调宽边界

export class EditorView implements EditorViewLike {
  readonly dom: HTMLDivElement
  readonly stage: Konva.Stage
  state: SheetState
  scrollX = 0
  scrollY = 0

  private readonly proxy: HTMLTextAreaElement
  private readonly gridLayer: Konva.Layer
  private readonly cellLayer: Konva.Layer
  private readonly overlayLayer: Konva.Layer
  private readonly pluginViews: (PluginView | undefined)[]
  private readonly listeners = new Set<() => void>()
  private readonly resizeObserver: ResizeObserver
  private geomCache: { sheet: SheetData; geom: GridGeometry } | null = null
  private rafId = 0
  private sbDrag: { axis: 'h' | 'v'; startScroll: number; startPos: number; ratio: number } | null = null

  constructor(mount: HTMLElement, props: DirectEditorProps) {
    this.state = props.state

    this.dom = document.createElement('div')
    this.dom.className = 'xcell-root'
    Object.assign(this.dom.style, {
      position: 'relative',
      overflow: 'hidden',
      outline: 'none',
      width: '100%',
      height: '100%',
    })
    const container = document.createElement('div')
    this.dom.appendChild(container)
    mount.appendChild(this.dom)

    this.stage = new Konva.Stage({
      container,
      width: mount.clientWidth || 0,
      height: mount.clientHeight || 0,
    })
    this.gridLayer = new Konva.Layer({ listening: false })
    this.cellLayer = new Konva.Layer({ listening: false })
    this.overlayLayer = new Konva.Layer({ listening: false })
    this.stage.add(this.gridLayer, this.cellLayer, this.overlayLayer)

    // 键盘/输入法/剪贴板事件代理（1px 隐形 textarea）
    this.proxy = document.createElement('textarea')
    this.proxy.className = 'xcell-input-proxy'
    Object.assign(this.proxy.style, {
      position: 'absolute',
      width: '1px',
      height: '1px',
      opacity: '0.01',
      top: '0',
      left: '0',
      border: 'none',
      padding: '0',
      resize: 'none',
      overflow: 'hidden',
    })
    this.dom.appendChild(this.proxy)

    this.pluginViews = this.state.plugins.map((p) => p.spec.view?.(this) || undefined)

    this.bindEvents()
    this.resizeObserver = new ResizeObserver(() => this.resize())
    this.resizeObserver.observe(mount)
    this.render()
  }

  // 按 activeSheet 引用缓存几何对象（行高列宽/冻结/筛选/隐藏变化都会得到新 SheetData）
  geometry(): GridGeometry {
    const sheet = this.state.activeSheet
    if (this.geomCache?.sheet !== sheet) {
      const extra = sheet.filter
        ? filterHiddenRows(this.state.doc.active, sheet, evaluatorFor(this.state.doc))
        : undefined
      this.geomCache = { sheet, geom: new GridGeometry(sheet, sheet.frozenRows, sheet.frozenCols, extra) }
    }
    return this.geomCache.geom
  }

  // 应用事务的唯一入口：applyTransaction → updateState → 插件 view.update → 重绘。
  // 宿主（React bridge 等）经 subscribe 感知 state 变化，不做事务级回调
  dispatch(tr: Transaction): void {
    const { state } = this.state.applyTransaction(tr)
    this.updateState(state)
    if (tr.scrolledIntoView) this.ensureVisible(state.selection.focus)
  }

  updateState(state: SheetState): void {
    const prev = this.state
    this.state = state
    // 切换活动表：视口归零，避免沿用前表的滚动偏移
    if (prev.doc.active !== state.doc.active) {
      this.scrollX = 0
      this.scrollY = 0
    }
    for (let i = 0; i < state.plugins.length; i++) {
      this.pluginViews[i]?.update?.(this, prev)
    }
    this.scheduleRender()
    for (const l of this.listeners) l()
  }

  subscribe(listener: () => void): () => void {
    this.listeners.add(listener)
    return () => {
      this.listeners.delete(listener)
    }
  }

  someProp(name: keyof PluginProps, fn: (prop: any) => boolean): boolean {
    for (const p of this.state.plugins) {
      const prop = p.spec.props?.[name]
      if (prop && fn(prop)) return true
    }
    return false
  }

  hitTest(clientX: number, clientY: number): HitResult {
    const rect = this.dom.getBoundingClientRect()
    const x = clientX - rect.left
    const y = clientY - rect.top
    if (x < 0 || y < 0 || x > this.stage.width() || y > this.stage.height()) {
      return { region: 'outside', row: -1, col: -1 }
    }
    // 滚动条区域优先于其他命中
    const geom0 = this.geometry()
    const vg = vScrollbar(geom0.contentHeight, this.scrollY, this.stage.width(), this.stage.height())
    if (vg && x >= vg.track.x && y >= vg.track.y && y <= vg.track.y + vg.track.h) {
      return { region: 'vscrollbar', row: -1, col: -1 }
    }
    const hg = hScrollbar(geom0.contentWidth, this.scrollX, this.stage.width(), this.stage.height())
    if (hg && y >= hg.track.y && x >= hg.track.x && x <= hg.track.x + hg.track.w) {
      return { region: 'hscrollbar', row: -1, col: -1 }
    }
    const geom = this.geometry()
    // 填充手柄：活动选区右下角 6px 方块 ±2px 容差（用活动格的视口位置）
    const sel = this.state.selection
    const sRange = selectionRange(sel)
    const br = this.cellViewportRect(sRange.er, sRange.ec)
    const hx = br.x + br.w
    const hy = br.y + br.h
    const half = FILL_HANDLE_SIZE / 2 + 2
    if (Math.abs(x - hx) <= half && Math.abs(y - hy) <= half) {
      return { region: 'fillhandle', row: sel.focus.row, col: sel.focus.col }
    }
    if (x < ROW_HEADER_WIDTH && y < COL_HEADER_HEIGHT) return { region: 'corner', row: -1, col: -1 }
    if (y < COL_HEADER_HEIGHT) {
      const cx = x - ROW_HEADER_WIDTH
      const col =
        cx < geom.frozenWidth
          ? geom.colAt(cx)
          : geom.colAt(geom.frozenWidth + this.scrollX + (cx - geom.frozenWidth))
      const left = col < geom.frozenCols ? geom.colLeft(col) : geom.frozenWidth + (geom.colLeft(col) - geom.frozenWidth - this.scrollX)
      const right = left + geom.colWidth(col)
      // 边界双侧 ±3px → 列调宽边界
      if (Math.abs(cx - right) <= BORDER_TOLERANCE) return { region: 'colborder', row: -1, col }
      if (col > 0) {
        const prevLeft =
          col - 1 < geom.frozenCols
            ? geom.colLeft(col - 1)
            : geom.frozenWidth + (geom.colLeft(col - 1) - geom.frozenWidth - this.scrollX)
        const prevRight = prevLeft + geom.colWidth(col - 1)
        if (Math.abs(cx - prevRight) <= BORDER_TOLERANCE) return { region: 'colborder', row: -1, col: col - 1 }
      }
      return { region: 'colheader', row: -1, col }
    }
    if (x < ROW_HEADER_WIDTH) {
      const cy = y - COL_HEADER_HEIGHT
      const row =
        cy < geom.frozenHeight
          ? geom.rowAt(cy)
          : geom.rowAt(geom.frozenHeight + this.scrollY + (cy - geom.frozenHeight))
      const top = row < geom.frozenRows ? geom.rowTop(row) : geom.frozenHeight + (geom.rowTop(row) - geom.frozenHeight - this.scrollY)
      const bottom = top + geom.rowHeight(row)
      if (Math.abs(cy - bottom) <= BORDER_TOLERANCE) return { region: 'rowborder', row, col: -1 }
      if (row > 0) {
        const prevTop =
          row - 1 < geom.frozenRows
            ? geom.rowTop(row - 1)
            : geom.frozenHeight + (geom.rowTop(row - 1) - geom.frozenHeight - this.scrollY)
        const prevBottom = prevTop + geom.rowHeight(row - 1)
        if (Math.abs(cy - prevBottom) <= BORDER_TOLERANCE) return { region: 'rowborder', row: row - 1, col: -1 }
      }
      return { region: 'rowheader', row, col: -1 }
    }
    const a = geom.cellAtContent(x - ROW_HEADER_WIDTH, y - COL_HEADER_HEIGHT, this.scrollX, this.scrollY)
    // 筛选箭头：表头行单元格右缘 18px 区域
    const f = this.state.activeSheet.filter
    if (f && a.row === f.range.sr && a.col >= f.range.sc && a.col <= f.range.ec) {
      const rect = this.cellViewportRect(a.row, a.col)
      if (x >= rect.x + rect.w - 18) return { region: 'filter', row: a.row, col: a.col }
    }
    return { region: 'cell', row: a.row, col: a.col }
  }

  // 指针 client 坐标 → 冻结感知 + clamp 到表内的单元格地址（供插件与 hitTest 共用）
  pointerToCell(clientX: number, clientY: number): CellAddr {
    const rect = this.dom.getBoundingClientRect()
    const geom = this.geometry()
    const sheet = this.state.activeSheet
    const a = geom.cellAtContent(
      clientX - rect.left - ROW_HEADER_WIDTH,
      clientY - rect.top - COL_HEADER_HEIGHT,
      this.scrollX,
      this.scrollY,
    )
    return {
      row: Math.max(0, Math.min(a.row, sheet.rowCount - 1)),
      col: Math.max(0, Math.min(a.col, sheet.colCount - 1)),
    }
  }

  cellViewportRect(row: number, col: number): Rect {
    const geom = this.geometry()
    const r = geom.cellRect(row, col)
    const sx = col < geom.frozenCols ? r.x : geom.frozenWidth + (r.x - geom.frozenWidth - this.scrollX)
    const sy = row < geom.frozenRows ? r.y : geom.frozenHeight + (r.y - geom.frozenHeight - this.scrollY)
    return { x: ROW_HEADER_WIDTH + sx, y: COL_HEADER_HEIGHT + sy, w: r.w, h: r.h }
  }

  ensureVisible(addr: CellAddr): void {
    const geom = this.geometry()
    const vp = contentViewport(
      geom.contentWidth,
      geom.contentHeight,
      this.stage.width() - ROW_HEADER_WIDTH,
      this.stage.height() - COL_HEADER_HEIGHT,
    )
    const viewW = vp.w
    const viewH = vp.h
    if (addr.col >= geom.frozenCols) {
      const left = geom.colLeft(addr.col) - geom.frozenWidth
      const right = left + geom.colWidth(addr.col)
      const span = viewW - geom.frozenWidth
      if (left < this.scrollX) this.scrollX = left
      else if (right > this.scrollX + span) this.scrollX = right - span
    }
    if (addr.row >= geom.frozenRows) {
      const top = geom.rowTop(addr.row) - geom.frozenHeight
      const bottom = top + geom.rowHeight(addr.row)
      const span = viewH - geom.frozenHeight
      if (top < this.scrollY) this.scrollY = top
      else if (bottom > this.scrollY + span) this.scrollY = bottom - span
    }
    this.clampScroll()
    this.render()
  }

  focus(): void {
    this.proxy.focus({ preventScroll: true })
  }

  // 清空 proxy 累计的输入字符（keymap 开编辑器前调用，防脏值混入下次输入）
  clearProxy(): void {
    this.proxy.value = ''
  }

  render(): void {
    renderAll(
      this.gridLayer,
      this.cellLayer,
      this.overlayLayer,
      this.state,
      this.geometry(),
      this.scrollX,
      this.scrollY,
      this.stage.width(),
      this.stage.height(),
    )
  }

  destroy(): void {
    if (this.rafId) cancelAnimationFrame(this.rafId)
    this.resizeObserver.disconnect()
    window.removeEventListener('mousemove', this.onMouseMove)
    window.removeEventListener('mouseup', this.onMouseUp)
    for (const pv of this.pluginViews) pv?.destroy?.()
    this.stage.destroy()
    this.dom.remove()
    this.listeners.clear()
  }

  private scheduleRender(): void {
    if (this.rafId) return
    this.rafId = requestAnimationFrame(() => {
      this.rafId = 0
      this.render()
    })
  }

  private resize(): void {
    this.stage.size({ width: this.dom.clientWidth, height: this.dom.clientHeight })
    this.clampScroll()
    this.render()
  }

  private maxScrollX(): number {
    const vp = contentViewport(
      this.geometry().contentWidth,
      this.geometry().contentHeight,
      this.stage.width() - ROW_HEADER_WIDTH,
      this.stage.height() - COL_HEADER_HEIGHT,
    )
    return Math.max(0, this.geometry().contentWidth - vp.w)
  }

  private maxScrollY(): number {
    const vp = contentViewport(
      this.geometry().contentWidth,
      this.geometry().contentHeight,
      this.stage.width() - ROW_HEADER_WIDTH,
      this.stage.height() - COL_HEADER_HEIGHT,
    )
    return Math.max(0, this.geometry().contentHeight - vp.h)
  }

  // 插件（selection 边缘自动滚动）也需要钳位，故公开
  clampScroll(): void {
    this.scrollX = Math.max(0, Math.min(this.scrollX, this.maxScrollX()))
    this.scrollY = Math.max(0, Math.min(this.scrollY, this.maxScrollY()))
  }

  private bindEvents(): void {
    const content = this.stage.content
    content.addEventListener('mousedown', this.onMouseDown)
    content.addEventListener('dblclick', this.onDblClick)
    content.addEventListener('wheel', this.onWheel, { passive: false })
    window.addEventListener('mousemove', this.onMouseMove)
    window.addEventListener('mouseup', this.onMouseUp)
    this.proxy.addEventListener('keydown', this.onKeyDown)
    this.proxy.addEventListener('paste', this.onPaste)
    this.proxy.addEventListener('copy', this.onCopy)
    this.proxy.addEventListener('cut', this.onCut)
  }

  private onMouseDown = (e: MouseEvent): void => {
    const hit = this.hitTest(e.clientX, e.clientY)
    if (hit.region === 'hscrollbar' || hit.region === 'vscrollbar') {
      const geom = this.geometry()
      const g =
        hit.region === 'vscrollbar'
          ? vScrollbar(geom.contentHeight, this.scrollY, this.stage.width(), this.stage.height())
          : hScrollbar(geom.contentWidth, this.scrollX, this.stage.width(), this.stage.height())
      if (!g) return
      const rect = this.dom.getBoundingClientRect()
      const x = e.clientX - rect.left
      const y = e.clientY - rect.top
      const axis = hit.region === 'vscrollbar' ? 'v' : 'h'
      if (thumbHit(g, x, y)) {
        this.sbDrag = { axis, startScroll: axis === 'v' ? this.scrollY : this.scrollX, startPos: axis === 'v' ? y : x, ratio: g.ratio }
      } else {
        // 点轨道翻页（向点击方向滚一个视口）
        const page = axis === 'v' ? this.stage.height() - COL_HEADER_HEIGHT : this.stage.width() - ROW_HEADER_WIDTH
        const onThumbSide = axis === 'v' ? y > g.thumb.y + g.thumb.h : x > g.thumb.x + g.thumb.w
        if (axis === 'v') this.scrollY += onThumbSide ? page : -page
        else this.scrollX += onThumbSide ? page : -page
        this.clampScroll()
        this.render()
      }
      this.focus()
      return
    }
    if (this.someProp('handleMouseDown', (p) => p(this, e, hit))) return
    // 默认行为：点击单元格 → 单格选区
    if (hit.region === 'cell') {
      this.dispatch(this.state.tr.setSelection(singleCell(hit.row, hit.col)).scrollIntoView())
      this.focus()
    }
  }

  private onDblClick = (e: MouseEvent): void => {
    const hit = this.hitTest(e.clientX, e.clientY)
    if (this.someProp('handleDoubleClick', (p) => p(this, e, hit))) return
    if (hit.region === 'cell') {
      const m = this.state.activeSheet.mergeAt(hit.row, hit.col)
      openEditor(this, m ? { row: m.sr, col: m.sc } : { row: hit.row, col: hit.col })
    }
  }

  private onMouseMove = (e: MouseEvent): void => {
    if (this.sbDrag) {
      const rect = this.dom.getBoundingClientRect()
      const pos = this.sbDrag.axis === 'v' ? e.clientY - rect.top : e.clientX - rect.left
      const next = this.sbDrag.startScroll + (pos - this.sbDrag.startPos) * this.sbDrag.ratio
      if (this.sbDrag.axis === 'v') this.scrollY = next
      else this.scrollX = next
      this.clampScroll()
      this.render()
      return
    }
    const hit = this.hitTest(e.clientX, e.clientY)
    this.someProp('handleMouseMove', (p) => p(this, e, hit))
  }

  private onMouseUp = (e: MouseEvent): void => {
    if (this.sbDrag) {
      this.sbDrag = null
      return
    }
    const hit = this.hitTest(e.clientX, e.clientY)
    this.someProp('handleMouseUp', (p) => p(this, e, hit))
  }

  private onWheel = (e: WheelEvent): void => {
    e.preventDefault()
    this.scrollX += e.deltaX
    this.scrollY += e.deltaY
    this.clampScroll()
    this.render()
  }

  private onKeyDown = (e: KeyboardEvent): void => {
    if (!this.someProp('handleKeyDown', (p) => p(this, e))) {
      // 未拦截的可打印输入会落入 proxy textarea，清除避免脏值累积
      if (e.key.length === 1) this.proxy.value = ''
    }
  }

  private onPaste = (e: ClipboardEvent): void => {
    const text = e.clipboardData?.getData('text/plain') ?? ''
    if (this.someProp('handlePaste', (p) => p(this, text))) e.preventDefault()
  }

  private onCopy = (e: ClipboardEvent): void => {
    if (this.someProp('handleCopy', (p) => p(this, false, e))) e.preventDefault()
  }

  private onCut = (e: ClipboardEvent): void => {
    if (this.someProp('handleCopy', (p) => p(this, true, e))) e.preventDefault()
  }
}
