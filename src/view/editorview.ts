// EditorView：命令式 Konva 视图。事件流：DOM 事件 → hitTest → 插件 someProp 链 →
// 未拦截走默认行为 → dispatch(tr) → applyTransaction → updateState → rAF 合帧重绘。
import Konva from 'konva'
import { CellAddr } from '../core/addr'
import { COL_HEADER_HEIGHT, ROW_HEADER_WIDTH, SheetData } from '../core/model'
import { EditorViewLike, PluginProps, PluginView } from '../core/plugin'
import { selectionRange, singleCell } from '../core/selection'
import type { SheetState } from '../core/state'
import type { Transaction } from '../core/transaction'
import { openEditor } from './editbox'
import { GridGeometry } from './geometry'
import { FILL_HANDLE_SIZE, renderAll } from './layers'
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

  // 按 activeSheet 引用缓存几何对象（行高列宽变化会得到新 SheetData）
  geometry(): GridGeometry {
    const sheet = this.state.activeSheet
    if (this.geomCache?.sheet !== sheet) {
      this.geomCache = { sheet, geom: new GridGeometry(sheet) }
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
    const geom = this.geometry()
    // 填充手柄：活动选区 rangeRect 右下角 6px 方块 ±2px 容差
    const sel = this.state.selection
    const rr = geom.rangeRect(selectionRange(sel))
    const hx = ROW_HEADER_WIDTH + rr.x + rr.w - this.scrollX
    const hy = COL_HEADER_HEIGHT + rr.y + rr.h - this.scrollY
    const half = FILL_HANDLE_SIZE / 2 + 2
    if (Math.abs(x - hx) <= half && Math.abs(y - hy) <= half) {
      return { region: 'fillhandle', row: sel.focus.row, col: sel.focus.col }
    }
    if (x < ROW_HEADER_WIDTH && y < COL_HEADER_HEIGHT) return { region: 'corner', row: -1, col: -1 }
    if (y < COL_HEADER_HEIGHT) {
      const cx = x - ROW_HEADER_WIDTH + this.scrollX
      const col = geom.colAt(cx)
      // 列边界双侧 ±3px → 列调宽边界（bsearch 在边界返回下一列，故需补左缘判定）
      if (Math.abs(cx - geom.colLeft(col + 1)) <= BORDER_TOLERANCE) {
        return { region: 'colborder', row: -1, col }
      }
      if (col > 0 && Math.abs(cx - geom.colLeft(col)) <= BORDER_TOLERANCE) {
        return { region: 'colborder', row: -1, col: col - 1 }
      }
      return { region: 'colheader', row: -1, col }
    }
    if (x < ROW_HEADER_WIDTH) {
      const cy = y - COL_HEADER_HEIGHT + this.scrollY
      const row = geom.rowAt(cy)
      // 行边界双侧 ±3px → 行调宽边界
      if (Math.abs(cy - geom.rowTop(row + 1)) <= BORDER_TOLERANCE) {
        return { region: 'rowborder', row, col: -1 }
      }
      if (row > 0 && Math.abs(cy - geom.rowTop(row)) <= BORDER_TOLERANCE) {
        return { region: 'rowborder', row: row - 1, col: -1 }
      }
      return { region: 'rowheader', row, col: -1 }
    }
    return {
      region: 'cell',
      row: geom.rowAt(y - COL_HEADER_HEIGHT + this.scrollY),
      col: geom.colAt(x - ROW_HEADER_WIDTH + this.scrollX),
    }
  }

  cellViewportRect(row: number, col: number): Rect {
    const r = this.geometry().cellRect(row, col)
    return {
      x: ROW_HEADER_WIDTH + r.x - this.scrollX,
      y: COL_HEADER_HEIGHT + r.y - this.scrollY,
      w: r.w,
      h: r.h,
    }
  }

  ensureVisible(addr: CellAddr): void {
    const geom = this.geometry()
    const r = geom.cellRect(addr.row, addr.col)
    const viewW = this.stage.width() - ROW_HEADER_WIDTH
    const viewH = this.stage.height() - COL_HEADER_HEIGHT
    if (r.x < this.scrollX) this.scrollX = r.x
    else if (r.x + r.w > this.scrollX + viewW) this.scrollX = r.x + r.w - viewW
    if (r.y < this.scrollY) this.scrollY = r.y
    else if (r.y + r.h > this.scrollY + viewH) this.scrollY = r.y + r.h - viewH
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
    return Math.max(0, this.geometry().contentWidth - (this.stage.width() - ROW_HEADER_WIDTH))
  }

  private maxScrollY(): number {
    return Math.max(0, this.geometry().contentHeight - (this.stage.height() - COL_HEADER_HEIGHT))
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
    if (hit.region === 'cell') openEditor(this, { row: hit.row, col: hit.col })
  }

  private onMouseMove = (e: MouseEvent): void => {
    const hit = this.hitTest(e.clientX, e.clientY)
    this.someProp('handleMouseMove', (p) => p(this, e, hit))
  }

  private onMouseUp = (e: MouseEvent): void => {
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
