// EditorView：命令式 Konva 视图。事件流：DOM 事件 → hitTest → 插件 someProp 链 →
// 未拦截走默认行为 → dispatch(tr) → applyTransaction → updateState → rAF 合帧重绘。
import Konva from 'konva'
import { CellAddr } from '@gmi/excel-core'
import { COL_HEADER_HEIGHT, ROW_HEADER_WIDTH, SEL_BORDER_HIT, SheetData } from '@gmi/excel-core'
import { EditorViewLike, PluginProps, PluginView } from '@gmi/excel-core'
import { rangeSelection, selectionRange, singleCell } from '@gmi/excel-core'
import type { SheetState } from '@gmi/excel-core'
import type { Transaction } from '@gmi/excel-core'
import { evaluatorFor } from '@gmi/excel-core'
import { filterHiddenRows } from '@gmi/excel-core'
import { autoRowHeights } from './autoheight'
import { openEditor } from './editbox'
import { GridGeometry } from './geometry'
import { fillHandleSize, renderAll } from './layers'
import { contentViewport, hScrollbar, SB_SIZE, thumbHit, vScrollbar } from './scrollbar'
import { HitResult, Rect, contextMenuKey, zoomKey } from './types'
import { anchoredScroll, nextZoomLevel, zoomOf } from './zoom'

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
  private geomCache: { sheet: SheetData; zoom: number; geom: GridGeometry } | null = null
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

  // 当前活动表的缩放档位（缺省 1）
  zoom(): number {
    return zoomOf(this.state, this.state.doc.active)
  }

  // 按 activeSheet 引用 + zoom 缓存几何对象（行高列宽/冻结/筛选/隐藏变化都会得到新 SheetData）
  geometry(): GridGeometry {
    const sheet = this.state.activeSheet
    const z = this.zoom()
    if (this.geomCache?.sheet !== sheet || this.geomCache.zoom !== z) {
      const ev = evaluatorFor(this.state.doc)
      const extra = sheet.filter ? filterHiddenRows(this.state.doc.active, sheet, ev) : undefined
      const auto = autoRowHeights(sheet, this.state.doc.active, ev)
      this.geomCache = {
        sheet,
        zoom: z,
        geom: new GridGeometry(sheet, sheet.frozenRows, sheet.frozenCols, extra, auto.size ? auto : undefined, z),
      }
    }
    return this.geomCache.geom
  }

  // 应用事务的唯一入口：applyTransaction → updateState → 插件 view.update → 重绘。
  // 宿主（React bridge 等）经 subscribe 感知 state 变化，不做事务级回调
  dispatch(tr: Transaction): void {
    const { state } = this.state.applyTransaction(tr)
    this.updateState(state)
    if (tr.scrolledIntoView) this.ensureVisible(state.selection.activeCell)
  }

  updateState(state: SheetState): void {
    const prev = this.state
    this.state = state
    // 切换活动表：视口归零，避免沿用前表的滚动偏移
    if (prev.doc.active !== state.doc.active) {
      this.scrollX = 0
      this.scrollY = 0
    }
    // zoom 变化（滚轮/状态栏菜单等一切路径）后钳位滚动，避免超出新上限留白
    if (zoomOf(prev, prev.doc.active) !== zoomOf(state, state.doc.active)) {
      this.clampScroll()
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

  // 内容视口与双条存在性：纵条存在 ⇔ vp.w < w0（纵条占宽），横条存在 ⇔ vp.h < h0
  private viewportWithBars(geom: GridGeometry): { vp: { w: number; h: number }; hVisible: boolean; vVisible: boolean } {
    const z = this.zoom()
    const w0 = this.stage.width() - ROW_HEADER_WIDTH * z
    const h0 = this.stage.height() - COL_HEADER_HEIGHT * z
    const vp = contentViewport(geom.contentWidth, geom.contentHeight, w0, h0, SB_SIZE * z)
    return { vp, hVisible: vp.h < h0, vVisible: vp.w < w0 }
  }

  hitTest(clientX: number, clientY: number): HitResult {
    const rect = this.dom.getBoundingClientRect()
    const x = clientX - rect.left
    const y = clientY - rect.top
    if (x < 0 || y < 0 || x > this.stage.width() || y > this.stage.height()) {
      return { region: 'outside', row: -1, col: -1 }
    }
    // 表头尺寸/滚动条厚度随 zoom 缩放
    const z = this.zoom()
    const hw = ROW_HEADER_WIDTH * z
    const hh = COL_HEADER_HEIGHT * z
    const sb = SB_SIZE * z
    // 滚动条区域优先于其他命中
    const geom0 = this.geometry()
    const { hVisible, vVisible } = this.viewportWithBars(geom0)
    const vg = vScrollbar(geom0.contentHeight, this.scrollY, this.stage.width(), this.stage.height(), sb, hh, hVisible)
    if (vg && x >= vg.track.x && y >= vg.track.y && y <= vg.track.y + vg.track.h) {
      return { region: 'vscrollbar', row: -1, col: -1 }
    }
    const hg = hScrollbar(geom0.contentWidth, this.scrollX, this.stage.width(), this.stage.height(), sb, hw, vVisible)
    if (hg && y >= hg.track.y && x >= hg.track.x && x <= hg.track.x + hg.track.w) {
      return { region: 'hscrollbar', row: -1, col: -1 }
    }
    const geom = this.geometry()
    // 填充手柄：活动选区右下角 6px 方块（×zoom）±2px 容差（用活动格的视口位置）
    const sel = this.state.selection
    const sRange = selectionRange(sel)
    const br = this.cellViewportRect(sRange.er, sRange.ec)
    const hx = br.x + br.w
    const hy = br.y + br.h
    const half = fillHandleSize(z) / 2 + 2
    if (Math.abs(x - hx) <= half && Math.abs(y - hy) <= half) {
      return { region: 'fillhandle', row: sel.activeCell.row, col: sel.activeCell.col }
    }
    // 选区边框：活动区域四边 ±SEL_BORDER_HIT 像素带（不含内部、不含右下角填充柄区）
    const tl = this.cellViewportRect(sRange.sr, sRange.sc)
    const sx = tl.x, sy = tl.y, sex = br.x + br.w, sey = br.y + br.h
    const edge = SEL_BORDER_HIT * z
    const onEdge =
      (Math.abs(x - sx) <= edge || Math.abs(x - sex) <= edge) && y >= sy - edge && y <= sey + edge ||
      (Math.abs(y - sy) <= edge || Math.abs(y - sey) <= edge) && x >= sx - edge && x <= sex + edge
    if (onEdge && !(x > sex - half && y > sey - half)) return { region: 'selborder', row: sel.activeCell.row, col: sel.activeCell.col }
    if (x < hw && y < hh) return { region: 'corner', row: -1, col: -1 }
    if (y < hh) {
      const cx = x - hw
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
    if (x < hw) {
      const cy = y - hh
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
    const a = geom.cellAtContent(x - hw, y - hh, this.scrollX, this.scrollY)
    // 筛选箭头：表头行单元格右缘 18px（×zoom）区域，与绘制偏移一致
    const f = this.state.activeSheet.filter
    if (f && a.row === f.range.sr && a.col >= f.range.sc && a.col <= f.range.ec) {
      const rect = this.cellViewportRect(a.row, a.col)
      if (x >= rect.x + rect.w - 18 * z) return { region: 'filter', row: a.row, col: a.col }
    }
    return { region: 'cell', row: a.row, col: a.col }
  }

  // 指针 client 坐标 → 冻结感知 + clamp 到表内的单元格地址（供插件与 hitTest 共用）
  pointerToCell(clientX: number, clientY: number): CellAddr {
    const rect = this.dom.getBoundingClientRect()
    const geom = this.geometry()
    const sheet = this.state.activeSheet
    const a = geom.cellAtContent(
      clientX - rect.left - ROW_HEADER_WIDTH * this.zoom(),
      clientY - rect.top - COL_HEADER_HEIGHT * this.zoom(),
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
    const z = this.zoom()
    return { x: ROW_HEADER_WIDTH * z + sx, y: COL_HEADER_HEIGHT * z + sy, w: r.w, h: r.h }
  }

  ensureVisible(addr: CellAddr): void {
    const geom = this.geometry()
    const z = this.zoom()
    const vp = contentViewport(
      geom.contentWidth,
      geom.contentHeight,
      this.stage.width() - ROW_HEADER_WIDTH * z,
      this.stage.height() - COL_HEADER_HEIGHT * z,
      SB_SIZE * z,
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
      this.zoom(),
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
    const z = this.zoom()
    const vp = contentViewport(
      this.geometry().contentWidth,
      this.geometry().contentHeight,
      this.stage.width() - ROW_HEADER_WIDTH * z,
      this.stage.height() - COL_HEADER_HEIGHT * z,
      SB_SIZE * z,
    )
    return Math.max(0, this.geometry().contentWidth - vp.w)
  }

  private maxScrollY(): number {
    const z = this.zoom()
    const vp = contentViewport(
      this.geometry().contentWidth,
      this.geometry().contentHeight,
      this.stage.width() - ROW_HEADER_WIDTH * z,
      this.stage.height() - COL_HEADER_HEIGHT * z,
      SB_SIZE * z,
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
    content.addEventListener('contextmenu', this.onContextMenu)
    content.addEventListener('wheel', this.onWheel, { passive: false })
    window.addEventListener('mousemove', this.onMouseMove)
    window.addEventListener('mouseup', this.onMouseUp)
    this.proxy.addEventListener('keydown', this.onKeyDown)
    this.proxy.addEventListener('paste', this.onPaste)
    this.proxy.addEventListener('copy', this.onCopy)
    this.proxy.addEventListener('cut', this.onCut)
  }

  private onMouseDown = (e: MouseEvent): void => {
    // 右键 mousedown 不做任何选择/拖拽：选区调整由 contextmenu 处理器负责（点在选区内保持选区），
    // 且右键菜单会拦截 mouseup 冒泡，此处若启动拖拽态将永不清除（issue #1）
    if (e.button !== 0) return
    const hit = this.hitTest(e.clientX, e.clientY)
    if (hit.region === 'hscrollbar' || hit.region === 'vscrollbar') {
      const geom = this.geometry()
      const z = this.zoom()
      const hw = ROW_HEADER_WIDTH * z
      const hh = COL_HEADER_HEIGHT * z
      const sb = SB_SIZE * z
      const { vp, hVisible, vVisible } = this.viewportWithBars(geom)
      const g =
        hit.region === 'vscrollbar'
          ? vScrollbar(geom.contentHeight, this.scrollY, this.stage.width(), this.stage.height(), sb, hh, hVisible)
          : hScrollbar(geom.contentWidth, this.scrollX, this.stage.width(), this.stage.height(), sb, hw, vVisible)
      if (!g) return
      const rect = this.dom.getBoundingClientRect()
      const x = e.clientX - rect.left
      const y = e.clientY - rect.top
      const axis = hit.region === 'vscrollbar' ? 'v' : 'h'
      if (thumbHit(g, x, y)) {
        this.sbDrag = { axis, startScroll: axis === 'v' ? this.scrollY : this.scrollX, startPos: axis === 'v' ? y : x, ratio: g.ratio }
      } else {
        // 点轨道翻页（向点击方向滚一个内容视口，自动扣对侧条厚度）
        const page = axis === 'v' ? vp.h : vp.w
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

  // 右键：点在选区外先改选区（整行/列头同理），再开菜单（不入 undo）
  private onContextMenu = (e: MouseEvent): void => {
    e.preventDefault()
    const hit = this.hitTest(e.clientX, e.clientY)
    const st = this.state
    const sheet = st.activeSheet
    const sel = selectionRange(st.selection)
    const tr = st.tr
    const openMenu = (kind: 'cell' | 'rowheader' | 'colheader', row: number, col: number): void => {
      tr.setMeta(contextMenuKey, { kind, x: e.clientX, y: e.clientY, row, col }).setMeta('addToHistory', false)
      this.dispatch(tr)
    }
    switch (hit.region) {
      case 'cell':
      case 'filter': {
        const inSel = st.selection.ranges.some(r => hit.row >= r.sr && hit.row <= r.er && hit.col >= r.sc && hit.col <= r.ec)
        if (!inSel) tr.setSelection(singleCell(hit.row, hit.col))
        openMenu('cell', hit.row, hit.col)
        break
      }
      case 'rowheader': {
        const fullRows = sel.sc === 0 && sel.ec === sheet.colCount - 1
        if (!(fullRows && hit.row >= sel.sr && hit.row <= sel.er)) {
          tr.setSelection(rangeSelection({ sr: hit.row, sc: 0, er: hit.row, ec: sheet.colCount - 1 }))
        }
        openMenu('rowheader', hit.row, -1)
        break
      }
      case 'colheader': {
        const fullCols = sel.sr === 0 && sel.er === sheet.rowCount - 1
        if (!(fullCols && hit.col >= sel.sc && hit.col <= sel.ec)) {
          tr.setSelection(rangeSelection({ sr: 0, sc: hit.col, er: sheet.rowCount - 1, ec: hit.col }))
        }
        openMenu('colheader', -1, hit.col)
        break
      }
      default:
        break // 滚动条/角落/外部不弹
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
    // Ctrl/Cmd+滚轮：以光标为锚点按档缩放（zoom 非文档态，不入 undo）
    if (e.ctrlKey || e.metaKey) {
      const st = this.state
      const z0 = this.zoom()
      const z1 = nextZoomLevel(z0, e.deltaY < 0 ? 1 : -1)
      if (z1 === z0) return
      const rect = this.dom.getBoundingClientRect()
      const hw = ROW_HEADER_WIDTH * z0
      const hh = COL_HEADER_HEIGHT * z0
      const cx = e.clientX - rect.left - hw
      const cy = e.clientY - rect.top - hh
      this.scrollX = anchoredScroll(this.scrollX, cx, z0, z1)
      this.scrollY = anchoredScroll(this.scrollY, cy, z0, z1)
      const field = (st.getField(zoomKey) as Record<string, number> | null) ?? {}
      this.dispatch(
        st.tr.setMeta(zoomKey, { ...field, [st.doc.active]: z1 }).setMeta('addToHistory', false),
      )
      this.clampScroll()
      this.render()
      return
    }
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
