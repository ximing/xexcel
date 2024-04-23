// 三层渲染：gridLayer 背景+表头；cellLayer 单元格 bg+网格线+文字（含合并区）；
// overlayLayer 选区+填充手柄+参考线。
// 冻结行列：内容区按 computeQuadrants 切成 ≤4 象限（交叉/冻结行/冻结列/主区），
// 每象限独立 clip+偏移复用同一套内容函数；选区覆盖层同样按象限裁切（自然分裂）。
// 每次 updateState 后整层重建可见节点（M1 从简，1000 格/帧量级可接受）。
import Konva from 'konva'
import { CellRange, colName, rangesIntersect } from '../core/addr'
import { BorderEdge, COL_HEADER_HEIGHT, ROW_HEADER_WIDTH } from '../core/model'
import { edgeDash, edgeWidth, resolveHEdge, resolveVEdge } from './borders'
import { selectionRange } from '../core/selection'
import type { SheetState } from '../core/state'
import type { CellEvaluator } from '../formula/engine'
import { evaluatorFor } from '../formula/engine'
import { GridGeometry } from './geometry'
import { CELL_PAD_X } from './measure'
import { hScrollbar, vScrollbar } from './scrollbar'
import { fillPreviewKey, ResizeGuide, resizeGuideKey } from './types'

const COLOR_GRID = '#d9dce1'
const COLOR_HEADER_BG = '#f7f8fa'
const COLOR_HEADER_TEXT = '#5f6368'
const COLOR_HEADER_ACTIVE_BG = '#e8f0fe'
const COLOR_HEADER_ACTIVE_TEXT = '#1a73e8'
const COLOR_SELECT_FILL = 'rgba(26, 115, 232, 0.12)'
const COLOR_SELECT_BORDER = '#1a73e8'
const COLOR_FROZEN_LINE = '#9aa0a6'
const COLOR_TEXT = '#202124'
const FONT_SIZE = 13
const FONT_FAMILY =
  '-apple-system, BlinkMacSystemFont, "Segoe UI", Roboto, "Helvetica Neue", Arial, sans-serif'
export const FILL_HANDLE_SIZE = 6

const noListen = { listening: false }

// 单条边框边：double 画两条平行 1px 线，其余按线宽/虚线
function drawBorderEdge(inner: Konva.Group, points: number[], e: BorderEdge): void {
  const stroke = e.color ?? '#000000'
  if (e.style === 'double') {
    const [x1, y1, x2, y2] = points
    const v = x1 === x2 // 竖线
    inner.add(new Konva.Line({ points: v ? [x1 - 1, y1, x2 - 1, y2] : [x1, y1 - 1, x2, y2 - 1], stroke, strokeWidth: 1, ...noListen }))
    inner.add(new Konva.Line({ points: v ? [x1 + 1, y1, x2 + 1, y2] : [x1, y1 + 1, x2, y2 + 1], stroke, strokeWidth: 1, ...noListen }))
    return
  }
  inner.add(new Konva.Line({ points, stroke, strokeWidth: edgeWidth(e.style), dash: edgeDash(e.style), ...noListen }))
}

// 一个渲染象限：clip 为内容区坐标（不含表头），off 为内容坐标的绘制偏移
interface Quadrant {
  clipX: number
  clipY: number
  clipW: number
  clipH: number
  offX: number
  offY: number
  sr: number
  sc: number
  er: number
  ec: number
}

// 主区（可滚区）可见范围
function scrollableRange(
  geom: GridGeometry,
  scrollX: number,
  scrollY: number,
  viewW: number,
  viewH: number,
): CellRange {
  const W = Math.max(0, viewW - ROW_HEADER_WIDTH)
  const H = Math.max(0, viewH - COL_HEADER_HEIGHT)
  const fw = geom.frozenWidth
  const fh = geom.frozenHeight
  const sr = Math.max(geom.frozenRows, geom.rowAt(fh + scrollY))
  const er = Math.max(sr, geom.rowAt(fh + scrollY + Math.max(0, H - fh) - 1))
  const sc = Math.max(geom.frozenCols, geom.colAt(fw + scrollX))
  const ec = Math.max(sc, geom.colAt(fw + scrollX + Math.max(0, W - fw) - 1))
  return { sr, sc, er, ec }
}

function computeQuadrants(
  geom: GridGeometry,
  scrollX: number,
  scrollY: number,
  viewW: number,
  viewH: number,
): Quadrant[] {
  const W = Math.max(0, viewW - ROW_HEADER_WIDTH)
  const H = Math.max(0, viewH - COL_HEADER_HEIGHT)
  const fw = geom.frozenWidth
  const fh = geom.frozenHeight
  const fr = geom.frozenRows
  const fc = geom.frozenCols
  const main = scrollableRange(geom, scrollX, scrollY, viewW, viewH)
  const qs: Quadrant[] = []
  if (W > fw && H > fh) {
    qs.push({ clipX: fw, clipY: fh, clipW: W - fw, clipH: H - fh, offX: -fw - scrollX, offY: -fh - scrollY, ...main })
  }
  if (fr > 0 && W > fw) {
    qs.push({ clipX: fw, clipY: 0, clipW: W - fw, clipH: fh, offX: -fw - scrollX, offY: 0, sr: 0, sc: main.sc, er: fr - 1, ec: main.ec })
  }
  if (fc > 0 && H > fh) {
    qs.push({ clipX: 0, clipY: fh, clipW: fw, clipH: H - fh, offX: 0, offY: -fh - scrollY, sr: main.sr, sc: 0, er: main.er, ec: fc - 1 })
  }
  if (fr > 0 && fc > 0) {
    qs.push({ clipX: 0, clipY: 0, clipW: fw, clipH: fh, offX: 0, offY: 0, sr: 0, sc: 0, er: fr - 1, ec: fc - 1 })
  }
  return qs
}

function quadrantGroup(q: Quadrant): Konva.Group {
  const g = new Konva.Group({
    x: ROW_HEADER_WIDTH + q.clipX,
    y: COL_HEADER_HEIGHT + q.clipY,
    clipX: 0,
    clipY: 0,
    clipWidth: Math.max(0, q.clipW),
    clipHeight: Math.max(0, q.clipH),
    ...noListen,
  })
  g.add(new Konva.Group({ x: q.offX, y: q.offY, ...noListen }))
  return g
}

export function renderAll(
  gridLayer: Konva.Layer,
  cellLayer: Konva.Layer,
  overlayLayer: Konva.Layer,
  state: SheetState,
  geom: GridGeometry,
  scrollX: number,
  scrollY: number,
  viewW: number,
  viewH: number,
): void {
  gridLayer.destroyChildren()
  cellLayer.destroyChildren()
  overlayLayer.destroyChildren()
  if (viewW <= 0 || viewH <= 0) return
  renderCellLayer(cellLayer, state, geom, scrollX, scrollY, viewW, viewH)
  renderGridLayer(gridLayer, state, geom, scrollX, scrollY, viewW, viewH)
  renderOverlayLayer(overlayLayer, state, geom, scrollX, scrollY, viewW, viewH)
  renderScrollbars(overlayLayer, geom, scrollX, scrollY, viewW, viewH)
  gridLayer.batchDraw()
  cellLayer.batchDraw()
  overlayLayer.batchDraw()
}

function renderGridLayer(
  layer: Konva.Layer,
  state: SheetState,
  geom: GridGeometry,
  scrollX: number,
  scrollY: number,
  viewW: number,
  viewH: number,
): void {
  layer.add(new Konva.Rect({ x: ROW_HEADER_WIDTH, y: COL_HEADER_HEIGHT, width: viewW, height: viewH, fill: '#ffffff', ...noListen }))
  layer.add(new Konva.Rect({ x: 0, y: 0, width: viewW, height: COL_HEADER_HEIGHT, fill: COLOR_HEADER_BG, ...noListen }))
  layer.add(new Konva.Rect({ x: 0, y: 0, width: ROW_HEADER_WIDTH, height: viewH, fill: COLOR_HEADER_BG, ...noListen }))

  const sel = selectionRange(state.selection)
  const main = scrollableRange(geom, scrollX, scrollY, viewW, viewH)
  const fw = geom.frozenWidth
  const fh = geom.frozenHeight
  const sheet = geom.sheet

  // 列头：冻结段（不滚）+ 可滚段（clip 到冻结线右侧）
  const drawColHeader = (c: number, x: number): void => {
    const w = geom.colWidth(c)
    const active = c >= sel.sc && c <= sel.ec
    if (active) {
      layer.add(new Konva.Rect({ x, y: 0, width: w, height: COL_HEADER_HEIGHT, fill: COLOR_HEADER_ACTIVE_BG, ...noListen }))
    }
    layer.add(
      new Konva.Text({
        x, y: 0, width: w, height: COL_HEADER_HEIGHT,
        text: colName(c), align: 'center', verticalAlign: 'middle',
        fontSize: FONT_SIZE, fontFamily: FONT_FAMILY,
        fill: active ? COLOR_HEADER_ACTIVE_TEXT : COLOR_HEADER_TEXT,
        wrap: 'none', ...noListen,
      }),
    )
    layer.add(new Konva.Line({ points: [x, 0, x, COL_HEADER_HEIGHT], stroke: COLOR_GRID, strokeWidth: 1, ...noListen }))
  }
  for (let c = 0; c < geom.frozenCols; c++) drawColHeader(c, ROW_HEADER_WIDTH + geom.colLeft(c))
  const colStrip = new Konva.Group({
    x: 0, y: 0,
    clipX: ROW_HEADER_WIDTH + fw, clipY: 0,
    clipWidth: Math.max(0, viewW - ROW_HEADER_WIDTH - fw), clipHeight: COL_HEADER_HEIGHT,
    ...noListen,
  })
  for (let c = main.sc; c <= main.ec; c++) {
    drawColHeaderInto(colStrip, c, ROW_HEADER_WIDTH + fw + (geom.colLeft(c) - fw - scrollX), geom, sel)
  }
  layer.add(colStrip)

  // 行头：冻结段 + 可滚段（clip 到冻结线下侧）
  const drawRowHeader = (r: number, y: number): void => {
    const h = geom.rowHeight(r)
    const active = r >= sel.sr && r <= sel.er
    if (active) {
      layer.add(new Konva.Rect({ x: 0, y, width: ROW_HEADER_WIDTH, height: h, fill: COLOR_HEADER_ACTIVE_BG, ...noListen }))
    }
    layer.add(
      new Konva.Text({
        x: 0, y, width: ROW_HEADER_WIDTH, height: h,
        text: String(r + 1), align: 'center', verticalAlign: 'middle',
        fontSize: FONT_SIZE, fontFamily: FONT_FAMILY,
        fill: active ? COLOR_HEADER_ACTIVE_TEXT : COLOR_HEADER_TEXT,
        wrap: 'none', ...noListen,
      }),
    )
    layer.add(new Konva.Line({ points: [0, y, ROW_HEADER_WIDTH, y], stroke: COLOR_GRID, strokeWidth: 1, ...noListen }))
  }
  for (let r = 0; r < geom.frozenRows; r++) drawRowHeader(r, COL_HEADER_HEIGHT + geom.rowTop(r))
  const rowStrip = new Konva.Group({
    x: 0, y: 0,
    clipX: 0, clipY: COL_HEADER_HEIGHT + fh,
    clipWidth: ROW_HEADER_WIDTH, clipHeight: Math.max(0, viewH - COL_HEADER_HEIGHT - fh),
    ...noListen,
  })
  for (let r = main.sr; r <= main.er; r++) {
    drawRowHeaderInto(rowStrip, r, COL_HEADER_HEIGHT + fh + (geom.rowTop(r) - fh - scrollY), geom, sel)
  }
  layer.add(rowStrip)

  // 表头带外缘分隔线
  layer.add(new Konva.Line({ points: [0, COL_HEADER_HEIGHT, viewW, COL_HEADER_HEIGHT], stroke: COLOR_GRID, strokeWidth: 1, ...noListen }))
  layer.add(new Konva.Line({ points: [ROW_HEADER_WIDTH, 0, ROW_HEADER_WIDTH, viewH], stroke: COLOR_GRID, strokeWidth: 1, ...noListen }))
  // 冻结分割线
  if (geom.frozenCols > 0) {
    const x = ROW_HEADER_WIDTH + fw
    layer.add(new Konva.Line({ points: [x, 0, x, viewH], stroke: COLOR_FROZEN_LINE, strokeWidth: 2, ...noListen }))
  }
  if (geom.frozenRows > 0) {
    const y = COL_HEADER_HEIGHT + fh
    layer.add(new Konva.Line({ points: [0, y, viewW, y], stroke: COLOR_FROZEN_LINE, strokeWidth: 2, ...noListen }))
  }
  // 隐藏行列提示：隐藏区塌缩处的表头画双线（按屏幕坐标去重，一段连续隐藏只画一次）
  const rowMarks = new Set<number>()
  for (const r of sheet.hiddenRows) {
    const y =
      r < geom.frozenRows
        ? COL_HEADER_HEIGHT + geom.rowTop(r)
        : COL_HEADER_HEIGHT + fh + (geom.rowTop(r) - fh - scrollY)
    if (y > COL_HEADER_HEIGHT && y < viewH) rowMarks.add(Math.round(y))
  }
  for (const y of rowMarks) {
    layer.add(new Konva.Line({ points: [0, y - 2, ROW_HEADER_WIDTH, y - 2], stroke: COLOR_FROZEN_LINE, strokeWidth: 1, ...noListen }))
    layer.add(new Konva.Line({ points: [0, y + 2, ROW_HEADER_WIDTH, y + 2], stroke: COLOR_FROZEN_LINE, strokeWidth: 1, ...noListen }))
  }
  const colMarks = new Set<number>()
  for (const c of sheet.hiddenCols) {
    const x =
      c < geom.frozenCols
        ? ROW_HEADER_WIDTH + geom.colLeft(c)
        : ROW_HEADER_WIDTH + fw + (geom.colLeft(c) - fw - scrollX)
    if (x > ROW_HEADER_WIDTH && x < viewW) colMarks.add(Math.round(x))
  }
  for (const x of colMarks) {
    layer.add(new Konva.Line({ points: [x - 2, 0, x - 2, COL_HEADER_HEIGHT], stroke: COLOR_FROZEN_LINE, strokeWidth: 1, ...noListen }))
    layer.add(new Konva.Line({ points: [x + 2, 0, x + 2, COL_HEADER_HEIGHT], stroke: COLOR_FROZEN_LINE, strokeWidth: 1, ...noListen }))
  }
}

// 与 drawColHeader/drawRowHeader 同逻辑，绘制到指定容器（供 clip strip 用）
function drawColHeaderInto(
  g: Konva.Group | Konva.Layer,
  c: number,
  x: number,
  geom: GridGeometry,
  sel: CellRange,
): void {
  const w = geom.colWidth(c)
  const active = c >= sel.sc && c <= sel.ec
  if (active) {
    g.add(new Konva.Rect({ x, y: 0, width: w, height: COL_HEADER_HEIGHT, fill: COLOR_HEADER_ACTIVE_BG, ...noListen }))
  }
  g.add(
    new Konva.Text({
      x, y: 0, width: w, height: COL_HEADER_HEIGHT,
      text: colName(c), align: 'center', verticalAlign: 'middle',
      fontSize: FONT_SIZE, fontFamily: FONT_FAMILY,
      fill: active ? COLOR_HEADER_ACTIVE_TEXT : COLOR_HEADER_TEXT,
      wrap: 'none', ...noListen,
    }),
  )
  g.add(new Konva.Line({ points: [x, 0, x, COL_HEADER_HEIGHT], stroke: COLOR_GRID, strokeWidth: 1, ...noListen }))
}

function drawRowHeaderInto(
  g: Konva.Group | Konva.Layer,
  r: number,
  y: number,
  geom: GridGeometry,
  sel: CellRange,
): void {
  const h = geom.rowHeight(r)
  const active = r >= sel.sr && r <= sel.er
  if (active) {
    g.add(new Konva.Rect({ x: 0, y, width: ROW_HEADER_WIDTH, height: h, fill: COLOR_HEADER_ACTIVE_BG, ...noListen }))
  }
  g.add(
    new Konva.Text({
      x: 0, y, width: ROW_HEADER_WIDTH, height: h,
      text: String(r + 1), align: 'center', verticalAlign: 'middle',
      fontSize: FONT_SIZE, fontFamily: FONT_FAMILY,
      fill: active ? COLOR_HEADER_ACTIVE_TEXT : COLOR_HEADER_TEXT,
      wrap: 'none', ...noListen,
    }),
  )
  g.add(new Konva.Line({ points: [0, y, ROW_HEADER_WIDTH, y], stroke: COLOR_GRID, strokeWidth: 1, ...noListen }))
}

function renderCellLayer(
  layer: Konva.Layer,
  state: SheetState,
  geom: GridGeometry,
  scrollX: number,
  scrollY: number,
  viewW: number,
  viewH: number,
): void {
  const evaluator = evaluatorFor(state.doc)
  for (const q of computeQuadrants(geom, scrollX, scrollY, viewW, viewH)) {
    const clip = quadrantGroup(q)
    renderCellsInto(clip.children[0] as Konva.Group, state, geom, evaluator, q)
    layer.add(clip)
  }
}

// 象限内单元格内容：bg → 合并区 → 网格线 → 文字（范围由象限给定）
function renderCellsInto(
  inner: Konva.Group,
  state: SheetState,
  geom: GridGeometry,
  evaluator: CellEvaluator,
  q: { sr: number; sc: number; er: number; ec: number },
): void {
  const sheet = state.activeSheet
  const sheetId = state.doc.active
  const merges = sheet.merges.filter((m) => rangesIntersect(m, q))

  const left = geom.colLeft(q.sc)
  const right = geom.colLeft(q.ec) + geom.colWidth(q.ec)
  const top = geom.rowTop(q.sr)
  const bottom = geom.rowTop(q.er) + geom.rowHeight(q.er)

  // bg（合并区内的格跳过，由合并区统一画）
  for (let r = q.sr; r <= q.er; r++) {
    for (let c = q.sc; c <= q.ec; c++) {
      if (merges.some((m) => r >= m.sr && r <= m.er && c >= m.sc && c <= m.ec)) continue
      const cell = sheet.getCell(r, c)
      if (!cell?.style?.bg) continue
      const rect = geom.cellRect(r, c)
      inner.add(new Konva.Rect({ x: rect.x, y: rect.y, width: rect.w, height: rect.h, fill: cell.style.bg, ...noListen }))
    }
  }
  // 网格线
  for (let c = q.sc; c <= q.ec + 1; c++) {
    const x = geom.colLeft(c)
    inner.add(new Konva.Line({ points: [x, top, x, bottom], stroke: COLOR_GRID, strokeWidth: 1, ...noListen }))
  }
  for (let r = q.sr; r <= q.er + 1; r++) {
    const y = geom.rowTop(r)
    inner.add(new Konva.Line({ points: [left, y, right, y], stroke: COLOR_GRID, strokeWidth: 1, ...noListen }))
  }
  // 边框：每条网格线按共享边权重裁决后画一次（覆盖默认网格线）
  for (let r = q.sr; r <= q.er; r++) {
    for (let c = q.sc; c <= q.ec + 1; c++) {
      const e = resolveVEdge(sheet, r, c)
      if (!e) continue
      const x = geom.colLeft(c)
      drawBorderEdge(inner, [x, geom.rowTop(r), x, geom.rowTop(r) + geom.rowHeight(r)], e)
    }
  }
  for (let r = q.sr; r <= q.er + 1; r++) {
    for (let c = q.sc; c <= q.ec; c++) {
      const e = resolveHEdge(sheet, r, c)
      if (!e) continue
      const y = geom.rowTop(r)
      drawBorderEdge(inner, [geom.colLeft(c), y, geom.colLeft(c) + geom.colWidth(c), y], e)
    }
  }
  // 合并区：白底盖网格线 → 锚点 bg → 外框 → 锚点文字
  for (const m of merges) {
    const rect = geom.rangeRect(m)
    inner.add(new Konva.Rect({ x: rect.x, y: rect.y, width: rect.w, height: rect.h, fill: '#ffffff', ...noListen }))
    const anchor = sheet.getCell(m.sr, m.sc)
    if (anchor?.style?.bg) {
      inner.add(new Konva.Rect({ x: rect.x, y: rect.y, width: rect.w, height: rect.h, fill: anchor.style.bg, ...noListen }))
    }
    inner.add(new Konva.Rect({ x: rect.x, y: rect.y, width: rect.w, height: rect.h, stroke: COLOR_GRID, strokeWidth: 1, ...noListen }))
    if (anchor && anchor.raw !== '') drawCellText(inner, evaluator, sheetId, m.sr, m.sc, anchor, geom.rangeRect(m))
  }
  // 文字（合并区内的格跳过）
  for (let r = q.sr; r <= q.er; r++) {
    for (let c = q.sc; c <= q.ec; c++) {
      if (merges.some((m) => r >= m.sr && r <= m.er && c >= m.sc && c <= m.ec)) continue
      const cell = sheet.getCell(r, c)
      if (!cell || cell.raw === '') continue
      drawCellText(inner, evaluator, sheetId, r, c, cell, geom.cellRect(r, c))
    }
  }
  // 筛选箭头：筛选区域表头行各列右缘小三角；有生效 criteria 的列高亮
  const f = sheet.filter
  if (f && f.range.sr >= q.sr && f.range.sr <= q.er) {
    for (let c = Math.max(f.range.sc, q.sc); c <= Math.min(f.range.ec, q.ec); c++) {
      const rect = geom.cellRect(f.range.sr, c)
      const cx = rect.x + rect.w - 13
      const cy = rect.y + rect.h / 2
      const crit = f.criteria[c]
      const on = crit !== undefined && (crit.type !== 'values' || crit.excluded.length > 0)
      inner.add(
        new Konva.Line({
          points: [cx - 4, cy - 3, cx + 4, cy - 3, cx, cy + 3],
          closed: true,
          fill: on ? '#1a73e8' : '#9aa0a6',
          ...noListen,
        }),
      )
    }
  }
}

function drawCellText(
  inner: Konva.Group,
  evaluator: CellEvaluator,
  sheetId: string,
  r: number,
  c: number,
  cell: { raw: string; style?: import('../core/model').CellStyle },
  rect: { x: number; y: number; w: number; h: number },
): void {
  const text = evaluator.displayText(sheetId, r, c)
  if (text === '') return
  let align = cell.style?.align
  if (!align) {
    const v = evaluator.get(sheetId, r, c)
    align = typeof v === 'number' || typeof v === 'boolean' ? 'right' : 'left'
  }
  const fontStyle =
    cell.style?.bold && cell.style?.italic
      ? 'bold italic'
      : cell.style?.bold
        ? 'bold'
        : cell.style?.italic
          ? 'italic'
          : 'normal'
  const g = new Konva.Group({ clipX: rect.x, clipY: rect.y, clipWidth: rect.w, clipHeight: rect.h, ...noListen })
  g.add(
    new Konva.Text({
      x: rect.x + CELL_PAD_X,
      y: rect.y,
      width: Math.max(0, rect.w - CELL_PAD_X * 2),
      height: rect.h,
      text,
      align,
      verticalAlign: cell.style?.vAlign ?? 'bottom',
      fontSize: cell.style?.fontSize ?? FONT_SIZE,
      fontFamily: cell.style?.fontFamily ?? FONT_FAMILY,
      fontStyle,
      textDecoration: cell.style?.underline
        ? cell.style?.strikethrough
          ? 'underline line-through'
          : 'underline'
        : cell.style?.strikethrough
          ? 'line-through'
          : '',
      fill: cell.style?.color ?? COLOR_TEXT,
      wrap: cell.style?.wrap ? 'char' : 'none',
      ...noListen,
    }),
  )
  inner.add(g)
}

function renderOverlayLayer(
  layer: Konva.Layer,
  state: SheetState,
  geom: GridGeometry,
  scrollX: number,
  scrollY: number,
  viewW: number,
  viewH: number,
): void {
  const sel = state.selection
  const rr = geom.rangeRect(selectionRange(sel))
  const fr = geom.cellRect(sel.focus.row, sel.focus.col)
  const preview = state.getField(fillPreviewKey) as CellRange | null | undefined
  const guide = state.getField(resizeGuideKey) as ResizeGuide | null | undefined

  for (const q of computeQuadrants(geom, scrollX, scrollY, viewW, viewH)) {
    const clip = quadrantGroup(q)
    const inner = clip.children[0] as Konva.Group
    // 选区填充 + 边框（clip 自然按象限分裂）
    inner.add(new Konva.Rect({ x: rr.x, y: rr.y, width: rr.w, height: rr.h, fill: COLOR_SELECT_FILL, ...noListen }))
    inner.add(new Konva.Rect({ x: rr.x, y: rr.y, width: rr.w, height: rr.h, stroke: COLOR_SELECT_BORDER, strokeWidth: 2, ...noListen }))
    // 活动格边框
    inner.add(new Konva.Rect({ x: fr.x, y: fr.y, width: fr.w, height: fr.h, stroke: COLOR_SELECT_BORDER, strokeWidth: 2, ...noListen }))
    // 填充手柄
    inner.add(
      new Konva.Rect({
        x: rr.x + rr.w - FILL_HANDLE_SIZE / 2,
        y: rr.y + rr.h - FILL_HANDLE_SIZE / 2,
        width: FILL_HANDLE_SIZE,
        height: FILL_HANDLE_SIZE,
        fill: COLOR_SELECT_BORDER,
        stroke: '#ffffff',
        strokeWidth: 1,
        ...noListen,
      }),
    )
    if (preview) {
      const pr = geom.rangeRect(preview)
      inner.add(
        new Konva.Rect({ x: pr.x, y: pr.y, width: pr.w, height: pr.h, stroke: COLOR_SELECT_BORDER, strokeWidth: 1, dash: [4, 3], ...noListen }),
      )
    }
    if (guide) {
      const points =
        guide.axis === 'col'
          ? [guide.pos, 0, guide.pos, geom.contentHeight]
          : [0, guide.pos, geom.contentWidth, guide.pos]
      inner.add(new Konva.Line({ points, stroke: COLOR_SELECT_BORDER, strokeWidth: 1, dash: [4, 3], ...noListen }))
    }
    layer.add(clip)
  }
}

// 滚动条：overlay 层屏幕坐标绘制（不进象限 clip）
function renderScrollbars(
  layer: Konva.Layer,
  geom: GridGeometry,
  scrollX: number,
  scrollY: number,
  viewW: number,
  viewH: number,
): void {
  const bars = [hScrollbar(geom.contentWidth, scrollX, viewW, viewH), vScrollbar(geom.contentHeight, scrollY, viewW, viewH)]
  for (const b of bars) {
    if (!b) continue
    layer.add(new Konva.Rect({ x: b.track.x, y: b.track.y, width: b.track.w, height: b.track.h, fill: '#f1f3f4', ...noListen }))
    layer.add(
      new Konva.Rect({
        x: b.thumb.x + 2,
        y: b.thumb.y + 2,
        width: Math.max(0, b.thumb.w - 4),
        height: Math.max(0, b.thumb.h - 4),
        fill: '#c1c7cd',
        cornerRadius: 4,
        ...noListen,
      }),
    )
  }
}
