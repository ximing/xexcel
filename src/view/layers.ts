// 三层渲染：gridLayer 背景+表头；cellLayer 单元格 bg+网格线+文字；overlayLayer 选区+填充手柄。
// 每次 updateState 后整层重建 visibleRange 内节点（M1 从简，1000 格/帧量级可接受；
// 如需进一步优化可在此引入 nodePool 节点池复用 Konva 节点）。
import Konva from 'konva'
import { CellRange, colName } from '../core/addr'
import { COL_HEADER_HEIGHT, ROW_HEADER_WIDTH } from '../core/model'
import { selectionRange } from '../core/selection'
import type { SheetState } from '../core/state'
import { evaluatorFor } from '../formula/engine'
import { GridGeometry } from './geometry'
import { fillPreviewKey, ResizeGuide, resizeGuideKey } from './types'

const COLOR_GRID = '#d9dce1'
const COLOR_HEADER_BG = '#f7f8fa'
const COLOR_HEADER_TEXT = '#5f6368'
const COLOR_HEADER_ACTIVE_BG = '#e8f0fe'
const COLOR_HEADER_ACTIVE_TEXT = '#1a73e8'
const COLOR_SELECT_FILL = 'rgba(26, 115, 232, 0.12)'
const COLOR_SELECT_BORDER = '#1a73e8'
const COLOR_TEXT = '#202124'
const FONT_SIZE = 13
const FONT_FAMILY =
  '-apple-system, BlinkMacSystemFont, "Segoe UI", Roboto, "Helvetica Neue", Arial, sans-serif'
const CELL_PAD_X = 6
export const FILL_HANDLE_SIZE = 6

const noListen = { listening: false }

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
  const vr = geom.visibleRange(
    scrollX,
    scrollY,
    Math.max(0, viewW - ROW_HEADER_WIDTH),
    Math.max(0, viewH - COL_HEADER_HEIGHT),
  )
  renderCellLayer(cellLayer, state, geom, vr, scrollX, scrollY, viewW, viewH)
  renderGridLayer(gridLayer, state, geom, vr, scrollX, scrollY, viewW, viewH)
  renderOverlayLayer(overlayLayer, state, geom, scrollX, scrollY, viewW, viewH)
  gridLayer.batchDraw()
  cellLayer.batchDraw()
  overlayLayer.batchDraw()
}

// 内容区裁剪容器：表头带以外才绘制，内部按 -scroll 偏移画内容坐标
function contentClip(scrollX: number, scrollY: number, viewW: number, viewH: number): Konva.Group {
  const clip = new Konva.Group({
    x: ROW_HEADER_WIDTH,
    y: COL_HEADER_HEIGHT,
    clipX: 0,
    clipY: 0,
    clipWidth: Math.max(0, viewW - ROW_HEADER_WIDTH),
    clipHeight: Math.max(0, viewH - COL_HEADER_HEIGHT),
    ...noListen,
  })
  clip.add(new Konva.Group({ x: -scrollX, y: -scrollY, ...noListen }))
  return clip
}

function renderGridLayer(
  layer: Konva.Layer,
  state: SheetState,
  geom: GridGeometry,
  vr: CellRange,
  scrollX: number,
  scrollY: number,
  viewW: number,
  viewH: number,
): void {
  // 内容区白底
  layer.add(
    new Konva.Rect({
      x: ROW_HEADER_WIDTH,
      y: COL_HEADER_HEIGHT,
      width: viewW,
      height: viewH,
      fill: '#ffffff',
      ...noListen,
    }),
  )
  // 表头带背景
  layer.add(new Konva.Rect({ x: 0, y: 0, width: viewW, height: COL_HEADER_HEIGHT, fill: COLOR_HEADER_BG, ...noListen }))
  layer.add(new Konva.Rect({ x: 0, y: 0, width: ROW_HEADER_WIDTH, height: viewH, fill: COLOR_HEADER_BG, ...noListen }))

  const sel = selectionRange(state.selection)
  // 列头：选区覆盖的列强调
  for (let c = vr.sc; c <= vr.ec; c++) {
    const x = ROW_HEADER_WIDTH + geom.colLeft(c) - scrollX
    const w = geom.sheet.colWidth(c)
    const active = c >= sel.sc && c <= sel.ec
    if (active) {
      layer.add(new Konva.Rect({ x, y: 0, width: w, height: COL_HEADER_HEIGHT, fill: COLOR_HEADER_ACTIVE_BG, ...noListen }))
    }
    layer.add(
      new Konva.Text({
        x,
        y: 0,
        width: w,
        height: COL_HEADER_HEIGHT,
        text: colName(c),
        align: 'center',
        verticalAlign: 'middle',
        fontSize: FONT_SIZE,
        fontFamily: FONT_FAMILY,
        fill: active ? COLOR_HEADER_ACTIVE_TEXT : COLOR_HEADER_TEXT,
        wrap: 'none',
        ...noListen,
      }),
    )
    // 列头间隔线
    layer.add(
      new Konva.Line({ points: [x, 0, x, COL_HEADER_HEIGHT], stroke: COLOR_GRID, strokeWidth: 1, ...noListen }),
    )
  }
  // 行头：选区覆盖的行强调
  for (let r = vr.sr; r <= vr.er; r++) {
    const y = COL_HEADER_HEIGHT + geom.rowTop(r) - scrollY
    const h = geom.sheet.rowHeight(r)
    const active = r >= sel.sr && r <= sel.er
    if (active) {
      layer.add(new Konva.Rect({ x: 0, y, width: ROW_HEADER_WIDTH, height: h, fill: COLOR_HEADER_ACTIVE_BG, ...noListen }))
    }
    layer.add(
      new Konva.Text({
        x: 0,
        y,
        width: ROW_HEADER_WIDTH,
        height: h,
        text: String(r + 1),
        align: 'center',
        verticalAlign: 'middle',
        fontSize: FONT_SIZE,
        fontFamily: FONT_FAMILY,
        fill: active ? COLOR_HEADER_ACTIVE_TEXT : COLOR_HEADER_TEXT,
        wrap: 'none',
        ...noListen,
      }),
    )
    layer.add(
      new Konva.Line({ points: [0, y, ROW_HEADER_WIDTH, y], stroke: COLOR_GRID, strokeWidth: 1, ...noListen }),
    )
  }
  // 表头带外缘分隔线
  layer.add(
    new Konva.Line({ points: [0, COL_HEADER_HEIGHT, viewW, COL_HEADER_HEIGHT], stroke: COLOR_GRID, strokeWidth: 1, ...noListen }),
  )
  layer.add(
    new Konva.Line({ points: [ROW_HEADER_WIDTH, 0, ROW_HEADER_WIDTH, viewH], stroke: COLOR_GRID, strokeWidth: 1, ...noListen }),
  )
}

function renderCellLayer(
  layer: Konva.Layer,
  state: SheetState,
  geom: GridGeometry,
  vr: CellRange,
  scrollX: number,
  scrollY: number,
  viewW: number,
  viewH: number,
): void {
  const clip = contentClip(scrollX, scrollY, viewW, viewH)
  const inner = clip.children[0] as Konva.Group
  const sheet = state.activeSheet
  const sheetId = state.doc.active
  const evaluator = evaluatorFor(state.doc)

  const left = geom.colLeft(vr.sc)
  const right = geom.colLeft(vr.ec) + sheet.colWidth(vr.ec)
  const top = geom.rowTop(vr.sr)
  const bottom = geom.rowTop(vr.er) + sheet.rowHeight(vr.er)

  // 绘制顺序：单元格 bg → 网格线 → 文字
  for (let r = vr.sr; r <= vr.er; r++) {
    for (let c = vr.sc; c <= vr.ec; c++) {
      const cell = sheet.getCell(r, c)
      if (!cell?.style?.bg) continue
      const rect = geom.cellRect(r, c)
      inner.add(new Konva.Rect({ x: rect.x, y: rect.y, width: rect.w, height: rect.h, fill: cell.style.bg, ...noListen }))
    }
  }
  for (let c = vr.sc; c <= vr.ec + 1; c++) {
    const x = geom.colLeft(c)
    inner.add(new Konva.Line({ points: [x, top, x, bottom], stroke: COLOR_GRID, strokeWidth: 1, ...noListen }))
  }
  for (let r = vr.sr; r <= vr.er + 1; r++) {
    const y = geom.rowTop(r)
    inner.add(new Konva.Line({ points: [left, y, right, y], stroke: COLOR_GRID, strokeWidth: 1, ...noListen }))
  }
  for (let r = vr.sr; r <= vr.er; r++) {
    for (let c = vr.sc; c <= vr.ec; c++) {
      const cell = sheet.getCell(r, c)
      if (!cell || cell.raw === '') continue
      const text = evaluator.displayText(sheetId, r, c)
      if (text === '') continue
      const rect = geom.cellRect(r, c)
      let align = cell.style?.align
      if (!align) {
        // 数字/布尔默认右对齐
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
      // M1 简化：文字始终裁剪到单元格矩形（不溢出到相邻空格）
      const g = new Konva.Group({ clipX: rect.x, clipY: rect.y, clipWidth: rect.w, clipHeight: rect.h, ...noListen })
      g.add(
        new Konva.Text({
          x: rect.x + CELL_PAD_X,
          y: rect.y,
          width: Math.max(0, rect.w - CELL_PAD_X * 2),
          height: rect.h,
          text,
          align,
          verticalAlign: 'middle',
          fontSize: FONT_SIZE,
          fontFamily: FONT_FAMILY,
          fontStyle,
          fill: cell.style?.color ?? COLOR_TEXT,
          wrap: 'none',
          ...noListen,
        }),
      )
      inner.add(g)
    }
  }
  layer.add(clip)
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
  const clip = contentClip(scrollX, scrollY, viewW, viewH)
  const inner = clip.children[0] as Konva.Group

  const sel = state.selection
  const rr = geom.rangeRect(selectionRange(sel))
  // 选区半透明蓝填充 + 2px 边框
  inner.add(new Konva.Rect({ x: rr.x, y: rr.y, width: rr.w, height: rr.h, fill: COLOR_SELECT_FILL, ...noListen }))
  inner.add(
    new Konva.Rect({ x: rr.x, y: rr.y, width: rr.w, height: rr.h, stroke: COLOR_SELECT_BORDER, strokeWidth: 2, ...noListen }),
  )
  // 活动单元格 2px 边框（单格选区时与选区框重合）
  const fr = geom.cellRect(sel.focus.row, sel.focus.col)
  inner.add(
    new Konva.Rect({ x: fr.x, y: fr.y, width: fr.w, height: fr.h, stroke: COLOR_SELECT_BORDER, strokeWidth: 2, ...noListen }),
  )
  // 选区右下角 6×6 填充手柄方块
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
  // 填充手柄拖拽预览：非空则画 1px 虚线框（Task 6 写入）
  const preview = state.getField(fillPreviewKey) as CellRange | null | undefined
  if (preview) {
    const pr = geom.rangeRect(preview)
    inner.add(
      new Konva.Rect({
        x: pr.x,
        y: pr.y,
        width: pr.w,
        height: pr.h,
        stroke: COLOR_SELECT_BORDER,
        strokeWidth: 1,
        dash: [4, 3],
        ...noListen,
      }),
    )
  }
  // 行列调宽拖拽参考线：非空则画通长 1px 虚线（selection 插件写入）
  const guide = state.getField(resizeGuideKey) as ResizeGuide | null | undefined
  if (guide) {
    const points =
      guide.axis === 'col'
        ? [guide.pos, 0, guide.pos, geom.contentHeight]
        : [0, guide.pos, geom.contentWidth, guide.pos]
    inner.add(
      new Konva.Line({ points, stroke: COLOR_SELECT_BORDER, strokeWidth: 1, dash: [4, 3], ...noListen }),
    )
  }
  layer.add(clip)
}
