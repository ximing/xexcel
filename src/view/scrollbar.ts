// 自绘滚动条几何：轨道/滑块矩形与拖拽映射比例。纯数学，node 可测。
import { COL_HEADER_HEIGHT, ROW_HEADER_WIDTH } from '../core/model'
import { Rect } from './types'

export const SB_SIZE = 12
export const SB_MIN_THUMB = 24

export interface ScrollbarGeom {
  track: Rect
  thumb: Rect
  maxScroll: number
  ratio: number // maxScroll / (trackLen - thumbLen)；拖拽 1px = ratio 滚动单位
}

function geom(track: Rect, viewportLen: number, contentLen: number, scroll: number, vertical: boolean): ScrollbarGeom | null {
  const maxScroll = Math.max(0, contentLen - viewportLen)
  if (maxScroll <= 0) return null
  const trackLen = vertical ? track.h : track.w
  const thumbLen = Math.max(SB_MIN_THUMB, trackLen * (viewportLen / contentLen))
  const ratio = maxScroll / (trackLen - thumbLen)
  const offset = scroll / ratio
  const thumb: Rect = vertical
    ? { x: track.x, y: track.y + offset, w: track.w, h: thumbLen }
    : { x: track.x + offset, y: track.y, w: thumbLen, h: track.h }
  return { track, thumb, maxScroll, ratio }
}

export function hScrollbar(contentWidth: number, scrollX: number, viewW: number, viewH: number): ScrollbarGeom | null {
  const track: Rect = { x: ROW_HEADER_WIDTH, y: viewH - SB_SIZE, w: viewW - ROW_HEADER_WIDTH - SB_SIZE, h: SB_SIZE }
  return geom(track, viewW - ROW_HEADER_WIDTH, contentWidth, scrollX, false)
}

export function vScrollbar(contentHeight: number, scrollY: number, viewW: number, viewH: number): ScrollbarGeom | null {
  const track: Rect = { x: viewW - SB_SIZE, y: COL_HEADER_HEIGHT, w: SB_SIZE, h: viewH - COL_HEADER_HEIGHT - SB_SIZE }
  return geom(track, viewH - COL_HEADER_HEIGHT, contentHeight, scrollY, true)
}

export function thumbHit(g: ScrollbarGeom, x: number, y: number): boolean {
  return x >= g.thumb.x && x <= g.thumb.x + g.thumb.w && y >= g.thumb.y && y <= g.thumb.y + g.thumb.h
}

// 内容视口尺寸（w0/h0 = 扣表头后的区域）：对侧滚动条存在时扣其厚度。
// 双轮判定：横条占高可能引发纵条需求，纵条占宽也可能反过来引发横条。
export function contentViewport(
  contentW: number,
  contentH: number,
  w0: number,
  h0: number,
): { w: number; h: number } {
  const needH = contentW > w0
  const needV = contentH > (needH ? h0 - SB_SIZE : h0)
  const needH2 = contentW > (needV ? w0 - SB_SIZE : w0)
  return { w: needV ? w0 - SB_SIZE : w0, h: needH2 ? h0 - SB_SIZE : h0 }
}
