// 缩放：档位常量与锚点换算（纯逻辑，node 可测）。zoom=1 为 100%。
import type { SheetId } from '@xexcel/core'
import type { SheetState } from '@xexcel/core'
import { zoomKey } from './types'

export const ZOOM_LEVELS = [0.5, 0.75, 0.9, 1, 1.25, 1.5, 2]

// 当前表 zoom（缺省 1）
export function zoomOf(state: SheetState, sheetId: SheetId): number {
  const z = state.getField(zoomKey) as Record<string, number> | null | undefined
  return z?.[sheetId] ?? 1
}

// 档位步进：dir=+1 放大；非档值先就近归位（放大取上档，缩小取下档）
export function nextZoomLevel(z: number, dir: 1 | -1): number {
  const eps = 1e-9
  if (dir === 1) {
    for (const l of ZOOM_LEVELS) if (l > z + eps) return l
    return ZOOM_LEVELS[ZOOM_LEVELS.length - 1]
  }
  for (let i = ZOOM_LEVELS.length - 1; i >= 0; i--) if (ZOOM_LEVELS[i] < z - eps) return ZOOM_LEVELS[i]
  return ZOOM_LEVELS[0]
}

// 锚点保持：光标相对内容区偏移 cursor（px），缩放后光标下内容不变
export function anchoredScroll(scroll: number, cursor: number, z0: number, z1: number): number {
  return (scroll + cursor) * (z1 / z0) - cursor
}
