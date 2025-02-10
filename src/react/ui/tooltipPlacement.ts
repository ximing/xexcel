// Tooltip 方位：auto 时下方够放就朝下，否则朝上
export type TipPlacement = 'top' | 'bottom'
export function resolvePlacement(
  anchorBottom: number,
  anchorTop: number,
  tipHeight: number,
  preferred: 'auto' | TipPlacement,
  viewportH: number,
): TipPlacement {
  if (preferred !== 'auto') return preferred
  return anchorBottom + 6 + tipHeight <= viewportH ? 'bottom' : 'top'
}
