// 边框预设计算：选区各格新 style（整体替换 style 键 border；其余样式键保留）。
import { CellRange } from './addr'
import { BorderEdge, CellStyle, SheetData } from './model'

export type BorderPreset = 'none' | 'all' | 'outer' | 'inner' | 'top' | 'bottom' | 'left' | 'right'
type Side = 'top' | 'right' | 'bottom' | 'left'

// edge=null（或 preset='none'）→ 清除相应边；返回逐格 style（空样式空格 → null）
export function computeBorderStyles(
  sheet: SheetData,
  range: CellRange,
  preset: BorderPreset,
  edge: BorderEdge | null,
): { row: number; col: number; style: CellStyle | null }[] {
  const out: { row: number; col: number; style: CellStyle | null }[] = []
  const clear = preset === 'none'
  for (let r = range.sr; r <= range.er; r++) {
    for (let c = range.sc; c <= range.ec; c++) {
      const cell = sheet.getCell(r, c)
      const style: CellStyle = { ...(cell?.style ?? {}) }
      const border: { [k in Side]?: BorderEdge } = { ...(style.border ?? {}) }
      const set = (side: Side, on: boolean): void => {
        if (!on) return
        if (clear || edge === null) delete border[side]
        else border[side] = edge // 后写覆盖：预设触及的边一律替换
      }
      const isTop = r === range.sr
      const isBottom = r === range.er
      const isLeft = c === range.sc
      const isRight = c === range.ec
      switch (preset) {
        case 'none':
          delete border.top; delete border.right; delete border.bottom; delete border.left
          break
        case 'all':
          set('top', true); set('right', true); set('bottom', true); set('left', true)
          break
        case 'outer':
          set('top', isTop); set('bottom', isBottom); set('left', isLeft); set('right', isRight)
          break
        case 'inner':
          set('top', !isTop); set('bottom', !isBottom); set('left', !isLeft); set('right', !isRight)
          break
        case 'top': set('top', isTop); break
        case 'bottom': set('bottom', isBottom); break
        case 'left': set('left', isLeft); break
        case 'right': set('right', isRight); break
      }
      if (Object.keys(border).length === 0) delete style.border
      else style.border = border
      const raw = cell?.raw ?? ''
      if (Object.keys(style).length === 0 && raw === '') out.push({ row: r, col: c, style: null })
      else out.push({ row: r, col: c, style })
    }
  }
  return out
}
