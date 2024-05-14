// 自动行高推导：wrap 格按显示文本折行需求撑高（推导值不入模型；手动行高行不参与）。
// 行高按 (宽度|字体|文本) 缓存；列宽/字号/内容变化经 sheet 引用变化自然失效。
import { CellStyle, SheetData } from '../core/model'
import type { CellEvaluator } from '../formula/engine'
import { CELL_PAD_X, measureTextWidth, wrapRowHeight } from './measure'

const heightCache = new Map<string, number>()

function fontKey(style?: CellStyle): string {
  return `${style?.bold ? 'b' : ''}${style?.italic ? 'i' : ''}${style?.fontSize ?? 13}|${style?.fontFamily ?? ''}`
}

// 缓存的是需求行高（wrapRowHeight 结果），key 含宽度与字体
function cachedRowHeight(text: string, style: CellStyle | undefined, width: number): number {
  const key = `${width}|${fontKey(style)}|${text}`
  const hit = heightCache.get(key)
  if (hit !== undefined) return hit
  const n = wrapRowHeight([{ text, style, contentWidth: width }], measureTextWidth)
  heightCache.set(key, n)
  if (heightCache.size > 5000) heightCache.clear() // 防无限增长
  return n
}

// 返回 row → 需求行高（仅含 wrap 格的行；调用侧与默认行高取 max）
export function autoRowHeights(sheet: SheetData, sheetId: string, ev: CellEvaluator): Map<number, number> {
  const out = new Map<number, number>()
  const used = sheet.usedRange()
  for (let row = used.sr; row <= used.er; row++) {
    let need = 0
    for (let col = used.sc; col <= used.ec; col++) {
      const cell = sheet.getCell(row, col)
      if (!cell?.style?.wrap || cell.raw === '') continue
      const m = sheet.mergeAt(row, col)
      if (m && (m.sr !== row || m.sc !== col)) continue // 非锚点不撑高
      let w = 0
      if (m) for (let c = m.sc; c <= m.ec; c++) w += sheet.colWidth(c)
      else w = sheet.colWidth(col)
      if (w <= CELL_PAD_X * 2) continue
      const text = ev.displayText(sheetId, row, col)
      if (text === '') continue
      need = Math.max(need, cachedRowHeight(text, cell.style, w - CELL_PAD_X * 2))
    }
    if (need > 0) out.set(row, need)
  }
  return out
}
