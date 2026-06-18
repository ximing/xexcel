// 文本测量：列宽自适应用。canvas 部分仅浏览器可用；纯逻辑部分（optimal*）注入测量函数，node 可测。
import { CellStyle } from '@xexcel/core'

export const CELL_PAD_X = 6
const DEFAULT_FONT_SIZE = 13
const DEFAULT_FONT_FAMILY =
  '-apple-system, BlinkMacSystemFont, "Segoe UI", Roboto, "Helvetica Neue", Arial, sans-serif'

let ctx: CanvasRenderingContext2D | null = null

function fontOf(style?: CellStyle): string {
  const size = style?.fontSize ?? DEFAULT_FONT_SIZE
  const family = style?.fontFamily ?? DEFAULT_FONT_FAMILY
  const weight = style?.bold ? 'bold ' : ''
  const italic = style?.italic ? 'italic ' : ''
  return `${italic}${weight}${size}px ${family}`
}

// 浏览器专用：canvas measureText（node 环境返回 0，调用侧仅浏览器路径使用）
export function measureTextWidth(text: string, style?: CellStyle): number {
  if (typeof document === 'undefined') return 0
  if (!ctx) ctx = document.createElement('canvas').getContext('2d')
  if (!ctx) return 0
  ctx.font = fontOf(style)
  return ctx.measureText(text).width
}

// 列最优宽：最大文本宽 + 左右 padding + 4px 余量；空列 → null（恢复默认）
export function optimalColWidth(
  items: { text: string; style?: CellStyle }[],
  measure: (t: string, s?: CellStyle) => number,
): number | null {
  if (items.length === 0) return null
  let max = 0
  for (const it of items) {
    if (it.text === '') continue
    max = Math.max(max, measure(it.text, it.style))
  }
  if (max === 0) return null
  return Math.ceil(max) + CELL_PAD_X * 2 + 4
}

// 行最优高：最大字号 × 1.35 + 6，下限 20；空行 → null（恢复默认）
export function optimalRowHeight(items: { style?: CellStyle }[]): number | null {
  if (items.length === 0) return null
  let size = 0
  for (const it of items) size = Math.max(size, it.style?.fontSize ?? DEFAULT_FONT_SIZE)
  if (size === 0) return null
  return Math.max(20, Math.ceil(size * 1.35) + 6)
}

export const WRAP_LINE_PAD = 6

// 折行数估计：\n 强制换行；段内贪心按字符填充（CJK 友好，与 Konva wrap:'char' 近似）。
// measure 注入便于 node 单测；调用侧传 measureTextWidth。
export function wrappedLineCount(
  text: string,
  style: CellStyle | undefined,
  contentWidth: number,
  measure: (t: string, s?: CellStyle) => number,
): number {
  if (text === '') return 1
  let lines = 0
  for (const para of text.split('\n')) {
    lines++
    let cur = ''
    for (const ch of para) {
      if (cur !== '' && measure(cur + ch, style) > contentWidth) {
        lines++
        cur = ch
      } else {
        cur += ch
      }
    }
  }
  return lines
}

// 一组 wrap 格的行高需求最大值；无项 → 0（调用侧与默认行高取 max）
export function wrapRowHeight(
  items: { text: string; style?: CellStyle; contentWidth: number }[],
  measure: (t: string, s?: CellStyle) => number,
): number {
  let need = 0
  for (const it of items) {
    const lines = wrappedLineCount(it.text, it.style, it.contentWidth, measure)
    const size = it.style?.fontSize ?? DEFAULT_FONT_SIZE
    need = Math.max(need, lines * Math.ceil(size * 1.35) + WRAP_LINE_PAD)
  }
  return need
}
