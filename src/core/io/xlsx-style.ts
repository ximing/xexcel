// src/core/io/xlsx-style.ts
// CellStyle ↔ exceljs 样式形状的纯互转。只镜像 exceljs 对象结构，不 import exceljs 运行时。
import { BorderEdge, BorderLineStyle, CFStyle, CellStyle } from '../model'

export interface XColor {
  argb: string
}
export interface XFont {
  bold?: boolean
  italic?: boolean
  underline?: boolean
  strike?: boolean
  color?: XColor
  name?: string
  size?: number
}
export interface XFill {
  type: 'pattern'
  pattern: 'solid'
  fgColor: XColor
}
export interface XAlignment {
  horizontal?: 'left' | 'center' | 'right'
  vertical?: 'top' | 'middle' | 'bottom'
  wrapText?: boolean
}
export interface XBorderEdge {
  style: string
  color?: XColor
}
export interface XBorders {
  top?: XBorderEdge
  right?: XBorderEdge
  bottom?: XBorderEdge
  left?: XBorderEdge
}
// CF differential style（dxf）：填充色在 bgColor，与单元格 fill 的 fgColor 不同
export interface XDiffFill {
  type: 'pattern'
  pattern: 'solid'
  bgColor: XColor
}
export interface XDiffStyle {
  font?: XFont
  fill?: XDiffFill
}

// '#rgb'/'#rrggbb' → 'FFRRGGBB'；非 hex 返回 null
export function cssToArgb(css: string): string | null {
  const m = /^#([0-9a-fA-F]{3}|[0-9a-fA-F]{6})$/.exec(css.trim())
  if (!m) return null
  let hex = m[1]
  if (hex.length === 3) hex = [...hex].map((c) => c + c).join('')
  return 'FF' + hex.toUpperCase()
}

// 'FFRRGGBB'（alpha 忽略）→ '#rrggbb'；非法返回 null
export function argbToCss(argb: string): string | null {
  if (!/^[0-9a-fA-F]{8}$/.test(argb.trim())) return null
  return '#' + argb.trim().slice(2).toLowerCase()
}

const emptyToUndefined = <T extends object>(o: T): T | undefined =>
  Object.keys(o).length === 0 ? undefined : o

export function styleToExcelFont(s: CellStyle): XFont | undefined {
  // 完全无 font 相关字段 → undefined；有字段但全部降级丢弃（如非 hex 颜色）→ 返回 {} 保留"有 font"语义
  const hasFontField =
    s.bold !== undefined || s.italic !== undefined || s.underline !== undefined ||
    s.strikethrough !== undefined || s.color !== undefined ||
    s.fontFamily !== undefined || s.fontSize !== undefined
  if (!hasFontField) return undefined
  const f: XFont = {}
  if (s.bold) f.bold = true
  if (s.italic) f.italic = true
  if (s.underline) f.underline = true
  if (s.strikethrough) f.strike = true
  if (s.color) {
    const argb = cssToArgb(s.color)
    if (argb) f.color = { argb }
  }
  if (s.fontFamily) f.name = s.fontFamily
  if (s.fontSize !== undefined) f.size = s.fontSize
  return f
}

export function styleToExcelFill(s: CellStyle): XFill | undefined {
  if (!s.bg) return undefined
  const argb = cssToArgb(s.bg)
  if (!argb) return undefined
  return { type: 'pattern', pattern: 'solid', fgColor: { argb } }
}

export function styleToExcelAlignment(s: CellStyle): XAlignment | undefined {
  const a: XAlignment = {}
  if (s.align) a.horizontal = s.align
  if (s.vAlign) a.vertical = s.vAlign
  if (s.wrap) a.wrapText = true
  return emptyToUndefined(a)
}

const edgeToExcel = (e: BorderEdge): XBorderEdge => {
  const out: XBorderEdge = { style: e.style }
  if (e.color) {
    const argb = cssToArgb(e.color)
    if (argb) out.color = { argb }
  }
  return out
}

export function styleToExcelBorders(s: CellStyle): XBorders | undefined {
  if (!s.border) return undefined
  const b: XBorders = {}
  if (s.border.top) b.top = edgeToExcel(s.border.top)
  if (s.border.right) b.right = edgeToExcel(s.border.right)
  if (s.border.bottom) b.bottom = edgeToExcel(s.border.bottom)
  if (s.border.left) b.left = edgeToExcel(s.border.left)
  return emptyToUndefined(b)
}

const KNOWN_BORDER_STYLES = new Set<BorderLineStyle>([
  'thin', 'medium', 'thick', 'dashed', 'dotted', 'double', 'hair', 'mediumDashed',
])

const edgeFromExcel = (e: XBorderEdge): BorderEdge | null => {
  if (!KNOWN_BORDER_STYLES.has(e.style as BorderLineStyle)) return null
  const out: BorderEdge = { style: e.style as BorderLineStyle }
  if (e.color?.argb) {
    const css = argbToCss(e.color.argb)
    if (css) out.color = css
  }
  return out
}

// theme 色/未知线型/未知 vertical 一律忽略（降级不打断）
export function styleFromExcel(f: {
  font?: XFont
  fill?: XFill
  alignment?: XAlignment
  border?: XBorders
}): CellStyle | undefined {
  const s: CellStyle = {}
  const font = f.font
  if (font) {
    if (font.bold) s.bold = true
    if (font.italic) s.italic = true
    if (font.underline) s.underline = true
    if (font.strike) s.strikethrough = true
    if (font.color?.argb) {
      const css = argbToCss(font.color.argb)
      if (css) s.color = css
    }
    if (font.name) s.fontFamily = font.name
    if (font.size !== undefined) s.fontSize = font.size
  }
  if (f.fill?.pattern === 'solid' && f.fill.fgColor?.argb) {
    const css = argbToCss(f.fill.fgColor.argb)
    if (css) s.bg = css
  }
  const a = f.alignment
  if (a) {
    if (a.horizontal === 'left' || a.horizontal === 'center' || a.horizontal === 'right') s.align = a.horizontal
    if (a.vertical === 'top' || a.vertical === 'middle' || a.vertical === 'bottom') s.vAlign = a.vertical
    if (a.wrapText) s.wrap = true
  }
  const b = f.border
  if (b) {
    const border: NonNullable<CellStyle['border']> = {}
    if (b.top) { const e = edgeFromExcel(b.top); if (e) border.top = e }
    if (b.right) { const e = edgeFromExcel(b.right); if (e) border.right = e }
    if (b.bottom) { const e = edgeFromExcel(b.bottom); if (e) border.bottom = e }
    if (b.left) { const e = edgeFromExcel(b.left); if (e) border.left = e }
    if (Object.keys(border).length) s.border = border
  }
  return emptyToUndefined(s)
}

// CFStyle → dxf（bgColor 语义）
export function cfStyleToExcel(s: CFStyle): XDiffStyle | undefined {
  const d: XDiffStyle = {}
  const font = styleToExcelFont(s as CellStyle) // CFStyle 是 CellStyle 子集字段
  if (font) d.font = font
  if (s.bg) {
    const argb = cssToArgb(s.bg)
    if (argb) d.fill = { type: 'pattern', pattern: 'solid', bgColor: { argb } }
  }
  return emptyToUndefined(d)
}

export function cfStyleFromExcel(d: XDiffStyle): CFStyle {
  const s: CFStyle = {}
  const font = d.font
  if (font) {
    if (font.bold) s.bold = true
    if (font.italic) s.italic = true
    if (font.underline) s.underline = true
    if (font.strike) s.strikethrough = true
    if (font.color?.argb) {
      const css = argbToCss(font.color.argb)
      if (css) s.color = css
    }
  }
  if (d.fill?.bgColor?.argb) {
    const css = argbToCss(d.fill.bgColor.argb)
    if (css) s.bg = css
  }
  return s
}
