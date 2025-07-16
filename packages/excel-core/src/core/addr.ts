// 单元格地址与区域工具。索引一律 0-based；A1 表示法仅在此层转换。
export interface CellAddr { row: number; col: number }
export interface CellRange { sr: number; sc: number; er: number; ec: number } // 闭区间

export function colName(col: number): string {
  let s = ''
  let n = col
  do { s = String.fromCharCode(65 + (n % 26)) + s; n = Math.floor(n / 26) - 1 } while (n >= 0)
  return s
}

export function parseColName(s: string): number {
  if (!/^[A-Za-z]+$/.test(s)) return -1
  let n = 0
  const up = s.toUpperCase()
  for (let i = 0; i < up.length; i++) n = n * 26 + (up.charCodeAt(i) - 64)
  return n - 1
}

export function toA1(row: number, col: number): string {
  return colName(col) + (row + 1)
}

export function fromA1(s: string): CellAddr | null {
  const m = /^([A-Za-z]+)([1-9][0-9]*)$/.exec(s)
  if (!m) return null
  const col = parseColName(m[1])
  if (col < 0) return null
  return { row: parseInt(m[2], 10) - 1, col }
}

export function normalizeRange(r: CellRange): CellRange {
  return {
    sr: Math.min(r.sr, r.er),
    sc: Math.min(r.sc, r.ec),
    er: Math.max(r.sr, r.er),
    ec: Math.max(r.sc, r.ec),
  }
}

export function parseRange(s: string): CellRange | null {
  const parts = s.split(':')
  if (parts.length > 2) return null
  const a = fromA1(parts[0])
  if (!a) return null
  if (parts.length === 1) return { sr: a.row, sc: a.col, er: a.row, ec: a.col }
  const b = fromA1(parts[1])
  if (!b) return null
  return normalizeRange({ sr: a.row, sc: a.col, er: b.row, ec: b.col })
}

// UI 输入用 A1 解析（公式层有自己的 lexer，不共用）
export function parseA1(text: string): CellAddr | null {
  const m = /^([A-Za-z]+)([0-9]+)$/.exec(text.trim())
  if (!m) return null
  let col = 0
  for (const ch of m[1].toUpperCase()) col = col * 26 + (ch.charCodeAt(0) - 64)
  return { row: parseInt(m[2], 10) - 1, col: col - 1 }
}

export function parseRangeA1(text: string): CellRange | null {
  const parts = text.trim().split(':')
  if (parts.length > 2) return null
  const a = parseA1(parts[0])
  const b = parts.length === 2 ? parseA1(parts[1]) : a
  if (!a || !b) return null
  return normalizeRange({ sr: a.row, sc: a.col, er: b.row, ec: b.col })
}

export function rangeContains(r: CellRange, row: number, col: number): boolean {
  return row >= r.sr && row <= r.er && col >= r.sc && col <= r.ec
}

export function rangeCellCount(r: CellRange): number {
  return (r.er - r.sr + 1) * (r.ec - r.sc + 1)
}

export function clampRange(r: CellRange, maxRow: number, maxCol: number): CellRange {
  const n = normalizeRange(r)
  return {
    sr: Math.max(0, Math.min(n.sr, maxRow - 1)),
    sc: Math.max(0, Math.min(n.sc, maxCol - 1)),
    er: Math.max(0, Math.min(n.er, maxRow - 1)),
    ec: Math.max(0, Math.min(n.ec, maxCol - 1)),
  }
}

export function rangesEqual(a: CellRange, b: CellRange): boolean {
  return a.sr === b.sr && a.sc === b.sc && a.er === b.er && a.ec === b.ec
}

export function rangesIntersect(a: CellRange, b: CellRange): boolean {
  return a.sr <= b.er && a.er >= b.sr && a.sc <= b.ec && a.ec >= b.sc
}

export function wholeRange(maxRow: number, maxCol: number): CellRange {
  return { sr: 0, sc: 0, er: maxRow - 1, ec: maxCol - 1 }
}
