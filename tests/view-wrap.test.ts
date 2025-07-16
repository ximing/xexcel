import { describe, expect, it } from 'vitest'
import { SheetData } from '@gmi/excel-core'
import { GridGeometry } from '../src/view/geometry'
import { wrappedLineCount, wrapRowHeight, WRAP_LINE_PAD } from '../src/view/measure'

// 等宽 mock：每字符 10px
const measure = (t: string): number => t.length * 10

describe('wrappedLineCount', () => {
  it('单行不超宽 → 1', () => {
    expect(wrappedLineCount('abc', undefined, 100, measure)).toBe(1)
  })
  it('贪心折行：abcd 宽 25 → 2 行', () => {
    expect(wrappedLineCount('abcd', undefined, 25, measure)).toBe(2)
  })
  it('\\n 强制换行；空段也算一行', () => {
    expect(wrappedLineCount('a\n\nb', undefined, 100, measure)).toBe(3)
  })
  it('空串 → 1', () => {
    expect(wrappedLineCount('', undefined, 100, measure)).toBe(1)
  })
  it('恰满一行不折', () => {
    expect(wrappedLineCount('abc', undefined, 30, measure)).toBe(1)
  })
})

describe('wrapRowHeight', () => {
  it('取行内最大需求：行数 × ceil(13×1.35) + pad', () => {
    const h = wrapRowHeight(
      [
        { text: 'abcd', style: undefined, contentWidth: 25 }, // 2 行
        { text: 'x', style: undefined, contentWidth: 100 }, // 1 行
      ],
      measure,
    )
    expect(h).toBe(2 * Math.ceil(13 * 1.35) + WRAP_LINE_PAD)
  })
  it('字号大者行距更大', () => {
    const h = wrapRowHeight([{ text: 'x', style: { fontSize: 24 }, contentWidth: 100 }], measure)
    expect(h).toBe(Math.ceil(24 * 1.35) + WRAP_LINE_PAD)
  })
})

describe('geometry 自动行高', () => {
  it('autoRows 撑高默认行；手动行高优先；隐藏行仍 0', () => {
    let s = SheetData.create({ rowCount: 5, colCount: 5 })
    s = s.setRowHeight(1, 50) // 手动
    s = s.setHidden('row', [2], true)
    const auto = new Map([[0, 60], [1, 80], [2, 90]])
    const g = new GridGeometry(s, 0, 0, undefined, auto)
    expect(g.rowHeight(0)).toBe(60)
    expect(g.rowHeight(1)).toBe(50) // 手动优先
    expect(g.rowHeight(2)).toBe(0) // 隐藏
    expect(g.rowHeight(3)).toBe(24) // 默认
  })
  it('autoRows 小于默认不压低', () => {
    const s = SheetData.create({ rowCount: 2, colCount: 2 })
    const g = new GridGeometry(s, 0, 0, undefined, new Map([[0, 10]]))
    expect(g.rowHeight(0)).toBe(24)
  })
})
