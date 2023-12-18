import { describe, it, expect } from 'vitest'
import { CELL_PAD_X, optimalColWidth, optimalRowHeight } from '../src/view/measure'

// 假测量器：宽度 = 字符数 × 10 + fontSize 加成
const fakeMeasure = (t: string, s?: { fontSize?: number }): number => t.length * 10 + (s?.fontSize ?? 13)

describe('optimalColWidth', () => {
  it('取最大宽度 + padding，向上取整', () => {
    const w = optimalColWidth(
      [{ text: 'abc' }, { text: 'abcdef', style: { fontSize: 20 } }, { text: '' }],
      fakeMeasure,
    )
    // max = 6*10+20 = 80 → 80 + 12 + 4 = 96
    expect(w).toBe(80 + CELL_PAD_X * 2 + 4)
  })
  it('空列 → null（调用侧恢复默认）', () => {
    expect(optimalColWidth([], fakeMeasure)).toBeNull()
  })
})

describe('optimalRowHeight', () => {
  it('按最大字号估算，下限 20', () => {
    expect(optimalRowHeight([{ style: { fontSize: 24 } }])).toBe(Math.ceil(24 * 1.35) + 6)
    expect(optimalRowHeight([{ style: {} }])).toBe(Math.max(20, Math.ceil(13 * 1.35) + 6))
    expect(optimalRowHeight([])).toBeNull()
  })
})
