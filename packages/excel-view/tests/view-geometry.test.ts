import { describe, it, expect } from 'vitest'
import { SheetData, DEFAULT_ROW_HEIGHT, DEFAULT_COL_WIDTH } from '@xexcel/core'
import { GridGeometry } from '../src/view/geometry'

describe('GridGeometry', () => {
  const sheet = SheetData.create({ rowCount: 1000, colCount: 26 }).setRowHeight(0, 40).setColWidth(0, 120)
  const g = new GridGeometry(sheet)
  it('坐标正反差拍', () => {
    expect(g.rowTop(0)).toBe(0)
    expect(g.rowTop(1)).toBe(40)               // 第 0 行自定义 40
    expect(g.rowTop(3)).toBe(40 + 2 * DEFAULT_ROW_HEIGHT)
    expect(g.colLeft(1)).toBe(120)
    expect(g.colLeft(3)).toBe(120 + 2 * DEFAULT_COL_WIDTH)
    expect(g.rowAt(0)).toBe(0)
    expect(g.rowAt(39)).toBe(0)
    expect(g.rowAt(40)).toBe(1)
    expect(g.colAt(119)).toBe(0)
    expect(g.colAt(120)).toBe(1)
  })
  it('cellRect / rangeRect / visibleRange / contentSize', () => {
    expect(g.cellRect(1, 1)).toEqual({ x: 120, y: 40, w: DEFAULT_COL_WIDTH, h: DEFAULT_ROW_HEIGHT })
    expect(g.rangeRect({ sr: 0, sc: 0, er: 1, ec: 1 })).toEqual({ x: 0, y: 0, w: 120 + DEFAULT_COL_WIDTH, h: 40 + DEFAULT_ROW_HEIGHT })
    const vr = g.visibleRange(0, 0, 480, 240)
    expect(vr.sr).toBe(0); expect(vr.sc).toBe(0)
    expect(vr.er).toBeGreaterThanOrEqual(4)
    expect(vr.ec).toBeGreaterThanOrEqual(2)
    expect(g.contentWidth).toBe(120 + 25 * DEFAULT_COL_WIDTH)
    expect(g.contentHeight).toBe(40 + 999 * DEFAULT_ROW_HEIGHT)
  })
  it('越界 clamp', () => {
    expect(g.rowAt(-100)).toBe(0)
    expect(g.rowAt(1e9)).toBe(999)
    expect(g.colAt(1e9)).toBe(25)
  })
})

describe('extraHiddenRows（筛选隐藏注入）', () => {
  it('额外隐藏行高度折叠为 0 且不可命中', () => {
    const sheet = SheetData.create({ rowCount: 5, colCount: 5 })
    const g = new GridGeometry(sheet, 0, 0, new Set([1, 3]))
    expect(g.rowHeight(1)).toBe(0)
    expect(g.rowHeight(3)).toBe(0)
    expect(g.rowAt(24)).toBe(2) // 行 1 被跳过
    expect(g.rowAt(48)).toBe(4) // 行 3 被跳过
  })
  it('cellRect：隐藏行 h=0，相邻可见行不受影响', () => {
    const sheet = SheetData.create({ rowCount: 5, colCount: 5 })
    const g = new GridGeometry(sheet, 0, 0, new Set([1, 3]))
    expect(g.cellRect(1, 2)).toEqual({ x: 2 * DEFAULT_COL_WIDTH, y: DEFAULT_ROW_HEIGHT, w: DEFAULT_COL_WIDTH, h: 0 })
    expect(g.cellRect(3, 2).h).toBe(0)
    // 相邻可见行：位置塌缩到隐藏行原位，但高度保持完整
    expect(g.cellRect(2, 2)).toEqual({ x: 2 * DEFAULT_COL_WIDTH, y: DEFAULT_ROW_HEIGHT, w: DEFAULT_COL_WIDTH, h: DEFAULT_ROW_HEIGHT })
    expect(g.cellRect(4, 2)).toEqual({ x: 2 * DEFAULT_COL_WIDTH, y: 2 * DEFAULT_ROW_HEIGHT, w: DEFAULT_COL_WIDTH, h: DEFAULT_ROW_HEIGHT })
    expect(g.cellRect(0, 2)).toEqual({ x: 2 * DEFAULT_COL_WIDTH, y: 0, w: DEFAULT_COL_WIDTH, h: DEFAULT_ROW_HEIGHT })
  })
  it('rangeRect：跨隐藏行时高度排除隐藏行', () => {
    const sheet = SheetData.create({ rowCount: 5, colCount: 5 })
    const g = new GridGeometry(sheet, 0, 0, new Set([1, 3]))
    // 行 0..4 含两个隐藏行 → 高度 = 3 个可见行
    expect(g.rangeRect({ sr: 0, sc: 0, er: 4, ec: 1 })).toEqual({
      x: 0, y: 0, w: 2 * DEFAULT_COL_WIDTH, h: 3 * DEFAULT_ROW_HEIGHT,
    })
    // 完全落在隐藏行上 → h=0
    expect(g.rangeRect({ sr: 1, sc: 0, er: 1, ec: 1 }).h).toBe(0)
    // 不含隐藏行的区间不受影响
    expect(g.rangeRect({ sr: 0, sc: 0, er: 0, ec: 1 })).toEqual({
      x: 0, y: 0, w: 2 * DEFAULT_COL_WIDTH, h: DEFAULT_ROW_HEIGHT,
    })
  })
})
