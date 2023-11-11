import { describe, it, expect } from 'vitest'
import { SheetData, DEFAULT_ROW_HEIGHT, DEFAULT_COL_WIDTH } from '../src/core/model'
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
