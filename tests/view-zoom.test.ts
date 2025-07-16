import { describe, expect, it } from 'vitest'
import { SheetData } from '@gmi/excel-core'
import { GridGeometry } from '../src/view/geometry'
import { anchoredScroll, nextZoomLevel, ZOOM_LEVELS } from '../src/view/zoom'

describe('zoom 档位', () => {
  it('档位表升序且覆盖 50%–200%，含 100%', () => {
    expect(ZOOM_LEVELS[0]).toBe(0.5)
    expect(ZOOM_LEVELS[ZOOM_LEVELS.length - 1]).toBe(2)
    expect(ZOOM_LEVELS).toContain(1)
    expect([...ZOOM_LEVELS].sort((a, b) => a - b)).toEqual(ZOOM_LEVELS)
  })
  it('放大/缩小按档步进，边界 clamp', () => {
    expect(nextZoomLevel(1, 1)).toBe(1.25)
    expect(nextZoomLevel(1, -1)).toBe(0.9)
    expect(nextZoomLevel(2, 1)).toBe(2)
    expect(nextZoomLevel(0.5, -1)).toBe(0.5)
  })
  it('非档值就近归位后步进', () => {
    expect(nextZoomLevel(1.1, 1)).toBe(1.25)
    expect(nextZoomLevel(1.1, -1)).toBe(1)
  })
})

describe('anchoredScroll', () => {
  it('光标下内容位置缩放前后不变：scroll\' = (scroll+cursor)×z1/z0 - cursor', () => {
    expect(anchoredScroll(100, 200, 1, 2)).toBe(400)
    expect(anchoredScroll(400, 200, 2, 1)).toBe(100)
    expect(anchoredScroll(0, 0, 1, 1.5)).toBe(0)
  })
})

describe('geometry zoom', () => {
  it('行列尺寸 ×zoom，前缀和一致', () => {
    let s = SheetData.create({ rowCount: 3, colCount: 3 })
    s = s.setRowHeight(0, 40)
    s = s.setColWidth(1, 120)
    const g = new GridGeometry(s, 0, 0, undefined, undefined, 2)
    expect(g.rowHeight(0)).toBe(80)
    expect(g.rowHeight(1)).toBe(48)
    expect(g.colWidth(1)).toBe(240)
    expect(g.rowTop(2)).toBe(80 + 48)
    expect(g.contentWidth).toBe((96 + 120 + 96) * 2)
  })
  it('autoRows 一并缩放', () => {
    const s = SheetData.create({ rowCount: 2, colCount: 2 })
    const g = new GridGeometry(s, 0, 0, undefined, new Map([[0, 60]]), 2)
    expect(g.rowHeight(0)).toBe(120)
  })
})
