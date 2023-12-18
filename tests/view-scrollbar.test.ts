import { describe, it, expect } from 'vitest'
import { hScrollbar, SB_SIZE, thumbHit, vScrollbar } from '../src/view/scrollbar'

// 视口 800×600，表头 48/26，内容 2400（100 行 × 24）高、2496（26 列 × 96）宽
describe('scrollbar geom', () => {
  it('垂直：滑块比例与位置', () => {
    const g = vScrollbar(2400, 0, 800, 600)!
    const trackLen = 600 - 26 - SB_SIZE
    expect(g.track).toEqual({ x: 800 - SB_SIZE, y: 26, w: SB_SIZE, h: trackLen })
    const thumbLen = Math.max(24, trackLen * ((600 - 26) / 2400))
    expect(g.thumb.y).toBe(26)
    expect(g.thumb.h).toBeCloseTo(thumbLen, 5)
    expect(g.ratio).toBeCloseTo((2400 - (600 - 26)) / (trackLen - thumbLen), 5)
  })
  it('滚动到一半 → 滑块居中', () => {
    const g0 = vScrollbar(2400, 0, 800, 600)!
    const g = vScrollbar(2400, (2400 - 574) / 2, 800, 600)!
    expect(g.thumb.y).toBeCloseTo(26 + (g.track.h - g.thumb.h) / 2, 5)
    expect(g0.maxScroll).toBe(2400 - 574)
  })
  it('水平：轨道在底部', () => {
    const g = hScrollbar(2496, 0, 800, 600)!
    expect(g.track).toEqual({ x: 48, y: 600 - SB_SIZE, w: 800 - 48 - SB_SIZE, h: SB_SIZE })
  })
  it('内容不足一屏 → null', () => {
    expect(vScrollbar(500, 0, 800, 600)).toBeNull()
    expect(hScrollbar(700, 0, 800, 600)).toBeNull()
  })
  it('thumbHit', () => {
    const g = vScrollbar(2400, 0, 800, 600)!
    expect(thumbHit(g, 795, 30)).toBe(true)
    expect(thumbHit(g, 795, 500)).toBe(false)
  })
})
