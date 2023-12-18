import { describe, it, expect } from 'vitest'
import { Workbook } from '../src/core/model'
import { GridGeometry } from '../src/view/geometry'

const geomOf = (fr: number, fc: number) =>
  new GridGeometry(Workbook.create({ rowCount: 100, colCount: 26 }).activeSheet.setFrozen(fr, fc), fr, fc)
// 默认行高 24、列宽 96

describe('freeze geometry', () => {
  it('frozenWidth/frozenHeight', () => {
    const g = geomOf(2, 3)
    expect(g.frozenHeight).toBe(48)
    expect(g.frozenWidth).toBe(288)
  })
  it('cellAtContent：冻结区不吃 scroll', () => {
    const g = geomOf(1, 1)
    // 冻结行内：scrollY=500 仍命中行 0
    expect(g.cellAtContent(10, 10, 500, 500)).toEqual({ row: 0, col: 0 })
    // 冻结列内：scrollX=500 仍命中列 0；行仍吃 scrollY（rowAt(24+500+6)=22）
    // 注：brief 原断言 row:1 与其 cellAtContent 实现/象限渲染（冻结列象限 offY=-fh-scrollY）矛盾，按实现修正
    expect(g.cellAtContent(50, 30, 500, 500)).toEqual({ row: 22, col: 0 })
  })
  it('cellAtContent：主区 = 冻结尺寸 + scroll + 偏移', () => {
    const g = geomOf(1, 1)
    // frozenH=24, scrollY=0，内容 y=24 → 行 1；y=48 → 行 2
    expect(g.cellAtContent(100, 48, 0, 0).row).toBe(2)
    // scrollY=24 时同样 y=48 → 行 3
    expect(g.cellAtContent(100, 48, 0, 24).row).toBe(3)
    // frozenW=96, scrollX=96，内容 x=96 → 列 2
    expect(g.cellAtContent(96, 100, 96, 0).col).toBe(2)
  })
  it('无冻结时与旧行为一致', () => {
    const g = geomOf(0, 0)
    expect(g.cellAtContent(100, 50, 0, 0)).toEqual({ row: 2, col: 1 })
    expect(g.cellAtContent(100, 50, 96, 24)).toEqual({ row: 3, col: 2 })
  })
  it('冻结数越界（陈旧数据）→ getter 钳位，不出 NaN', () => {
    // 模拟 delete 前的旧数据：4 行的表带着 frozenRows=9
    const sheet4 = Workbook.create({ rowCount: 4, colCount: 3 }).activeSheet
    const g = new GridGeometry(sheet4, 9, 5)
    expect(g.frozenHeight).toBe(g.rowTop(4))
    expect(g.frozenWidth).toBe(g.colLeft(3))
    const a = g.cellAtContent(10, 10, 0, 0)
    expect(Number.isFinite(a.row)).toBe(true)
    expect(Number.isFinite(a.col)).toBe(true)
    const b = g.cellAtContent(500, 500, 100, 100)
    expect(Number.isFinite(b.row)).toBe(true)
    expect(Number.isFinite(b.col)).toBe(true)
  })
})
