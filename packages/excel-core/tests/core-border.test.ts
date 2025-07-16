import { describe, expect, it } from 'vitest'
import { computeBorderStyles } from '../src/core/border'
import { SheetData } from '../src/core/model'

const sheet = (): SheetData => SheetData.create({ rowCount: 10, colCount: 10 })
const edge = { style: 'thin' as const, color: '#ff0000' }
const range = { sr: 1, sc: 1, er: 2, ec: 2 }

describe('computeBorderStyles', () => {
  it('all：选区每格四边', () => {
    const entries = computeBorderStyles(sheet(), range, 'all', edge)
    expect(entries).toHaveLength(4)
    for (const e of entries) {
      expect(e.style?.border).toEqual({ top: edge, right: edge, bottom: edge, left: edge })
    }
  })
  it('outer：仅周界边', () => {
    const entries = computeBorderStyles(sheet(), range, 'outer', edge)
    const at = (r: number, c: number) => entries.find((e) => e.row === r && e.col === c)!
    expect(at(1, 1).style?.border).toEqual({ top: edge, left: edge })
    expect(at(2, 2).style?.border).toEqual({ bottom: edge, right: edge })
    expect(at(1, 2).style?.border).toEqual({ top: edge, right: edge })
  })
  it('inner：仅内部网格线；单格选区为空', () => {
    const entries = computeBorderStyles(sheet(), range, 'inner', edge)
    const at = (r: number, c: number) => entries.find((e) => e.row === r && e.col === c)!
    expect(at(1, 1).style?.border).toEqual({ bottom: edge, right: edge })
    expect(at(2, 2).style?.border).toEqual({ top: edge, left: edge })
    expect(at(1, 2).style?.border).toEqual({ bottom: edge, left: edge })
    const single = computeBorderStyles(sheet(), { sr: 0, sc: 0, er: 0, ec: 0 }, 'inner', edge)
    expect(single[0].style).toBeNull() // 无边框且格空 → style null
  })
  it('none：清除已有边框，保留其他样式', () => {
    const s = sheet().setCell(1, 1, { raw: 'x', style: { bold: true, border: { top: edge } } })
    const entries = computeBorderStyles(s, { sr: 1, sc: 1, er: 1, ec: 1 }, 'none', null)
    expect(entries[0].style).toEqual({ bold: true })
  })
  it('top/bottom/left/right：仅周界对应边', () => {
    const entries = computeBorderStyles(sheet(), range, 'bottom', edge)
    const at = (r: number, c: number) => entries.find((e) => e.row === r && e.col === c)!
    expect(at(2, 1).style?.border).toEqual({ bottom: edge })
    expect(at(1, 1).style?.border).toBeUndefined()
  })
  it('后写覆盖同侧边，未触及边保留', () => {
    const s = sheet().setCell(1, 1, {
      raw: '',
      style: { border: { left: { style: 'thick' }, right: { style: 'dotted', color: '#00ff00' } } },
    })
    const entries = computeBorderStyles(s, range, 'outer', edge)
    const at = entries.find((e) => e.row === 1 && e.col === 1)!
    // outer 触及 left/top（周界边）：同侧已声明边被后写覆盖
    expect(at.style?.border?.left).toEqual(edge)
    expect(at.style?.border?.top).toEqual(edge)
    // right 非 (1,1) 周界边，outer 未触及：原有声明保留
    expect(at.style?.border?.right).toEqual({ style: 'dotted', color: '#00ff00' })
  })
})
