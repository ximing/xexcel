import { describe, expect, it } from 'vitest'
import { SheetData } from '@xexcel/core'
import { edgeDash, edgeWeight, edgeWidth, resolveHEdge, resolveVEdge } from '../src/view/borders'

const mk = (): SheetData => SheetData.create({ rowCount: 5, colCount: 5 })

describe('共享边裁决', () => {
  it('权重高者胜（thick 压 thin）', () => {
    let s = mk()
    s = s.setCell(1, 1, { raw: '', style: { border: { right: { style: 'thick' } } } })
    s = s.setCell(1, 2, { raw: '', style: { border: { left: { style: 'thin' } } } })
    expect(resolveVEdge(s, 1, 2)).toEqual({ style: 'thick' })
  })
  it('同权重取左/上格', () => {
    let s = mk()
    s = s.setCell(1, 1, { raw: '', style: { border: { bottom: { style: 'thin', color: '#111' } } } })
    s = s.setCell(2, 1, { raw: '', style: { border: { top: { style: 'thin', color: '#222' } } } })
    expect(resolveHEdge(s, 2, 1)).toEqual({ style: 'thin', color: '#111' })
  })
  it('单边声明直接生效；无声明返回 undefined', () => {
    let s = mk()
    s = s.setCell(0, 0, { raw: '', style: { border: { left: { style: 'dashed' } } } })
    expect(resolveVEdge(s, 0, 0)).toEqual({ style: 'dashed' })
    expect(resolveHEdge(s, 4, 4)).toBeUndefined()
  })
})

describe('线型映射', () => {
  it('权重档位', () => {
    expect(edgeWeight({ style: 'hair' })).toBe(0.5)
    expect(edgeWeight({ style: 'dotted' })).toBe(1)
    expect(edgeWeight({ style: 'mediumDashed' })).toBe(2)
    expect(edgeWeight({ style: 'thick' })).toBe(3)
  })
  it('宽度与虚线', () => {
    expect(edgeWidth('medium')).toBe(2)
    expect(edgeWidth('double')).toBe(3)
    expect(edgeDash('dashed')).toEqual([6, 3])
    expect(edgeDash('thin')).toBeUndefined()
  })
})
