import { describe, expect, it } from 'vitest'
import { SheetData, Workbook } from '../src/core/model'
import { condFormatStyle, duplicateSets } from '../src/formula/condformat'
import { evaluatorFor } from '../src/formula/engine'

const mkEv = (cells: [number, number, string][]) => {
  let data = SheetData.create({ rowCount: 10, colCount: 10 })
  for (const [r, c, raw] of cells) data = data.setCell(r, c, { raw })
  const wb = Workbook.create({ rowCount: 10, colCount: 10 }).setSheet('s1', data)
  return evaluatorFor(wb)
}

const range = { sr: 0, sc: 0, er: 5, ec: 5 }

describe('condFormatStyle', () => {
  it('value gt 命中数字；未命中返回 undefined', () => {
    const ev = mkEv([[0, 0, '15'], [1, 0, '5']])
    const rules = [{ id: 'cf1', range, type: 'value' as const, op: 'gt' as const, v1: '10', style: { bg: '#f00' } }]
    expect(condFormatStyle(rules, 's1', 0, 0, ev, new Map())).toEqual({ bg: '#f00' })
    expect(condFormatStyle(rules, 's1', 1, 0, ev, new Map())).toBeUndefined()
  })
  it('value between 闭区间', () => {
    const ev = mkEv([[0, 0, '5'], [1, 0, '11']])
    const rules = [{ id: 'cf1', range, type: 'value' as const, op: 'between' as const, v1: '5', v2: '10', style: { bold: true } }]
    expect(condFormatStyle(rules, 's1', 0, 0, ev, new Map())).toEqual({ bold: true })
    expect(condFormatStyle(rules, 's1', 1, 0, ev, new Map())).toBeUndefined()
  })
  it('textContains 不区分大小写', () => {
    const ev = mkEv([[0, 0, 'Hello World']])
    const rules = [{ id: 'cf1', range, type: 'textContains' as const, text: 'world', style: { color: '#00f' } }]
    expect(condFormatStyle(rules, 's1', 0, 0, ev, new Map())).toEqual({ color: '#00f' })
  })
  it('duplicate：出现 ≥2 次的显示文本命中（不区分大小写）', () => {
    const ev = mkEv([[0, 0, 'a'], [1, 0, 'A'], [2, 0, 'b']])
    const rules = [{ id: 'cf1', range, type: 'duplicate' as const, style: { bg: '#0f0' } }]
    const dups = duplicateSets(rules, 's1', ev)
    expect(condFormatStyle(rules, 's1', 0, 0, ev, dups)).toEqual({ bg: '#0f0' })
    expect(condFormatStyle(rules, 's1', 1, 0, ev, dups)).toEqual({ bg: '#0f0' })
    expect(condFormatStyle(rules, 's1', 2, 0, ev, dups)).toBeUndefined()
  })
  it('优先级：数组前者优先', () => {
    const ev = mkEv([[0, 0, '15']])
    const rules = [
      { id: 'cf1', range, type: 'value' as const, op: 'gt' as const, v1: '10', style: { bg: '#f00' } },
      { id: 'cf2', range, type: 'value' as const, op: 'gt' as const, v1: '1', style: { bg: '#0f0' } },
    ]
    expect(condFormatStyle(rules, 's1', 0, 0, ev, new Map())).toEqual({ bg: '#f00' })
  })
  it('range 外不命中', () => {
    const ev = mkEv([[8, 8, '15']])
    const rules = [{ id: 'cf1', range, type: 'value' as const, op: 'gt' as const, v1: '10', style: { bg: '#f00' } }]
    expect(condFormatStyle(rules, 's1', 8, 8, ev, new Map())).toBeUndefined()
  })
})
