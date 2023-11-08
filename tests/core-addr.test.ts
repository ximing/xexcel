import { describe, it, expect } from 'vitest'
import { colName, parseColName, toA1, fromA1, parseRange, normalizeRange, rangeContains, rangeCellCount, clampRange, rangesEqual, wholeRange } from '../src/core/addr'

describe('addr', () => {
  it('colName/parseColName 互逆', () => {
    expect(colName(0)).toBe('A')
    expect(colName(25)).toBe('Z')
    expect(colName(26)).toBe('AA')
    expect(colName(51)).toBe('AZ')
    expect(colName(701)).toBe('ZZ')
    for (const n of [0, 1, 25, 26, 51, 52, 701, 702, 16383]) {
      expect(parseColName(colName(n))).toBe(n)
    }
  })
  it('parseColName 非法输入 → -1', () => {
    expect(parseColName('')).toBe(-1)
    expect(parseColName('1A')).toBe(-1)
    expect(parseColName('a b')).toBe(-1)
  })
  it('toA1/fromA1', () => {
    expect(toA1(0, 0)).toBe('A1')
    expect(toA1(0, 1)).toBe('B1')
    expect(toA1(9, 27)).toBe('AB10')
    expect(fromA1('B1')).toEqual({ row: 0, col: 1 })
    expect(fromA1('ab10')).toEqual({ row: 9, col: 27 })
    expect(fromA1('A0')).toBeNull()
    expect(fromA1('A')).toBeNull()
    expect(fromA1('1')).toBeNull()
    expect(fromA1('')).toBeNull()
  })
  it('range 工具', () => {
    expect(normalizeRange({ sr: 5, sc: 3, er: 2, ec: 1 })).toEqual({ sr: 2, sc: 1, er: 5, ec: 3 })
    expect(parseRange('B2:D5')).toEqual({ sr: 1, sc: 1, er: 4, ec: 3 })
    expect(parseRange('B2')).toEqual({ sr: 1, sc: 1, er: 1, ec: 1 })
    expect(parseRange('D5:B2')).toEqual({ sr: 1, sc: 1, er: 4, ec: 3 })
    expect(parseRange('bad')).toBeNull()
    expect(rangeContains({ sr: 1, sc: 1, er: 4, ec: 3 }, 4, 3)).toBe(true)
    expect(rangeContains({ sr: 1, sc: 1, er: 4, ec: 3 }, 5, 3)).toBe(false)
    expect(rangeCellCount({ sr: 1, sc: 1, er: 4, ec: 3 })).toBe(12)
    expect(clampRange({ sr: 0, sc: 0, er: 99, ec: 99 }, 10, 5)).toEqual({ sr: 0, sc: 0, er: 9, ec: 4 })
    expect(rangesEqual({ sr: 1, sc: 1, er: 2, ec: 2 }, { sr: 1, sc: 1, er: 2, ec: 2 })).toBe(true)
    expect(wholeRange(100, 26)).toEqual({ sr: 0, sc: 0, er: 99, ec: 25 })
  })
})
