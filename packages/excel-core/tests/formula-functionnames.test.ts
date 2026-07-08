import { describe, expect, it } from 'vitest'
import { functionNames } from '../src/formula/eval'

describe('functionNames', () => {
  it('聚合 AGGREGATES ∪ FUNCTIONS 键，全大写', () => {
    const names = functionNames()
    expect(names).toContain('SUM')
    expect(names).toContain('ABS')
    expect(names).toContain('COUNTIF')
    expect(names).toContain('AVERAGEIF')
    expect(names.every(n => n === n.toUpperCase())).toBe(true)
    expect(new Set(names).size).toBe(names.length) // 无重复
  })
  it('前缀匹配候选', () => {
    const names = functionNames().filter(n => n.startsWith('SU'))
    expect(names).toEqual(['SUM', 'SUMIF']) // AGGREGATES(SUM) 在前，FUNCTIONS(SUMIF) 在后
  })
})
