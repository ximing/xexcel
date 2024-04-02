import { describe, expect, it } from 'vitest'
import { matchCriteria } from '../src/formula/criteria'

describe('matchCriteria', () => {
  it('数值前缀比较', () => {
    expect(matchCriteria('>10', 11)).toBe(true)
    expect(matchCriteria('>10', 10)).toBe(false)
    expect(matchCriteria('>=10', 10)).toBe(true)
    expect(matchCriteria('<10', 9)).toBe(true)
    expect(matchCriteria('<=10', 10)).toBe(true)
    expect(matchCriteria('<>10', 9)).toBe(true)
    expect(matchCriteria('=10', 10)).toBe(true)
    expect(matchCriteria('>10', '11')).toBe(false) // 数值前缀只匹配数值值
  })

  it('文本前缀比较（不区分大小写）', () => {
    expect(matchCriteria('=abc', 'ABC')).toBe(true)
    expect(matchCriteria('<>abc', 'abd')).toBe(true)
    expect(matchCriteria('>b', 'c')).toBe(true)
    expect(matchCriteria('=abc', 123)).toBe(false)
  })

  it('无前缀：数值串按数值等值', () => {
    expect(matchCriteria('42', 42)).toBe(true)
    expect(matchCriteria('42', '42')).toBe(false)
  })

  it('无前缀：文本精确匹配（不区分大小写）', () => {
    expect(matchCriteria('hello', 'Hello')).toBe(true)
    expect(matchCriteria('hello', 'hello!')).toBe(false)
  })

  it('通配符 * 与 ?', () => {
    expect(matchCriteria('a*', 'abc')).toBe(true)
    expect(matchCriteria('a*', 'ABC')).toBe(false)
    expect(matchCriteria('A*', 'abc')).toBe(true) // 不区分大小写
    expect(matchCriteria('a?c', 'abc')).toBe(true)
    expect(matchCriteria('a?c', 'ac')).toBe(false)
    expect(matchCriteria('a*c', 'abbbc')).toBe(true)
  })

  it('空 criteria 匹配空单元格', () => {
    expect(matchCriteria('', '')).toBe(true)
    expect(matchCriteria('', 'x')).toBe(false)
  })

  it('数值 criteria 直接相等', () => {
    expect(matchCriteria(7, 7)).toBe(true)
    expect(matchCriteria(7, '7')).toBe(false)
  })

  it('错误值永不匹配', () => {
    expect(matchCriteria('*', { error: '#REF!' })).toBe(false)
    expect(matchCriteria('>0', { error: '#DIV/0!' })).toBe(false)
  })
})
