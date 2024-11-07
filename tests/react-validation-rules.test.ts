import { describe, expect, it } from 'vitest'
import { ValidationRule } from '../src/core/model'
import { describeRule, nextValidationId, parseItems, ruleInvalid } from '../src/react/validationRules'

const num: ValidationRule = { id: 'v1', range: { sr: 0, sc: 0, er: 9, ec: 0 }, type: 'numRange', op: 'between', v1: '1', v2: '9' }
const len: ValidationRule = { id: 'v2', range: { sr: 1, sc: 1, er: 5, ec: 1 }, type: 'textLen', op: 'lte', v1: '3' }
const list: ValidationRule = { id: 'v3', range: { sr: 0, sc: 2, er: 5, ec: 2 }, type: 'list', items: ['a', 'b'] }

describe('validationRules 纯逻辑', () => {
  it('nextValidationId 递增跳过占用', () => {
    expect(nextValidationId([num, list])).toBe('v4')
    expect(nextValidationId([])).toBe('v1')
  })
  it('ruleInvalid：范围/数字/between/items', () => {
    expect(ruleInvalid(num, 'A1:A10')).toBe(false)
    expect(ruleInvalid(num, 'bad')).toBe(true)
    expect(ruleInvalid({ ...num, v1: '' }, 'A1:A10')).toBe(true)
    expect(ruleInvalid({ ...num, v2: '' }, 'A1:A10')).toBe(true)
    expect(ruleInvalid({ ...list, items: [] }, 'C1:C6')).toBe(true)
  })
  it('parseItems：逗号/中文逗号/trim/去空', () => {
    expect(parseItems('a, b，c,,')).toEqual(['a', 'b', 'c'])
    expect(parseItems('')).toEqual([])
  })
  it('describeRule', () => {
    expect(describeRule(num)).toBe('A1:A10 数值介于 1 与 9')
    expect(describeRule(len)).toBe('B2:B6 文本长度 <= 3')
    expect(describeRule(list)).toBe('C1:C6 序列：a, b')
  })
})
