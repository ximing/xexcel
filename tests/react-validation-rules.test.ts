import { describe, expect, it } from 'vitest'
import { ValidationRule } from '@gmi/excel-core'
import {
  describeRule,
  nextValidationId,
  normalizeRuleRange,
  parseItems,
  ruleInvalid,
} from '../src/react/validationRules'

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
  it('ruleInvalid：between 上下界倒置（v1>v2）非法', () => {
    expect(ruleInvalid({ ...num, v1: '9', v2: '1' }, 'A1:A10')).toBe(true)
    expect(ruleInvalid({ ...len, op: 'between', v1: '10', v2: '3' }, 'B2:B6')).toBe(true)
    // 等界合法；非 between 不受影响
    expect(ruleInvalid({ ...num, v1: '5', v2: '5' }, 'A1:A10')).toBe(false)
    expect(ruleInvalid({ ...num, op: 'gt', v1: '9', v2: undefined }, 'A1:A10')).toBe(false)
    // 不可解析值仍由既有分支挡下（不误判为合法倒置）
    expect(ruleInvalid({ ...num, v1: 'x', v2: '1' }, 'A1:A10')).toBe(true)
  })
  it('normalizeRuleRange：越界 clamp 到表界，界内原样', () => {
    const sheet = { rowCount: 100, colCount: 26 }
    expect(normalizeRuleRange({ sr: 0, sc: 0, er: 999, ec: 0 }, sheet)).toEqual({ sr: 0, sc: 0, er: 99, ec: 0 })
    expect(normalizeRuleRange({ sr: 0, sc: 20, er: 5, ec: 99 }, sheet)).toEqual({ sr: 0, sc: 20, er: 5, ec: 25 })
    expect(normalizeRuleRange(num.range, sheet)).toEqual(num.range)
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
