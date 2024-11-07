import { describe, expect, it } from 'vitest'
import { FilterOp, ValidationRule } from '../src/core/model'
import { notifyValidationReject, registerValidationNotice, validateInput } from '../src/core/validation'

const num = (op: FilterOp, v1: string, v2?: string): ValidationRule =>
  ({ id: 'v1', range: { sr: 0, sc: 0, er: 9, ec: 0 }, type: 'numRange', op, v1, ...(v2 !== undefined ? { v2 } : {}) })
const len = (op: FilterOp, v1: string): ValidationRule =>
  ({ id: 'v2', range: { sr: 0, sc: 0, er: 9, ec: 0 }, type: 'textLen', op, v1 })
const list = (items: string[]): ValidationRule =>
  ({ id: 'v3', range: { sr: 0, sc: 0, er: 9, ec: 0 }, type: 'list', items })

describe('validateInput', () => {
  it('numRange：区间内放行，区间外拒绝，非数字拒绝', () => {
    const rules = [num('between', '1', '9')]
    expect(validateInput(rules, 0, 0, '5')).toBeNull()
    expect(validateInput(rules, 0, 0, '1')).toBeNull()
    expect(validateInput(rules, 0, 0, '9')).toBeNull()
    expect(validateInput(rules, 0, 0, '10')).toMatch(/1.*9/)
    expect(validateInput(rules, 0, 0, 'abc')).toMatch(/数字/)
  })
  it('numRange 全 op', () => {
    expect(validateInput([num('gt', '5')], 0, 0, '6')).toBeNull()
    expect(validateInput([num('gt', '5')], 0, 0, '5')).not.toBeNull()
    expect(validateInput([num('gte', '5')], 0, 0, '5')).toBeNull()
    expect(validateInput([num('lt', '5')], 0, 0, '4')).toBeNull()
    expect(validateInput([num('lte', '5')], 0, 0, '6')).not.toBeNull()
    expect(validateInput([num('eq', '5')], 0, 0, '5')).toBeNull()
    expect(validateInput([num('neq', '5')], 0, 0, '5')).not.toBeNull()
    expect(validateInput([num('neq', '5')], 0, 0, '6')).toBeNull()
  })
  it('textLen：按字符数', () => {
    const rules = [len('lte', '3')]
    expect(validateInput(rules, 0, 0, 'abc')).toBeNull()
    expect(validateInput(rules, 0, 0, 'abcd')).not.toBeNull()
    expect(validateInput(rules, 0, 0, '中文测')).toBeNull()
  })
  it('list：trim + 不区分大小写', () => {
    const rules = [list(['Apple', 'Banana'])]
    expect(validateInput(rules, 0, 0, 'apple')).toBeNull()
    expect(validateInput(rules, 0, 0, ' Banana ')).toBeNull()
    expect(validateInput(rules, 0, 0, 'cherry')).not.toBeNull()
  })
  it('跳过：公式原文 / 空串 / 无命中规则 / 范围外坐标', () => {
    const rules = [num('between', '1', '9')]
    expect(validateInput(rules, 0, 0, '=A2')).toBeNull()
    expect(validateInput(rules, 0, 0, '')).toBeNull()
    expect(validateInput([], 0, 0, 'abc')).toBeNull()
    expect(validateInput(rules, 50, 50, 'abc')).toBeNull()
  })
  it('多规则：任一命中即校验（第一条命中的原因返回）', () => {
    const rules = [num('gt', '100'), list(['a'])]
    expect(validateInput(rules, 0, 0, '5')).not.toBeNull()
  })
})

describe('notifyValidationReject 注入', () => {
  it('注册后回调收到消息；未注册不抛', () => {
    const got: string[] = []
    registerValidationNotice((m) => got.push(m))
    notifyValidationReject('测试原因')
    expect(got).toEqual(['测试原因'])
  })
})
