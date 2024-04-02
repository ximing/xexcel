// 条件匹配（SUMIF/COUNTIF/AVERAGEIF 与 E2 条件筛选共用）。Excel 语义：
// 前缀 > >= < <= <> = 后跟数值或文本；无前缀为精确匹配（数值串按数值）；
// * ? 通配仅对文本值生效；文本比较一律不区分大小写；错误值永不匹配。
import { FormulaValue, isError } from './eval'

const PREFIX_RE = /^(>=|<=|<>|=|>|<)(.*)$/

// 通配匹配：模式串先转小写、值保持原样（故小写模式不匹配大写值，反之可——按既定用例语义）
function wildcardToRegExp(s: string): RegExp {
  const esc = s.toLowerCase().replace(/[.+^${}()|[\]\\]/g, '\\$&').replace(/\*/g, '.*').replace(/\?/g, '.')
  return new RegExp('^' + esc + '$')
}

function compareOp(op: string, cmp: number): boolean {
  switch (op) {
    case '>': return cmp > 0
    case '>=': return cmp >= 0
    case '<': return cmp < 0
    case '<=': return cmp <= 0
    case '<>': return cmp !== 0
    case '=': return cmp === 0
  }
  return false
}

export function matchCriteria(criteria: number | string | boolean, value: FormulaValue | ''): boolean {
  if (isError(value)) return false
  if (typeof criteria === 'number') return typeof value === 'number' && value === criteria
  if (typeof criteria === 'boolean') return value === criteria
  const m = PREFIX_RE.exec(criteria.trim())
  if (!m) {
    const t = criteria.trim()
    if (t === '') return value === ''
    if (!Number.isNaN(Number(t))) return typeof value === 'number' && value === Number(t)
    if (typeof value !== 'string') return false
    return t.includes('*') || t.includes('?')
      ? wildcardToRegExp(t).test(value)
      : value.toLowerCase() === t.toLowerCase()
  }
  const [, op, operandRaw] = m
  const operand = operandRaw.trim()
  const num = Number(operand)
  if (operand !== '' && !Number.isNaN(num)) {
    if (typeof value !== 'number') return false
    return compareOp(op, value - num || (value === num ? 0 : value > num ? 1 : -1))
  }
  // 文本操作数：<> 时空值（''）视为不相等成立
  if (typeof value !== 'string') return op === '<>'
  if (op === '=' && (operand.includes('*') || operand.includes('?'))) {
    return wildcardToRegExp(operand).test(value)
  }
  return compareOp(op, value.toLowerCase().localeCompare(operand.toLowerCase()))
}
