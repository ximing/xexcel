// 数据验证对话框纯逻辑（组件外可单测）。
import { CellRange, clampRange, parseRangeA1, toA1 } from '@xexcel/core'
import { FilterOp, ValidationRule } from '@xexcel/core'
import type { SheetState } from '@xexcel/core'

export const VALIDATION_OPS: { op: FilterOp; label: string }[] = [
  { op: 'between', label: '介于' },
  { op: 'gt', label: '>' }, { op: 'gte', label: '>=' },
  { op: 'lt', label: '<' }, { op: 'lte', label: '<=' },
  { op: 'eq', label: '=' }, { op: 'neq', label: '≠' },
]

export function nextValidationId(rules: ValidationRule[]): string {
  let max = 0
  for (const r of rules) {
    const m = /^v(\d+)$/.exec(r.id)
    if (m) max = Math.max(max, parseInt(m[1], 10))
  }
  return `v${max + 1}`
}

export function rangeText(r: { sr: number; sc: number; er: number; ec: number }): string {
  return `${toA1(r.sr, r.sc)}:${toA1(r.er, r.ec)}`
}

// 非法判定：范围解析失败；numRange/textLen 的 v1 非数字；between 的 v2 非数字；list items 全空
// between 且 v1>v2（上下界倒置）亦非法
export function ruleInvalid(rule: ValidationRule, text: string): boolean {
  if (parseRangeA1(text) === null) return true
  if (rule.type === 'list') return rule.items.length === 0
  if (!Number.isFinite(Number(rule.v1)) || rule.v1.trim() === '') return true
  if (rule.op === 'between') {
    if ((rule.v2 ?? '').trim() === '' || !Number.isFinite(Number(rule.v2))) return true
    if (Number(rule.v1) > Number(rule.v2)) return true
  }
  return false
}

// 提交前把解析出的 range 收进表界（防越界规则写入 doc）
export function normalizeRuleRange(
  range: { sr: number; sc: number; er: number; ec: number },
  sheet: { rowCount: number; colCount: number },
): CellRange {
  return clampRange(range, sheet.rowCount, sheet.colCount)
}

// 序列输入框解析：逗号/中文逗号分隔，逐项 trim，去空
export function parseItems(text: string): string[] {
  return text.split(/[,，]/).map((s) => s.trim()).filter((s) => s !== '')
}

const OP_LABEL: Partial<Record<FilterOp, string>> = {
  eq: '=', neq: '≠', gt: '>', gte: '>=', lt: '<', lte: '<=', between: '介于',
}

// 规则列表行描述
export function describeRule(rule: ValidationRule): string {
  const r = rangeText(rule.range)
  if (rule.type === 'list') return `${r} 序列：${rule.items.join(', ')}`
  if (rule.type === 'numRange') {
    return rule.op === 'between'
      ? `${r} 数值介于 ${rule.v1} 与 ${rule.v2}`
      : `${r} 数值 ${OP_LABEL[rule.op]} ${rule.v1}`
  }
  return rule.op === 'between'
    ? `${r} 文本长度介于 ${rule.v1} 与 ${rule.v2}`
    : `${r} 文本长度 ${OP_LABEL[rule.op]} ${rule.v1}`
}

// 多区域选区下入口按钮禁用（单区域零回归）
export function canValidation(state: SheetState): boolean {
  return state.selection.ranges.length === 1
}

export function validationRejection(state: SheetState): string | null {
  return state.selection.ranges.length > 1 ? '数据验证仅支持单区域选择' : null
}
