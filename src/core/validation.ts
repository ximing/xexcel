// 数据验证：交互录入校验纯函数 + 拒绝通知注入点（view 层不能 import app，仿 registerStructureCascade）。
import { CellRange } from './addr'
import { FilterOp, ValidationRule } from './model'

const numCmp = (op: FilterOp, v: number, v1: number, v2?: number): boolean => {
  switch (op) {
    case 'eq': return v === v1
    case 'neq': return v !== v1
    case 'gt': return v > v1
    case 'gte': return v >= v1
    case 'lt': return v < v1
    case 'lte': return v <= v1
    case 'between': return v2 !== undefined && v >= v1 && v <= v2
    default: return true // 文本 op 不用于数值验证
  }
}

const inRange = (r: CellRange, row: number, col: number): boolean =>
  row >= r.sr && row <= r.er && col >= r.sc && col <= r.ec

const OP_LABEL: Partial<Record<FilterOp, string>> = {
  eq: '等于', neq: '不等于', gt: '大于', gte: '大于等于', lt: '小于', lte: '小于等于', between: '介于',
}

function check(rule: ValidationRule, raw: string): string | null {
  if (rule.type === 'list') {
    const t = raw.trim().toLowerCase()
    if (rule.items.some((i) => i.trim().toLowerCase() === t)) return null
    return `输入值须在序列内：${rule.items.join(', ')}`
  }
  if (rule.type === 'numRange') {
    const v = Number(raw.trim())
    if (raw.trim() === '' || !Number.isFinite(v)) return '请输入数字'
    const v1 = Number(rule.v1)
    const v2 = rule.v2 !== undefined ? Number(rule.v2) : undefined
    if (numCmp(rule.op, v, v1, v2)) return null
    return rule.op === 'between'
      ? `请输入${OP_LABEL[rule.op]} ${rule.v1} 与 ${rule.v2} 之间的数字`
      : `请输入${OP_LABEL[rule.op] ?? rule.op} ${rule.v1} 的数字`
  }
  // textLen：按字符数
  const n = raw.length
  const v1 = Number(rule.v1)
  const v2 = rule.v2 !== undefined ? Number(rule.v2) : undefined
  if (numCmp(rule.op, n, v1, v2)) return null
  return rule.op === 'between'
    ? `文本长度须介于 ${rule.v1} 与 ${rule.v2} 之间`
    : `文本长度须${OP_LABEL[rule.op] ?? rule.op} ${rule.v1}`
}

// 命中任一规则且非法 → 原因串；公式原文/空串跳过（清格/公式不校验）
export function validateInput(
  validations: ValidationRule[],
  row: number,
  col: number,
  raw: string,
): string | null {
  if (raw === '' || raw.startsWith('=')) return null
  for (const rule of validations) {
    if (!inRange(rule.range, row, col)) continue
    const reason = check(rule, raw)
    if (reason) return reason
  }
  return null
}

// ---- 拒绝通知注入（app 层注册 showNotice；未注册静默丢弃） ----
let noticeFn: ((msg: string) => void) | null = null
export function registerValidationNotice(fn: (msg: string) => void): void {
  noticeFn = fn
}
export function notifyValidationReject(msg: string): void {
  noticeFn?.(msg)
}
