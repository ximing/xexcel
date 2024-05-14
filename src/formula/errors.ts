// 公式值与错误值表示：lexer 错误字面量与求值期错误共用。
// 独立成模块以打破 criteria ↔ eval 循环 import（双方只需类型与守卫，无需彼此实现）。
export type FormulaValue = number | string | boolean | FormulaError
export interface FormulaError {
  // '#NULL!' '#DIV/0!' '#VALUE!' '#REF!' '#NAME?' '#NUM!' '#N/A' '#CYCLE!'（以 lexer 支持为准）
  error: string
}

export function isError(v: unknown): v is FormulaError {
  return typeof v === 'object' && v !== null && 'error' in v
}

export const err = (error: string): FormulaError => ({ error })
