// 数字格式显示：numfmt（ECMA-376 格式串）封装 + 工具栏小数位调整。
// 纯函数，零 DOM 依赖。
import { format as numfmtFormat } from 'numfmt'

// 按格式串渲染数字；非法格式串 → null（调用侧回退 formatNumber）
export function formatValue(code: string, value: number): string | null {
  try {
    return numfmtFormat(code, value)
  } catch {
    return null
  }
}

// 增/减小数位：只操作第一段（正数段）的小数部分，面向工具栏生成的简单格式串。
// 无 numFmt 时以 '#,##0.00' 为底（spec §3）。
export function adjustDecimals(code: string | undefined, delta: 1 | -1): string {
  const base = (code ?? '#,##0.00').split(';')[0]
  const dot = base.indexOf('.')
  if (dot < 0) {
    if (delta < 0) return base
    // 在末尾数字占位符后、非数字前缀（如 %）前插入小数位
    const m = /[0#?]([^0#?]*)$/.exec(base)
    if (!m) return base
    const tail = m[1]
    return base.slice(0, base.length - tail.length) + '.0' + tail
  }
  const after = base.slice(dot + 1)
  const m = /^([0#?]+)(.*)$/.exec(after)
  if (!m) return base
  const decimals = m[1].length
  const next = Math.max(0, decimals + delta)
  const tail = m[2]
  if (next === 0) return base.slice(0, dot) + tail
  return base.slice(0, dot + 1) + '0'.repeat(next) + tail
}
