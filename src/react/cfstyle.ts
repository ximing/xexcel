// 条件格式草稿的样式开关纯逻辑（CondFormatDialog 用，独立模块便于单测）。
import { CFStyle } from '../core/model'

export type CFToggleKey = 'bold' | 'italic' | 'underline' | 'strikethrough'

// 切换布尔样式键：开态写 true，关态剔除该键（不留显式 undefined own property，
// 否则渲染合并 {...cell.style, ...cf} 时 undefined 会覆盖静态样式）
export function toggleCFStyle(style: CFStyle, key: CFToggleKey): CFStyle {
  if (style[key]) {
    const next = { ...style }
    delete next[key]
    return next
  }
  return { ...style, [key]: true }
}
