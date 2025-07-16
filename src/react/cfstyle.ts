// 条件格式草稿的样式开关纯逻辑（CondFormatDialog 用，独立模块便于单测）。
// 注：本文件的十六进制色值是写入文档的样式数据（CF 预设与 color input 兜底），
// 非 UI chrome，不参与 token 体系，故豁免裸色守卫（tests/react-no-raw-color.test.ts 白名单）。
import { CFStyle } from '@gmi/excel-core'

// 新增规则时的默认样式（写入文档的样式数据）
export const DEFAULT_CF_STYLE: CFStyle = { bg: '#ffc7ce', color: '#9c0006' }

// color input 的 value 兜底：未设置对应样式时的中性色（input[type=color] 必须有合法值）
export const COLOR_INPUT_FALLBACK_TEXT = '#000000'
export const COLOR_INPUT_FALLBACK_BG = '#ffffff'

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
