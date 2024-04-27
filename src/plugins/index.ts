// 内建交互插件组装。builtinPlugins() 顺序：selection → fillhandle → clipboard → keymap；
// 返回数组不含 history——由调用方插入最前（如 [history(), ...builtinPlugins()]）。
import { Plugin } from '../core/plugin'
import { filterDropdownKey, findBarKey, formatPainterKey } from '../view/types'
import { clipboard } from './clipboard'
import { fillhandle } from './fillhandle'
import { keymap } from './keymap'
import { selection } from './selection'
import { filterui, metaField, painter } from './uistate'

export { clipboard } from './clipboard'
export { fillhandle } from './fillhandle'
export { keymap } from './keymap'
export { selection } from './selection'
// fillPreviewKey/filterDropdownKey 定义在 src/view/types.ts（避免 layers ↔ plugins 循环依赖），此处再导出
export { fillPreviewKey, filterDropdownKey, formatPainterKey } from '../view/types'
export { filterui, metaField, painter } from './uistate'

export function builtinPlugins(): Plugin[] {
  return [
    selection(),
    fillhandle(),
    clipboard(),
    keymap(),
    metaField(filterDropdownKey, null),
    metaField(findBarKey, false),
    filterui(),
    metaField(formatPainterKey, null),
    painter(),
  ]
}
