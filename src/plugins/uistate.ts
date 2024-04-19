// 临时 UI 态宿主：metaField 工厂（PluginKey + StateField 桥接 tr meta 通道，
// 配 addToHistory:false 不入 undo 栈）与 filterui（筛选箭头点击 → 开启下拉面板）。
import { EditorViewLike, HitResult, Plugin, PluginKey } from '../core/plugin'
import type { EditorView } from '../view/editorview'
import { filterDropdownKey } from '../view/types'

// 通用「meta 透传」state field：tr.setMeta(key, v) → field 值为 v；无 meta 保持原值
export function metaField<T>(key: PluginKey, initial: T): Plugin {
  return new Plugin({
    key,
    state: {
      init: (): T => initial,
      apply: (tr, value: T): T => {
        const v = tr.getMeta(key)
        return v === undefined ? value : (v as T)
      },
    },
  })
}

export function filterui(): Plugin {
  return new Plugin({
    props: {
      handleMouseDown(view: EditorViewLike, e: MouseEvent, hit: HitResult): boolean {
        if (hit.region !== 'filter') return false
        const v = view as EditorView
        v.dispatch(
          v.state.tr
            .setMeta(filterDropdownKey, { col: hit.col, x: e.clientX, y: e.clientY })
            .setMeta('addToHistory', false),
        )
        return true
      },
    },
  })
}
