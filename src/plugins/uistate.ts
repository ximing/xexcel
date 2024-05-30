// 临时 UI 态宿主：metaField 工厂（PluginKey + StateField 桥接 tr meta 通道，
// 配 addToHistory:false 不入 undo 栈）与 filterui（筛选箭头点击 → 开启下拉面板）。
import { EditorViewLike, HitResult, Plugin, PluginKey } from '../core/plugin'
import { forEachSelectionRange } from '../core/selection'
import type { EditorView } from '../view/editorview'
import { contextMenuKey, filterDropdownKey, formatPainterKey, FormatPainterState } from '../view/types'

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

// 格式刷：激活后下一次 mouseup（选区已最终）把快照 style 整体刷到选区；
// 非锁定刷一次自动解除。注册顺序在 selection 之后，保证选区先更新。
export function painter(): Plugin {
  return new Plugin({
    props: {
      handleMouseUp(view: EditorViewLike, e: MouseEvent, hit: HitResult): boolean {
        const v = view as EditorView
        // 仅响应左键；右键菜单打开期间的 mouseup 穿透（window 监听）不得误刷
        if (e.button !== 0) return false
        if (v.state.getField(contextMenuKey)) return false
        const fp = v.state.getField(formatPainterKey) as FormatPainterState | null | undefined
        if (!fp || hit.region !== 'cell') return false
        const st = v.state
        const entries: { row: number; col: number; style: import('../core/model').CellStyle }[] = []
        forEachSelectionRange(st.selection, r => {
          for (let row = r.sr; row <= r.er; row++)
            for (let col = r.sc; col <= r.ec; col++) entries.push({ row, col, style: fp.style })
        })
        v.dispatch(
          st.tr.setCellStyles(entries).setMeta(formatPainterKey, fp.locked ? fp : null),
        )
        return true
      },
    },
  })
}
