// 键盘映射插件（handleKeyDown 来自 proxy textarea）。职责：
// - Arrow 移动 focus 并塌缩单格；Shift+Arrow 扩展 focus；Tab/Shift+Tab 右/左；Enter/Shift+Enter 下/上
// - Ctrl/Cmd+A → selectAll；Delete/Backspace → clearSelection
// - Ctrl/Cmd+Z → undo；Ctrl/Cmd+Shift+Z 或 Ctrl+Y → redo
// - F2 → openEditor；可打印单字符（无 Ctrl/Cmd/Alt）→ 清 proxy 后 openEditor(focus, key)
// 命中的按键一律 preventDefault 并返回 true；未命中返回 false（不拦截 copy/paste 等）。
// 移动类 dispatch 带 scrollIntoView()。
import { clearSelection, selectAll } from '../core/commands'
import { redo, undo } from '../core/history'
import { EditorViewLike, Plugin } from '../core/plugin'
import { singleCell } from '../core/selection'
import { isEditing, openEditor } from '../view/editbox'
import type { EditorView } from '../view/editorview'

export function keymap(): Plugin {
  return new Plugin({
    props: {
      handleKeyDown(view: EditorViewLike, e: KeyboardEvent): boolean {
        if (isEditing()) return false // 编辑器有自己的 keydown（stopPropagation），此处兜底
        const v = view as EditorView
        const state = v.state
        const sheet = state.activeSheet
        const sel = state.selection
        const mod = e.ctrlKey || e.metaKey

        const move = (dr: number, dc: number, extend: boolean): void => {
          const focus = {
            row: Math.max(0, Math.min(sel.focus.row + dr, sheet.rowCount - 1)),
            col: Math.max(0, Math.min(sel.focus.col + dc, sheet.colCount - 1)),
          }
          const m = sheet.mergeAt(focus.row, focus.col)
          if (m) {
            focus.row = m.sr
            focus.col = m.sc
          }
          v.dispatch(
            state.tr
              .setSelection(extend ? { anchor: sel.anchor, focus } : singleCell(focus.row, focus.col))
              .scrollIntoView(),
          )
        }

        switch (e.key) {
          case 'ArrowUp':
            move(-1, 0, e.shiftKey)
            break
          case 'ArrowDown':
            move(1, 0, e.shiftKey)
            break
          case 'ArrowLeft':
            move(0, -1, e.shiftKey)
            break
          case 'ArrowRight':
            move(0, 1, e.shiftKey)
            break
          case 'Tab':
            move(0, e.shiftKey ? -1 : 1, false)
            break
          case 'Enter':
            move(e.shiftKey ? -1 : 1, 0, false)
            break
          case 'Delete':
          case 'Backspace':
            clearSelection(state, (tr) => v.dispatch(tr))
            break
          case 'F2':
            openEditor(v, sel.focus)
            break
          default: {
            const k = e.key.toLowerCase()
            if (mod && k === 'a') {
              selectAll(state, (tr) => v.dispatch(tr))
            } else if (mod && k === 'z' && !e.shiftKey) {
              undo(state, (tr) => v.dispatch(tr))
            } else if ((mod && k === 'z' && e.shiftKey) || (e.ctrlKey && k === 'y')) {
              redo(state, (tr) => v.dispatch(tr))
            } else if (e.key.length === 1 && !mod && !e.altKey) {
              v.clearProxy() // 清掉 proxy 可能积累的脏值，再开编辑器
              openEditor(v, sel.focus, e.key)
            } else {
              return false
            }
          }
        }
        e.preventDefault()
        return true
      },
    },
  })
}
