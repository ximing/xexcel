// 键盘映射插件（handleKeyDown 来自 proxy textarea）。职责：
// - Arrow 移动 focus 并塌缩单格；Shift+Arrow 扩展 focus；Tab/Shift+Tab 右/左；Enter/Shift+Enter 下/上
// - Ctrl/Cmd+A → selectAll；Ctrl/Cmd+F → 打开查找栏；Delete/Backspace → clearSelection
// - Ctrl/Cmd+Z → undo；Ctrl/Cmd+Shift+Z 或 Ctrl+Y → redo
// - F2 → openEditor；可打印单字符（无 Ctrl/Cmd/Alt）→ 清 proxy 后 openEditor(focus, key)
// 命中的按键一律 preventDefault 并返回 true；未命中返回 false（不拦截 copy/paste 等）。
// 移动类 dispatch 带 scrollIntoView()。
import { CellAddr, CellRange } from '../core/addr'
import { clearSelection, selectAll } from '../core/commands'
import { redo, undo } from '../core/history'
import { EditorViewLike, Plugin } from '../core/plugin'
import { singleCell } from '../core/selection'
import { isEditing, openEditor } from '../view/editbox'
import type { EditorView } from '../view/editorview'
import { findBarKey, formatPainterKey } from '../view/types'

// 导航落点修正（导出供单测）：目标在合并区内时——
// 当前 focus 已是该合并区锚点（继续向外移动）→ 跳过整个合并区到远侧之外；
// 否则（从外部进入）→ 落在锚点。
export function navigateFocus(
  sheet: { mergeAt(row: number, col: number): CellRange | null; rowCount: number; colCount: number },
  current: CellAddr,
  dr: number,
  dc: number,
  isHidden: (row: number, col: number) => boolean = () => false,
): CellAddr {
  const clamp = (v: number, max: number): number => Math.max(0, Math.min(v, max))
  const target: CellAddr = {
    row: clamp(current.row + dr, sheet.rowCount - 1),
    col: clamp(current.col + dc, sheet.colCount - 1),
  }
  const m = sheet.mergeAt(target.row, target.col)
  let landed = target
  if (m) {
    const atAnchor = current.row === m.sr && current.col === m.sc
    landed = atAnchor
      ? {
          row: clamp(dr > 0 ? m.er + 1 : dr < 0 ? m.sr - 1 : target.row, sheet.rowCount - 1),
          col: clamp(dc > 0 ? m.ec + 1 : dc < 0 ? m.sc - 1 : target.col, sheet.colCount - 1),
        }
      : { row: m.sr, col: m.sc }
  }
  // 隐藏行列跳过：沿移动方向继续步进，直到可见格或边界（guard 防全隐藏死循环）
  let guard = 0
  const limit = Math.max(sheet.rowCount, sheet.colCount) + 1
  while (isHidden(landed.row, landed.col) && guard++ < limit) {
    const nr = clamp(landed.row + Math.sign(dr), sheet.rowCount - 1)
    const nc = clamp(landed.col + Math.sign(dc), sheet.colCount - 1)
    if (nr === landed.row && nc === landed.col) break // 已到边界（边界格本身隐藏则停留）
    landed = { row: nr, col: nc }
  }
  // 跳跃后落入合并区 → 锚点
  const m2 = sheet.mergeAt(landed.row, landed.col)
  if (m2) landed = { row: m2.sr, col: m2.sc }
  return landed
}

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
          const geom = v.geometry()
          const isHidden = (r: number, c: number): boolean => geom.rowHeight(r) === 0 || geom.colWidth(c) === 0
          const focus = navigateFocus(sheet, sel.focus, dr, dc, isHidden)
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
          case 'Escape': {
            const fp = state.getField(formatPainterKey)
            if (!fp) return false
            v.dispatch(state.tr.setMeta(formatPainterKey, null).setMeta('addToHistory', false))
            break
          }
          default: {
            const k = e.key.toLowerCase()
            if (mod && k === 'f') {
              v.dispatch(state.tr.setMeta(findBarKey, true).setMeta('addToHistory', false))
            } else if (mod && k === 'a') {
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
