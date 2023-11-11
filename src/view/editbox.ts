// 单元格编辑器：DOM textarea 覆盖层（不走 Konva，输入法体验与原生一致）。
// Esc 取消；Enter 提交并下移；Tab 提交并右移；blur 提交。编辑期间焦点在编辑器上，
// proxy 处于 blur，插件 keydown 不触发。
import type { CellAddr } from '../core/addr'
import { singleCell } from '../core/selection'
import type { EditorView } from './editorview'

interface EditSession {
  view: EditorView
  addr: CellAddr
  el: HTMLTextAreaElement
  done: boolean // 防 blur 与按键路径重复关闭
}

let session: EditSession | null = null

export function isEditing(): boolean {
  return session !== null
}

export function openEditor(view: EditorView, addr: CellAddr, initialText?: string): void {
  if (session) closeEditor(true)
  const cell = view.state.activeSheet.getCell(addr.row, addr.col)
  const r = view.cellViewportRect(addr.row, addr.col)
  const el = document.createElement('textarea')
  el.className = 'xcell-editor'
  Object.assign(el.style, {
    position: 'absolute',
    left: `${r.x - 1}px`,
    top: `${r.y - 1}px`,
    minWidth: `${r.w + 2}px`,
    minHeight: `${r.h + 2}px`,
    boxSizing: 'border-box',
    border: '2px solid #1a73e8',
    outline: 'none',
    margin: '0',
    padding: '1px 5px',
    font: '13px -apple-system, BlinkMacSystemFont, "Segoe UI", Roboto, "Helvetica Neue", Arial, sans-serif',
    resize: 'none',
    overflow: 'hidden',
    zIndex: '10',
    background: '#ffffff',
    color: '#202124',
  })
  el.value = initialText ?? cell?.raw ?? ''
  el.addEventListener('keydown', onEditorKeyDown)
  el.addEventListener('blur', onEditorBlur)
  view.dom.appendChild(el)
  session = { view, addr, el, done: false }
  el.focus()
  el.setSelectionRange(el.value.length, el.value.length)
}

export function closeEditor(commit: boolean): void {
  finish(commit)
}

// 关闭会话；commit 时提交文本。仅 Enter/Tab 显式传 next 时移动选区；
// blur 提交（如点击其他单元格）不动选区——选区已被 mousedown 的 tr 更新为目标格。
function finish(commit: boolean, next?: CellAddr): void {
  const s = session
  if (!s || s.done) return
  s.done = true
  session = null
  const text = s.el.value
  s.el.removeEventListener('keydown', onEditorKeyDown)
  s.el.removeEventListener('blur', onEditorBlur)
  s.el.remove()
  if (commit) {
    const tr = s.view.state.tr
    const old = s.view.state.activeSheet.getCell(s.addr.row, s.addr.col)?.raw ?? ''
    if (text !== old) tr.setCell(s.addr.row, s.addr.col, text)
    if (next) tr.setSelection(singleCell(next.row, next.col)).scrollIntoView()
    if (tr.steps.length > 0 || tr.selection) s.view.dispatch(tr)
  }
  s.view.focus()
}

function onEditorKeyDown(e: KeyboardEvent): void {
  const s = session
  if (!s) return
  e.stopPropagation()
  if (e.isComposing) return // 输入法组合中不响应提交键
  const sheet = s.view.state.activeSheet
  if (e.key === 'Escape') {
    e.preventDefault()
    finish(false)
  } else if (e.key === 'Enter' && !e.shiftKey) {
    e.preventDefault()
    finish(true, { row: Math.min(s.addr.row + 1, sheet.rowCount - 1), col: s.addr.col })
  } else if (e.key === 'Tab') {
    e.preventDefault()
    finish(true, { row: s.addr.row, col: Math.min(s.addr.col + 1, sheet.colCount - 1) })
  }
}

function onEditorBlur(): void {
  finish(true)
}
