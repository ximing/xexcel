// 单元格编辑器：DOM textarea 覆盖层（不走 Konva，输入法体验与原生一致）。
// Esc 取消；Enter 提交并下移；Tab 提交并右移；blur 提交。编辑期间焦点在编辑器上，
// proxy 处于 blur，插件 keydown 不触发。
// F5：编辑 =… 公式时画布高亮被引区域（refHighlightKey meta）；函数名补全下拉。
import type { CellAddr } from '../core/addr'
import { singleCell } from '../core/selection'
import { notifyValidationReject, validateInput } from '../core/validation'
import { functionNames } from '../formula/eval'
import { normalizedCell } from '../formula/input'
import { completionCandidates } from '../formula/rangeRefs'
import type { EditorView } from './editorview'
import { THEME } from './theme'
import { refHighlightKey } from './types'

interface EditSession {
  view: EditorView
  addr: CellAddr
  el: HTMLTextAreaElement
  done: boolean // 防 blur 与按键路径重复关闭
  dropdown: HTMLDivElement | null // 函数名补全下拉（null=未显示）
  candidates: string[] // 当前补全候选
  selIndex: number // 当前选中候选下标（-1=无）
}

let session: EditSession | null = null

export function isEditing(): boolean {
  return session !== null
}

export function openEditor(view: EditorView, addr: CellAddr, initialText?: string): void {
  if (session) {
    closeEditor(true)
    if (session) return // 验证拒绝：原会话存活，忽略重开（防僵尸双编辑器）
  }
  const cell = view.state.activeSheet.getCell(addr.row, addr.col)
  const r = view.cellViewportRect(addr.row, addr.col)
  // 编辑器字号 = 格样式字号 × zoom，与画布几何缩放一致
  const fs = (cell?.style?.fontSize ?? 13) * view.zoom()
  const el = document.createElement('textarea')
  el.className = 'xcell-editor'
  Object.assign(el.style, {
    position: 'absolute',
    left: `${r.x - 1}px`,
    top: `${r.y - 1}px`,
    minWidth: `${r.w + 2}px`,
    minHeight: `${r.h + 2}px`,
    boxSizing: 'border-box',
    border: `2px solid ${THEME.primary}`,
    outline: 'none',
    margin: '0',
    padding: '1px 5px',
    font: `${fs}px -apple-system, BlinkMacSystemFont, "Segoe UI", Roboto, "Helvetica Neue", Arial, sans-serif`,
    resize: 'none',
    overflow: 'hidden',
    zIndex: '10',
    background: THEME.surface,
    color: THEME.ink,
  })
  el.value = initialText ?? cell?.raw ?? ''
  // wrap 格：编辑框随内容向下增长（不小于行高）
  const grow = cell?.style?.wrap
    ? (): void => {
        el.style.height = `${Math.max(r.h + 2, el.scrollHeight + 2)}px`
      }
    : null
  if (grow) el.addEventListener('input', grow)
  el.addEventListener('keydown', onEditorKeyDown)
  el.addEventListener('blur', onEditorBlur)
  el.addEventListener('input', onInput)
  view.dom.appendChild(el)
  // 首次测量须在挂载后：detached 节点 scrollHeight 恒为 0
  grow?.()
  session = { view, addr, el, done: false, dropdown: null, candidates: [], selIndex: -1 }
  el.focus()
  el.setSelectionRange(el.value.length, el.value.length)
  // 以 = 开头的初始文本（如按键 = 入口）即触发高亮 + 补全
  onInput()
}

export function closeEditor(commit: boolean): void {
  finish(commit)
}

// 关闭会话；commit 时提交文本。仅 Enter/Tab 显式传 next 时移动选区；
// blur 提交（如点击其他单元格）不动选区——选区已被 mousedown 的 tr 更新为目标格。
function finish(commit: boolean, next?: CellAddr): void {
  const s = session
  if (!s || s.done) return
  // 数据验证：提交前先校验，拒绝则不拆会话（编辑框保持打开）
  if (commit) {
    const reason = validateInput(s.view.state.activeSheet.validations, s.addr.row, s.addr.col, s.el.value)
    if (reason) {
      notifyValidationReject(reason)
      // blur 路径：mousedown 的选区 tr 已先提交到目标格，须把选区派回编辑格（对齐 Excel 锚定）。
      // 选区 tr 无 steps 本就不入历史，无需 addToHistory meta（同 selection 插件惯例）
      s.view.dispatch(s.view.state.tr.setSelection(singleCell(s.addr.row, s.addr.col)))
      s.el.focus() // blur 路径拒绝后焦点回到编辑框（对齐 Excel）
      return
    }
  }
  s.done = true
  session = null
  const text = s.el.value
  closeCompletion(s)
  s.el.removeEventListener('keydown', onEditorKeyDown)
  s.el.removeEventListener('blur', onEditorBlur)
  s.el.removeEventListener('input', onInput)
  s.el.remove()
  // 清画布引用高亮（非文档态 meta，不入 undo）
  s.view.dispatch(s.view.state.tr.setMeta(refHighlightKey, null).setMeta('addToHistory', false))
  if (commit) {
    const tr = s.view.state.tr
    const oldCell = s.view.state.activeSheet.getCell(s.addr.row, s.addr.col)
    const nextCell = normalizedCell(text, oldCell)
    const changed =
      nextCell.raw !== (oldCell?.raw ?? '') ||
      nextCell.style?.numFmt !== oldCell?.style?.numFmt
    if (changed) tr.setCell(s.addr.row, s.addr.col, nextCell.raw, nextCell.style)
    if (next) tr.setSelection(singleCell(next.row, next.col)).scrollIntoView()
    if (tr.steps.length > 0 || tr.selection) s.view.dispatch(tr)
  }
  s.view.focus()
}

// 编辑输入：更新画布引用高亮 meta + 维护函数名补全下拉
function onInput(): void {
  const s = session
  if (!s) return
  const t = s.el.value
  s.view.dispatch(
    s.view.state.tr.setMeta(refHighlightKey, t.startsWith('=') ? t : null).setMeta('addToHistory', false),
  )
  updateCompletion(s, t)
}

// 函数名补全：= 后末尾标识符 token 前缀匹配 functionNames()，≤8 个
function updateCompletion(s: EditSession, text: string): void {
  const candidates = completionCandidates(text, functionNames())
  s.candidates = candidates
  if (candidates.length === 0) {
    closeCompletion(s)
    return
  }
  if (!s.dropdown) {
    const dd = document.createElement('div')
    dd.className = 'xcell-autocomplete'
    Object.assign(dd.style, {
      position: 'absolute',
      left: `${s.el.offsetLeft}px`,
      top: `${s.el.offsetTop + s.el.offsetHeight}px`,
      minWidth: `${s.el.offsetWidth}px`,
      background: THEME.surface,
      border: `1px solid ${THEME.lineStrong}`,
      borderTop: 'none',
      zIndex: '11',
      maxHeight: '200px',
      overflowY: 'auto',
      font: '13px -apple-system, BlinkMacSystemFont, "Segoe UI", Roboto, "Helvetica Neue", Arial, sans-serif',
      boxShadow: '0 2px 6px rgba(0,0,0,0.12)',
    })
    s.view.dom.appendChild(dd)
    s.dropdown = dd
  }
  s.selIndex = 0
  renderCompletion(s)
}

function renderCompletion(s: EditSession): void {
  const dd = s.dropdown
  if (!dd) return
  dd.innerHTML = ''
  s.candidates.forEach((name, i) => {
    const item = document.createElement('div')
    item.textContent = name
    Object.assign(item.style, {
      padding: '2px 8px',
      cursor: 'pointer',
      background: i === s.selIndex ? THEME.primarySoft : THEME.surface,
      color: i === s.selIndex ? THEME.primary : THEME.ink,
    })
    // mousedown 而非 click：在 blur 前 intercept 并阻止 textarea 失焦
    item.addEventListener('mousedown', (e) => {
      e.preventDefault()
      s.selIndex = i
      acceptCompletion(s)
    })
    dd.appendChild(item)
  })
}

// 接受补全：末尾标识符 token 替换为 NAME(
function acceptCompletion(s: EditSession): void {
  if (!s.candidates.length || s.selIndex < 0 || s.selIndex >= s.candidates.length) {
    closeCompletion(s)
    return
  }
  const name = s.candidates[s.selIndex]
  const text = s.el.value
  const eq = text.lastIndexOf('=')
  if (eq < 0) {
    closeCompletion(s)
    return
  }
  const m = /[A-Za-z]+$/.exec(text.slice(eq + 1))
  if (!m) {
    closeCompletion(s)
    return
  }
  s.el.value = text.slice(0, eq + 1 + m.index) + name + '('
  s.el.setSelectionRange(s.el.value.length, s.el.value.length)
  closeCompletion(s)
  // 同步画布高亮 meta 到接受后的文本
  s.view.dispatch(
    s.view.state.tr.setMeta(refHighlightKey, s.el.value.startsWith('=') ? s.el.value : null).setMeta('addToHistory', false),
  )
}

function closeCompletion(s: EditSession): void {
  if (s.dropdown) {
    s.dropdown.remove()
    s.dropdown = null
  }
  s.candidates = []
  s.selIndex = -1
}

function onEditorKeyDown(e: KeyboardEvent): void {
  const s = session
  if (!s) return
  e.stopPropagation()
  if (e.isComposing) return // 输入法组合中不响应提交键
  // 补全下拉可见时优先处理导航/接受/关闭
  if (s.dropdown) {
    if (e.key === 'ArrowDown') {
      e.preventDefault()
      s.selIndex = Math.min(s.selIndex + 1, s.candidates.length - 1)
      renderCompletion(s)
      return
    }
    if (e.key === 'ArrowUp') {
      e.preventDefault()
      s.selIndex = Math.max(s.selIndex - 1, 0)
      renderCompletion(s)
      return
    }
    if (e.key === 'Tab' || (e.key === 'Enter' && !e.shiftKey)) {
      e.preventDefault()
      acceptCompletion(s)
      return
    }
    if (e.key === 'Escape') {
      e.preventDefault()
      closeCompletion(s)
      return
    }
  }
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
