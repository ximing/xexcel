// 查找/替换栏：Ctrl+F 或工具栏开启（findBarKey state field）。
// 匹配显示文本与公式原文；逐个定位（选中 + ensureVisible，跨表切 active 不入 undo）；
// 替换仅改 raw（经 normalizedCell 归一）；全部替换跨表合并为一个事务。
import { useMemo, useState } from 'react'
import { singleCell } from '../core/selection'
import { evaluatorFor } from '../formula/engine'
import { findAll, FindMatch, FindQuery, replaceInRaw } from '../formula/find'
import { normalizedCell } from '../formula/input'
import type { EditorView } from '../view/editorview'
import { findBarKey } from '../view/types'
import { useSheetState } from './bridge'

export function FindBar({ view }: { view: EditorView }) {
  const state = useSheetState(view)
  const open = state.getField(findBarKey) as boolean
  const [text, setText] = useState('')
  const [replacement, setReplacement] = useState('')
  const [caseSensitive, setCaseSensitive] = useState(false)
  const [wholeCell, setWholeCell] = useState(false)
  const [wb, setWb] = useState(false)
  const [showReplace, setShowReplace] = useState(false)
  const [idx, setIdx] = useState(0)

  const query: FindQuery = { text, caseSensitive, wholeCell, workbook: wb }
  const matches: FindMatch[] = useMemo(
    () => (open ? findAll(state.doc, evaluatorFor(state.doc), query) : []),
    // state.doc 变化（替换/编辑）后自动重算
    // eslint-disable-next-line react-hooks/exhaustive-deps
    [open, text, caseSensitive, wholeCell, wb, state.doc],
  )

  if (!open) return null

  const close = (): void => {
    view.dispatch(view.state.tr.setMeta(findBarKey, false).setMeta('addToHistory', false))
    view.focus()
  }

  const gotoMatch = (m: FindMatch): void => {
    const tr = view.state.tr
    if (m.sheet !== view.state.doc.active) {
      tr.setActiveSheet(m.sheet).setMeta('addToHistory', false)
    }
    tr.setSelection(singleCell(m.row, m.col)).scrollIntoView()
    view.dispatch(tr)
  }

  const next = (delta: number): void => {
    if (matches.length === 0) return
    const i = ((idx + delta) % matches.length + matches.length) % matches.length
    setIdx(i)
    gotoMatch(matches[i])
  }

  const replaceCurrent = (): void => {
    if (matches.length === 0) return
    const m = matches[Math.min(idx, matches.length - 1)]
    const cell = view.state.doc.sheet(m.sheet).getCell(m.row, m.col)
    if (!cell) return
    const newRaw = replaceInRaw(cell.raw, query, replacement)
    if (newRaw === null) return // 显示值命中但 raw 不可替换（如格式化数字）
    const nextCell = normalizedCell(newRaw, cell)
    view.dispatch(view.state.tr.setCells(m.sheet, [{ row: m.row, col: m.col, cell: nextCell }]))
    // matches 随 state.doc 重算；idx 保持即指向原下一条
  }

  const replaceAll = (): void => {
    const bySheet = new Map<string, { row: number; col: number; cell: ReturnType<typeof normalizedCell> }[]>()
    let count = 0
    for (const m of matches) {
      const cell = view.state.doc.sheet(m.sheet).getCell(m.row, m.col)
      if (!cell) continue
      const newRaw = replaceInRaw(cell.raw, query, replacement)
      if (newRaw === null) continue
      const list = bySheet.get(m.sheet) ?? []
      list.push({ row: m.row, col: m.col, cell: normalizedCell(newRaw, cell) })
      bySheet.set(m.sheet, list)
      count++
    }
    if (count === 0) return
    const tr = view.state.tr
    for (const [sid, entries] of bySheet) tr.setCells(sid, entries)
    view.dispatch(tr)
    view.focus()
  }

  return (
    <div className="find-bar">
      <input
        className="find-input"
        placeholder="查找"
        value={text}
        autoFocus
        onChange={(e) => {
          setText(e.target.value)
          setIdx(0)
        }}
        onKeyDown={(e) => {
          if (e.key === 'Enter') next(e.shiftKey ? -1 : 1)
          if (e.key === 'Escape') close()
        }}
      />
      <span className="find-count">{text === '' ? '' : `${matches.length === 0 ? 0 : Math.min(idx + 1, matches.length)}/${matches.length}`}</span>
      <button onClick={() => next(-1)}>↑</button>
      <button onClick={() => next(1)}>↓</button>
      <label><input type="checkbox" checked={caseSensitive} onChange={(e) => setCaseSensitive(e.target.checked)} />Aa</label>
      <label><input type="checkbox" checked={wholeCell} onChange={(e) => setWholeCell(e.target.checked)} />整格</label>
      <label><input type="checkbox" checked={wb} onChange={(e) => setWb(e.target.checked)} />全簿</label>
      <button onClick={() => setShowReplace(!showReplace)}>替换</button>
      <button onClick={close}>×</button>
      {showReplace && (
        <>
          <input
            className="find-input"
            placeholder="替换为"
            value={replacement}
            onChange={(e) => setReplacement(e.target.value)}
            onKeyDown={(e) => {
              if (e.key === 'Escape') close()
            }}
          />
          <button onClick={replaceCurrent}>替换当前</button>
          <button onClick={replaceAll}>全部替换</button>
        </>
      )}
    </div>
  )
}
