// 查找/替换栏：Ctrl+F 或工具栏开启（findBarKey state field）。
// 匹配显示文本与公式原文；逐个定位（选中 + ensureVisible，跨表切 active 不入 undo）；
// 替换仅改 raw（经 normalizedCell 归一）；全部替换跨表合并为一个事务。
import { useMemo, useState } from 'react'
import { ChevronDown, ChevronUp, X } from 'lucide-react'
import { singleCell } from '../core/selection'
import { evaluatorFor } from '../formula/engine'
import { findAll, FindMatch, FindQuery, indexAfterReplace, replaceInRaw } from '../formula/find'
import { normalizedCell } from '../formula/input'
import type { EditorView } from '../view/editorview'
import { findBarKey } from '../view/types'
import { useSheetState } from './bridge'
import { Button } from './ui/Button'
import { IconButton } from './ui/IconButton'
import { TextInput } from './ui/TextInput'

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
    const i = Math.min(idx, matches.length - 1)
    const m = matches[i]
    const cell = view.state.doc.sheet(m.sheet).getCell(m.row, m.col)
    if (!cell) return
    const newRaw = replaceInRaw(cell.raw, query, replacement)
    if (newRaw === null) return // 显示值命中但 raw 不可替换（如格式化数字）
    const nextCell = normalizedCell(newRaw, cell)
    view.dispatch(view.state.tr.setCells(m.sheet, [{ row: m.row, col: m.col, cell: nextCell }]))
    // 替换后前进到下一匹配（替换文本仍命中时不反复替换同格），选中跟随；
    // dispatch 同步生效，此处直接按新 doc 重算，与 useMemo 重算结果一致
    const newMatches = findAll(view.state.doc, evaluatorFor(view.state.doc), query)
    const ni = indexAfterReplace(matches, i, newMatches)
    setIdx(ni)
    if (newMatches.length > 0) gotoMatch(newMatches[ni])
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
    <div className="flex items-center gap-1.5 border-b border-line-strong bg-surface-2 px-2 py-1 text-sm">
      <TextInput
        placeholder="查找"
        width={180}
        value={text}
        autoFocus
        onChange={(v) => {
          setText(v)
          setIdx(0)
        }}
        onKeyDown={(e) => {
          // 栏内 Ctrl/Cmd+F 不唤起浏览器默认查找（焦点留在栏内）
          if ((e.ctrlKey || e.metaKey) && e.key.toLowerCase() === 'f') e.preventDefault()
          if (e.key === 'Enter') next(e.shiftKey ? -1 : 1)
          if (e.key === 'Escape') close()
        }}
      />
      <span className="min-w-10 text-xs text-ink-2">
        {text === '' ? '' : `${matches.length === 0 ? 0 : Math.min(idx + 1, matches.length)}/${matches.length}`}
      </span>
      <IconButton icon={ChevronUp} tip="上一个" kbd="Shift+Enter" onClick={() => next(-1)} />
      <IconButton icon={ChevronDown} tip="下一个" kbd="Enter" onClick={() => next(1)} />
      <Button variant="ghost" size="sm" active={caseSensitive} onClick={() => setCaseSensitive(!caseSensitive)}>
        区分大小写
      </Button>
      <Button variant="ghost" size="sm" active={wholeCell} onClick={() => setWholeCell(!wholeCell)}>
        整个单元格
      </Button>
      <Button variant="ghost" size="sm" active={wb} onClick={() => setWb(!wb)}>
        全工作簿
      </Button>
      <Button variant="ghost" size="sm" active={showReplace} onClick={() => setShowReplace(!showReplace)}>
        替换
      </Button>
      <IconButton icon={X} tip="关闭" kbd="Esc" onClick={close} />
      {showReplace && (
        <>
          <TextInput
            placeholder="替换为"
            width={180}
            value={replacement}
            onChange={setReplacement}
            onKeyDown={(e) => {
              if ((e.ctrlKey || e.metaKey) && e.key.toLowerCase() === 'f') e.preventDefault()
              if (e.key === 'Escape') close()
            }}
          />
          <Button variant="ghost" size="sm" onClick={replaceCurrent}>
            替换当前
          </Button>
          <Button variant="ghost" size="sm" onClick={replaceAll}>
            全部替换
          </Button>
        </>
      )}
    </div>
  )
}
