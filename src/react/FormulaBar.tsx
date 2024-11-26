// 公式栏：名称框（活动格 A1）+ raw 输入框。聚焦时走本地编辑态，
// Enter 提交 setCell 并把焦点还给表格，Esc 还原；失焦提交（Excel 习惯）。
// F5：输入 =… 时画布高亮被引区域（refHighlightKey meta）+ 函数名补全下拉。
import { useEffect, useRef, useState } from 'react'
import { toA1 } from '../core/addr'
import { validateInput } from '../core/validation'
import { showNotice } from '../app/notice'
import { functionNames } from '../formula/eval'
import { normalizedCell } from '../formula/input'
import { completionCandidates } from '../formula/rangeRefs'
import type { EditorView } from '../view/editorview'
import { refHighlightKey } from '../view/types'
import { useSheetState } from './bridge'

const FONT_STACK =
  '-apple-system, BlinkMacSystemFont, "Segoe UI", Roboto, "Helvetica Neue", Arial, sans-serif'

interface Props {
  view: EditorView
}

export function FormulaBar({ view }: Props) {
  const state = useSheetState(view)
  const { row, col } = state.selection.activeCell
  const raw = state.activeSheet.getCell(row, col)?.raw ?? ''

  const [text, setText] = useState(raw)
  const [completion, setCompletion] = useState<string[]>([])
  const [selIndex, setSelIndex] = useState(-1)
  // refs 与 state 双轨：blur/Esc 时序里 React state 可能尚未刷新，提交判定一律读 ref
  const textRef = useRef(raw)
  const rawRef = useRef(raw)
  const focusedRef = useRef(false)

  // 外部 state 变化（选区移动/编辑提交/undo）且非自身编辑时同步输入框
  useEffect(() => {
    rawRef.current = raw
    if (!focusedRef.current) {
      textRef.current = raw
      setText(raw)
    }
  }, [raw, row, col])

  // 画布引用高亮 meta：以 = 开头 → 设为文本，否则清 null
  const setHl = (t: string): void => {
    view.dispatch(
      view.state.tr.setMeta(refHighlightKey, t.startsWith('=') ? t : null).setMeta('addToHistory', false),
    )
  }

  // 返回值：false = 验证拒绝（调用方不得 blur/移焦，否则 blur 二次 commit 双 notice）
  const commit = (): boolean => {
    if (textRef.current !== rawRef.current) {
      // 数据验证：拒绝则不提交（输入框内容保留，用户可继续修改）；未变更文本失焦不校验
      const reason = validateInput(state.activeSheet.validations, row, col, textRef.current)
      if (reason) {
        showNotice(reason)
        return false
      }
      const oldCell = state.activeSheet.getCell(row, col)
      const next = normalizedCell(textRef.current, oldCell)
      view.dispatch(view.state.tr.setCell(row, col, next.raw, next.style))
      rawRef.current = textRef.current
    }
    setCompletion([])
    setSelIndex(-1)
    setHl('') // 清画布引用高亮
    return true
  }

  // 接受补全：末尾标识符 token 替换为 NAME(
  const accept = (name: string): void => {
    const t = textRef.current
    const eq = t.lastIndexOf('=')
    if (eq < 0) return
    const m = /[A-Za-z]+$/.exec(t.slice(eq + 1))
    if (!m) return
    const next = t.slice(0, eq + 1 + m.index) + name + '('
    textRef.current = next
    setText(next)
    setHl(next)
    setCompletion([])
    setSelIndex(-1)
  }

  return (
    <div className="formula-bar">
      <div className="name-box">{toA1(row, col)}</div>
      <div style={{ position: 'relative', flex: 1, minWidth: 0, display: 'flex' }}>
        <input
          className="formula-input"
          value={text}
          onFocus={() => {
            focusedRef.current = true
          }}
          onChange={(e) => {
            const t = e.target.value
            textRef.current = t
            setText(t)
            setHl(t)
            const cands = completionCandidates(t, functionNames())
            setCompletion(cands)
            setSelIndex(cands.length ? 0 : -1)
          }}
          onBlur={() => {
            focusedRef.current = false
            commit()
          }}
          onKeyDown={(e) => {
            if (e.nativeEvent.isComposing) return
            if (completion.length > 0) {
              if (e.key === 'ArrowDown') {
                e.preventDefault()
                setSelIndex((i) => Math.min(i + 1, completion.length - 1))
                return
              }
              if (e.key === 'ArrowUp') {
                e.preventDefault()
                setSelIndex((i) => Math.max(i - 1, 0))
                return
              }
              if (e.key === 'Tab' || (e.key === 'Enter' && !e.shiftKey)) {
                e.preventDefault()
                accept(completion[selIndex < 0 ? 0 : selIndex])
                return
              }
              if (e.key === 'Escape') {
                e.preventDefault()
                setCompletion([])
                setSelIndex(-1)
                return
              }
            }
            if (e.key === 'Enter') {
              e.preventDefault()
              if (commit()) view.focus() // 焦点回表格（input 随之 blur，commit 因 rawRef 已同步而幂等）；拒绝则保持焦点
            } else if (e.key === 'Escape') {
              e.preventDefault()
              textRef.current = rawRef.current
              setText(rawRef.current)
              setCompletion([])
              setSelIndex(-1)
              setHl('')
              view.focus()
            }
          }}
        />
        {completion.length > 0 && (
          <div
            className="autocomplete"
            style={{
              position: 'absolute',
              top: '100%',
              left: 0,
              minWidth: '100%',
              background: '#ffffff',
              border: '1px solid #d9dce1',
              borderTop: 'none',
              zIndex: 11,
              maxHeight: '200px',
              overflowY: 'auto',
              font: `13px ${FONT_STACK}`,
              boxShadow: '0 2px 6px rgba(0,0,0,0.12)',
            }}
          >
            {completion.map((name, i) => (
              <div
                key={name}
                onMouseDown={(e) => {
                  e.preventDefault()
                  accept(name)
                }}
                style={{
                  padding: '2px 8px',
                  cursor: 'pointer',
                  background: i === selIndex ? '#e8f0fe' : '#ffffff',
                  color: i === selIndex ? '#1a73e8' : '#202124',
                }}
              >
                {name}
              </div>
            ))}
          </div>
        )}
      </div>
    </div>
  )
}
