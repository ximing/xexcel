// 公式栏：名称框（活动格 A1）+ raw 输入框。聚焦时走本地编辑态，
// Enter 提交 setCell 并把焦点还给表格，Esc 还原；失焦提交（Excel 习惯）。
import { useEffect, useRef, useState } from 'react'
import { toA1 } from '../core/addr'
import { normalizedCell } from '../formula/input'
import type { EditorView } from '../view/editorview'
import { useSheetState } from './bridge'

interface Props {
  view: EditorView
}

export function FormulaBar({ view }: Props) {
  const state = useSheetState(view)
  const { row, col } = state.selection.activeCell
  const raw = state.activeSheet.getCell(row, col)?.raw ?? ''

  const [text, setText] = useState(raw)
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

  const commit = (): void => {
    if (textRef.current !== rawRef.current) {
      const oldCell = state.activeSheet.getCell(row, col)
      const next = normalizedCell(textRef.current, oldCell)
      view.dispatch(view.state.tr.setCell(row, col, next.raw, next.style))
      rawRef.current = textRef.current
    }
  }

  return (
    <div className="formula-bar">
      <div className="name-box">{toA1(row, col)}</div>
      <input
        className="formula-input"
        value={text}
        onFocus={() => {
          focusedRef.current = true
        }}
        onChange={(e) => {
          textRef.current = e.target.value
          setText(e.target.value)
        }}
        onBlur={() => {
          focusedRef.current = false
          commit()
        }}
        onKeyDown={(e) => {
          if (e.nativeEvent.isComposing) return
          if (e.key === 'Enter') {
            e.preventDefault()
            commit()
            view.focus() // 焦点回表格（input 随之 blur，commit 因 rawRef 已同步而幂等）
          } else if (e.key === 'Escape') {
            e.preventDefault()
            textRef.current = rawRef.current
            setText(rawRef.current)
            view.focus()
          }
        }}
      />
    </div>
  )
}
