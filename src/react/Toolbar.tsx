// 工具栏：撤销/重做 + 粗体/斜体/文字色/背景色/对齐。操作回走 view.dispatch(applyStylePatch)。
import { applyStylePatch } from '../core/commands'
import { redo, redoDepth, undo, undoDepth } from '../core/history'
import type { CellStyle } from '../core/model'
import type { EditorView } from '../view/editorview'
import { useSheetState } from './bridge'
import { adjustDecimals } from '../formula/format'

interface Props {
  view: EditorView
}

export function Toolbar({ view }: Props) {
  const state = useSheetState(view)
  const { row, col } = state.selection.focus
  const active: CellStyle = state.activeSheet.getCell(row, col)?.style ?? {}

  const patch = (p: Partial<CellStyle>): void => {
    applyStylePatch(p)(view.state, (tr) => view.dispatch(tr))
  }

  return (
    <div className="toolbar">
      <button
        className="tool-btn"
        title="撤销"
        disabled={undoDepth(state) === 0}
        onClick={() => {
          undo(view.state, (tr) => view.dispatch(tr))
          view.focus()
        }}
      >
        ↩
      </button>
      <button
        className="tool-btn"
        title="重做"
        disabled={redoDepth(state) === 0}
        onClick={() => {
          redo(view.state, (tr) => view.dispatch(tr))
          view.focus()
        }}
      >
        ↪
      </button>
      <span className="tool-sep" />
      <button
        className={'tool-btn' + (active.bold ? ' active' : '')}
        title="加粗"
        style={{ fontWeight: 'bold' }}
        onClick={() => {
          patch(active.bold ? { bold: undefined } : { bold: true })
          view.focus()
        }}
      >
        B
      </button>
      <button
        className={'tool-btn' + (active.italic ? ' active' : '')}
        title="斜体"
        style={{ fontStyle: 'italic' }}
        onClick={() => {
          patch(active.italic ? { italic: undefined } : { italic: true })
          view.focus()
        }}
      >
        I
      </button>
      <span className="tool-sep" />
      <input
        className="tool-color"
        type="color"
        title="文字颜色"
        defaultValue="#202124"
        onChange={(e) => {
          patch({ color: e.target.value })
          view.focus()
        }}
      />
      <input
        className="tool-color"
        type="color"
        title="背景颜色"
        defaultValue="#ffffff"
        onChange={(e) => {
          patch({ bg: e.target.value })
          view.focus()
        }}
      />
      <span className="tool-sep" />
      {(['left', 'center', 'right'] as const).map((a) => (
        <button
          key={a}
          className={'tool-btn' + ((active.align ?? 'left') === a ? ' active' : '')}
          title={{ left: '左对齐', center: '居中', right: '右对齐' }[a]}
          onClick={() => {
            patch({ align: a })
            view.focus()
          }}
        >
          {{ left: '左', center: '中', right: '右' }[a]}
        </button>
      ))}
      <span className="tool-sep" />
      <select
        className="tool-select"
        title="字号"
        value={active.fontSize ?? 13}
        onChange={(e) => {
          patch({ fontSize: Number(e.target.value) })
          view.focus()
        }}
      >
        {[10, 11, 12, 13, 14, 16, 18, 24].map((s) => (
          <option key={s} value={s}>
            {s}
          </option>
        ))}
      </select>
      <select
        className="tool-select"
        title="字体"
        value={active.fontFamily ?? ''}
        onChange={(e) => {
          patch({ fontFamily: e.target.value || undefined })
          view.focus()
        }}
      >
        <option value="">默认</option>
        <option value="SimSun, serif">宋体</option>
        <option value="SimHei, sans-serif">黑体</option>
        <option value="monospace">等宽</option>
      </select>
      <button
        className={'tool-btn' + (active.underline ? ' active' : '')}
        title="下划线"
        style={{ textDecoration: 'underline' }}
        onClick={() => {
          patch(active.underline ? { underline: undefined } : { underline: true })
          view.focus()
        }}
      >
        U
      </button>
      <button
        className={'tool-btn' + (active.strikethrough ? ' active' : '')}
        title="删除线"
        style={{ textDecoration: 'line-through' }}
        onClick={() => {
          patch(active.strikethrough ? { strikethrough: undefined } : { strikethrough: true })
          view.focus()
        }}
      >
        S
      </button>
      <span className="tool-sep" />
      <button
        className="tool-btn"
        title="千分位"
        onClick={() => {
          patch({ numFmt: '#,##0.00' })
          view.focus()
        }}
      >
        ,
      </button>
      <button
        className={'tool-btn' + (active.numFmt === '0%' ? ' active' : '')}
        title="百分比"
        onClick={() => {
          patch({ numFmt: '0%' })
          view.focus()
        }}
      >
        %
      </button>
      <button
        className="tool-btn"
        title="货币"
        onClick={() => {
          patch({ numFmt: '¥#,##0.00' })
          view.focus()
        }}
      >
        ¥
      </button>
      <button
        className="tool-btn"
        title="增加小数位"
        onClick={() => {
          patch({ numFmt: adjustDecimals(active.numFmt, 1) })
          view.focus()
        }}
      >
        .0+
      </button>
      <button
        className="tool-btn"
        title="减少小数位"
        onClick={() => {
          patch({ numFmt: adjustDecimals(active.numFmt, -1) })
          view.focus()
        }}
      >
        .0-
      </button>
    </div>
  )
}
