// 工具栏：撤销/重做 + 粗体/斜体/文字色/背景色/对齐。操作回走 view.dispatch(applyStylePatch)。
import { applyStylePatch, mergeSelection, unmergeSelection } from '../core/commands'
import { selectionRange } from '../core/selection'
import { redo, redoDepth, undo, undoDepth } from '../core/history'
import type { CellStyle } from '../core/model'
import type { SheetState } from '../core/state'
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
      <span className="tool-sep" />
      <button
        className="tool-btn"
        title="合并单元格"
        onClick={() => {
          const r = selectionRange(view.state.selection)
          if (r.sr === r.er && r.sc === r.ec) return
          let nonEmpty = 0
          view.state.activeSheet.forEachInRange(r, (cell) => {
            if (cell && cell.raw !== '') nonEmpty++
          })
          if (nonEmpty > 1 && !window.confirm('合并仅保留左上角的值，其余内容将被清除。继续？')) return
          mergeSelection(view.state, (tr) => view.dispatch(tr))
          view.focus()
        }}
      >
        合
      </button>
      <button
        className="tool-btn"
        title="拆分单元格"
        onClick={() => {
          unmergeSelection(view.state, (tr) => view.dispatch(tr))
          view.focus()
        }}
      >
        拆
      </button>
      <span className="tool-sep" />
      <button
        className="tool-btn"
        title="上方插入行（选中整行时按行数）"
        disabled={!isFullRowSel(state)}
        onClick={() => {
          const r = selectionRange(view.state.selection)
          view.dispatch(view.state.tr.structure('row', r.sr, r.er - r.sr + 1, 'insert'))
          view.focus()
        }}
      >
        +行
      </button>
      <button
        className="tool-btn"
        title="删除选中行"
        disabled={!isFullRowSel(state)}
        onClick={() => {
          const r = selectionRange(view.state.selection)
          view.dispatch(view.state.tr.structure('row', r.sr, r.er - r.sr + 1, 'delete'))
          view.focus()
        }}
      >
        -行
      </button>
      <button
        className="tool-btn"
        title="左侧插入列"
        disabled={!isFullColSel(state)}
        onClick={() => {
          const r = selectionRange(view.state.selection)
          view.dispatch(view.state.tr.structure('col', r.sc, r.ec - r.sc + 1, 'insert'))
          view.focus()
        }}
      >
        +列
      </button>
      <button
        className="tool-btn"
        title="删除选中列"
        disabled={!isFullColSel(state)}
        onClick={() => {
          const r = selectionRange(view.state.selection)
          view.dispatch(view.state.tr.structure('col', r.sc, r.ec - r.sc + 1, 'delete'))
          view.focus()
        }}
      >
        -列
      </button>
      <span className="tool-sep" />
      <button
        className={'tool-btn' + (state.activeSheet.frozenRows === 1 && state.activeSheet.frozenCols === 0 ? ' active' : '')}
        title="冻结首行"
        onClick={() => {
          view.dispatch(view.state.tr.setFrozen(state.activeSheet.frozenRows === 1 && state.activeSheet.frozenCols === 0 ? 0 : 1, state.activeSheet.frozenCols))
          view.focus()
        }}
      >
        冻行
      </button>
      <button
        className={'tool-btn' + (state.activeSheet.frozenCols === 1 && state.activeSheet.frozenRows === 0 ? ' active' : '')}
        title="冻结首列"
        onClick={() => {
          view.dispatch(view.state.tr.setFrozen(state.activeSheet.frozenRows, state.activeSheet.frozenCols === 1 && state.activeSheet.frozenRows === 0 ? 0 : 1))
          view.focus()
        }}
      >
        冻列
      </button>
      <button
        className="tool-btn"
        title="冻结到当前选区（其上方与左侧）"
        onClick={() => {
          const { row, col } = view.state.selection.focus
          view.dispatch(view.state.tr.setFrozen(row, col))
          view.focus()
        }}
      >
        冻至
      </button>
      <button
        className="tool-btn"
        title="取消冻结"
        disabled={state.activeSheet.frozenRows === 0 && state.activeSheet.frozenCols === 0}
        onClick={() => {
          view.dispatch(view.state.tr.setFrozen(0, 0))
          view.focus()
        }}
      >
        解冻
      </button>
      <span className="tool-sep" />
      <button
        className="tool-btn"
        title="隐藏选中行"
        disabled={!isFullRowSel(state)}
        onClick={() => {
          const r = selectionRange(view.state.selection)
          const indices: number[] = []
          for (let row = r.sr; row <= r.er; row++) indices.push(row)
          view.dispatch(view.state.tr.setHidden('row', indices, true))
          view.focus()
        }}
      >
        隐行
      </button>
      <button
        className="tool-btn"
        title="隐藏选中列"
        disabled={!isFullColSel(state)}
        onClick={() => {
          const r = selectionRange(view.state.selection)
          const indices: number[] = []
          for (let col = r.sc; col <= r.ec; col++) indices.push(col)
          view.dispatch(view.state.tr.setHidden('col', indices, true))
          view.focus()
        }}
      >
        隐列
      </button>
      <button
        className="tool-btn"
        title="取消隐藏（选区范围内的隐藏行列）"
        disabled={!hasHiddenInSel(state)}
        onClick={() => {
          const r = selectionRange(view.state.selection)
          const sheet = view.state.activeSheet
          const rows = sheet.hiddenRows.filter((i) => i >= r.sr && i <= r.er)
          const cols = sheet.hiddenCols.filter((i) => i >= r.sc && i <= r.ec)
          const tr = view.state.tr
          if (rows.length) tr.setHidden('row', rows, false)
          if (cols.length) tr.setHidden('col', cols, false)
          view.dispatch(tr)
          view.focus()
        }}
      >
        取消隐
      </button>
      <span className="tool-sep" />
      <button
        className="tool-btn"
        title="重置选中行/列尺寸为默认"
        onClick={() => {
          const r = selectionRange(view.state.selection)
          const sheet = view.state.activeSheet
          const tr = view.state.tr
          if (r.sc === 0 && r.ec === sheet.colCount - 1) {
            for (let row = r.sr; row <= r.er; row++) tr.resize('row', row, null)
          } else if (r.sr === 0 && r.er === sheet.rowCount - 1) {
            for (let col = r.sc; col <= r.ec; col++) tr.resize('col', col, null)
          } else {
            return
          }
          view.dispatch(tr)
          view.focus()
        }}
      >
        重置
      </button>
    </div>
  )
}

function isFullRowSel(state: SheetState): boolean {
  const r = selectionRange(state.selection)
  return r.sc === 0 && r.ec === state.activeSheet.colCount - 1 && !(r.sr === 0 && r.er === state.activeSheet.rowCount - 1 && r.sc === 0 && r.ec === state.activeSheet.colCount - 1)
}

function isFullColSel(state: SheetState): boolean {
  const r = selectionRange(state.selection)
  return r.sr === 0 && r.er === state.activeSheet.rowCount - 1 && !(r.sr === 0 && r.er === state.activeSheet.rowCount - 1 && r.sc === 0 && r.ec === state.activeSheet.colCount - 1)
}

function hasHiddenInSel(state: SheetState): boolean {
  const r = selectionRange(state.selection)
  const sheet = state.activeSheet
  return (
    sheet.hiddenRows.some((i) => i >= r.sr && i <= r.er) ||
    sheet.hiddenCols.some((i) => i >= r.sc && i <= r.ec)
  )
}
