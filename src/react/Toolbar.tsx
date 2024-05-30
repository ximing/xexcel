// 工具栏：撤销/重做 + 粗体/斜体/文字色/背景色/对齐。操作回走 view.dispatch(applyStylePatch)。
import { useState } from 'react'
import { applyStylePatch, mergeSelection, unmergeSelection } from '../core/commands'
import { computeBorderStyles, BorderPreset } from '../core/border'
import { selectionRange } from '../core/selection'
import { redo, redoDepth, undo, undoDepth } from '../core/history'
import type { BorderLineStyle, CellStyle } from '../core/model'
import type { SheetState } from '../core/state'
import type { EditorView } from '../view/editorview'
import { findBarKey, formatPainterKey, FormatPainterState } from '../view/types'
import { useSheetState } from './bridge'
import { adjustDecimals } from '../formula/format'
import { evaluatorFor } from '../formula/engine'
import { computeSortEntries, sortBlockedByMerges } from '../formula/sort'
import { SortDialog } from './SortDialog'
import { CondFormatDialog } from './CondFormatDialog'

interface Props {
  view: EditorView
}

export function Toolbar({ view }: Props) {
  const state = useSheetState(view)
  const { row, col } = state.selection.activeCell
  const active: CellStyle = state.activeSheet.getCell(row, col)?.style ?? {}
  const fp = state.getField(formatPainterKey) as FormatPainterState | null | undefined
  const [showSort, setShowSort] = useState(false)
  const [showCF, setShowCF] = useState(false)
  const [showBorder, setShowBorder] = useState(false)
  const [borderLine, setBorderLine] = useState<BorderLineStyle>('thin')
  const [borderColor, setBorderColor] = useState('#000000')

  const applyBorder = (preset: BorderPreset): void => {
    const sheet = view.state.activeSheet
    const edge = preset === 'none' ? null : { style: borderLine, color: borderColor }
    const entries: ReturnType<typeof computeBorderStyles>[] = []
    for (const r of view.state.selection.ranges) entries.push(computeBorderStyles(sheet, r, preset, edge))
    view.dispatch(view.state.tr.setCellStyles(entries.flat()))
    setShowBorder(false)
    view.focus()
  }

  const quickSort = (asc: boolean): void => {
    const r = selectionRange(view.state.selection)
    if (r.sr === r.er) return
    const sheet = view.state.activeSheet
    if (sortBlockedByMerges(sheet, r)) {
      window.alert('排序区域包含合并单元格，无法排序')
      return
    }
    const entries = computeSortEntries(sheet, view.state.doc.active, evaluatorFor(view.state.doc), r, [{ col: r.sc, asc }], false)
    view.dispatch(view.state.tr.setCells(view.state.doc.active, entries))
    view.focus()
  }

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
        className={'tool-btn' + (fp ? ' active' : '')}
        title="格式刷（单击取格式刷一次，双击锁定连刷，Esc 解除）"
        onClick={() => {
          if (fp) {
            view.dispatch(view.state.tr.setMeta(formatPainterKey, null).setMeta('addToHistory', false))
          } else {
            const src = { ...(view.state.activeSheet.getCell(row, col)?.style ?? {}) }
            view.dispatch(view.state.tr.setMeta(formatPainterKey, { style: src, locked: false }).setMeta('addToHistory', false))
          }
          view.focus()
        }}
        onDoubleClick={() => {
          const src = { ...(view.state.activeSheet.getCell(row, col)?.style ?? {}) }
          view.dispatch(view.state.tr.setMeta(formatPainterKey, { style: src, locked: true }).setMeta('addToHistory', false))
          view.focus()
        }}
      >
        刷
      </button>
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
      <button
        className={'tool-btn' + (active.wrap ? ' active' : '')}
        title="自动换行"
        onClick={() => {
          patch({ wrap: active.wrap ? undefined : true })
          view.focus()
        }}
      >
        换行
      </button>
      {(['top', 'middle', 'bottom'] as const).map((v) => (
        <button
          key={v}
          className={'tool-btn' + ((active.vAlign ?? 'bottom') === v ? ' active' : '')}
          title={{ top: '顶端对齐', middle: '垂直居中', bottom: '底端对齐' }[v]}
          onClick={() => {
            patch({ vAlign: v })
            view.focus()
          }}
        >
          {{ top: '上', middle: '中', bottom: '下' }[v]}
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
          const ranges = view.state.selection.ranges
          if (ranges.every(r => r.sr === r.er && r.sc === r.ec)) return
          let nonEmpty = 0
          const sheet = view.state.activeSheet
          for (const r of ranges) {
            sheet.forEachInRange(r, (cell) => {
              if (cell && cell.raw !== '') nonEmpty++
            })
          }
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
      <span className="tool-border-wrap">
        <button className="tool-btn" title="边框" onClick={() => setShowBorder(!showBorder)}>
          框
        </button>
        {showBorder && (
          <div className="tool-border-panel">
            <div className="tool-border-presets">
              {([
                ['none', '无框'], ['all', '全框'], ['outer', '外框'], ['inner', '内框'],
                ['top', '上'], ['bottom', '下'], ['left', '左'], ['right', '右'],
              ] as [BorderPreset, string][]).map(([p, label]) => (
                <button key={p} className="tool-btn" onClick={() => applyBorder(p)}>{label}</button>
              ))}
            </div>
            <select
              className="tool-select"
              title="线型"
              value={borderLine}
              onChange={(e) => setBorderLine(e.target.value as BorderLineStyle)}
            >
              <option value="thin">细线</option>
              <option value="medium">中线</option>
              <option value="thick">粗线</option>
              <option value="dashed">虚线</option>
              <option value="dotted">点线</option>
              <option value="double">双线</option>
              <option value="hair">极细</option>
              <option value="mediumDashed">中虚线</option>
            </select>
            <input
              className="tool-color"
              type="color"
              title="边框颜色"
              value={borderColor}
              onChange={(e) => setBorderColor(e.target.value)}
            />
          </div>
        )}
      </span>
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
          const { row, col } = view.state.selection.activeCell
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
      <button className="tool-btn" title="按选区首列升序" disabled={!canSort(state)} onClick={() => quickSort(true)}>
        A↓
      </button>
      <button className="tool-btn" title="按选区首列降序" disabled={!canSort(state)} onClick={() => quickSort(false)}>
        Z↓
      </button>
      <button className="tool-btn" title="自定义排序" disabled={!canSort(state)} onClick={() => setShowSort(true)}>
        排序
      </button>
      {showSort && <SortDialog view={view} range={selectionRange(view.state.selection)} onClose={() => setShowSort(false)} />}
      <button
        className={'tool-btn' + (state.activeSheet.filter ? ' active' : '')}
        title="自动筛选（对选区启用/清除全表筛选）"
        onClick={() => {
          const sheet = view.state.activeSheet
          if (sheet.filter) {
            view.dispatch(view.state.tr.setFilter(undefined))
          } else {
            const r = selectionRange(view.state.selection)
            const range = r.sr === r.er && r.sc === r.ec ? sheet.usedRange() : r
            if (range.er <= range.sr) {
              window.alert('筛选区域至少需要两行（表头 + 数据）')
              return
            }
            view.dispatch(view.state.tr.setFilter({ range, criteria: {} }))
          }
          view.focus()
        }}
      >
        筛
      </button>
      <button
        className="tool-btn"
        title="查找/替换（Ctrl+F）"
        onClick={() => {
          view.dispatch(view.state.tr.setMeta(findBarKey, true).setMeta('addToHistory', false))
        }}
      >
        查
      </button>
      <button className="tool-btn" title="条件格式" onClick={() => setShowCF(true)}>
        条件
      </button>
      {showCF && <CondFormatDialog view={view} onClose={() => setShowCF(false)} />}
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

function canSort(state: SheetState): boolean {
  const r = selectionRange(state.selection)
  return r.er > r.sr
}
