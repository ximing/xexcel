// 自定义排序对话框：最多 3 个关键字（列 + 升/降），含表头勾选。
// 纯受控 UI 态；确定时计算排序条目并 dispatch 一个事务。
import { useState } from 'react'
import { CellRange, colName } from '../core/addr'
import { evaluatorFor } from '../formula/engine'
import { computeSortEntries, sortBlockedByMerges } from '../formula/sort'
import type { EditorView } from '../view/editorview'

interface Props {
  view: EditorView
  range: CellRange
  onClose: () => void
}

interface KeyDraft {
  col: number | null // null = 未启用
  asc: boolean
}

export function SortDialog({ view, range, onClose }: Props) {
  const [keys, setKeys] = useState<KeyDraft[]>([
    { col: range.sc, asc: true },
    { col: null, asc: true },
    { col: null, asc: true },
  ])
  const [hasHeader, setHasHeader] = useState(true)

  const cols: number[] = []
  for (let c = range.sc; c <= range.ec; c++) cols.push(c)

  const apply = (): void => {
    // 主关键字无「（无）」选项，active 恒 ≥1
    const active = keys.filter((k): k is { col: number; asc: boolean } => k.col !== null)
    const sheet = view.state.activeSheet
    if (sortBlockedByMerges(sheet, range)) {
      window.alert('排序区域包含合并单元格，无法排序')
      return
    }
    const entries = computeSortEntries(sheet, view.state.doc.active, evaluatorFor(view.state.doc), range, active, hasHeader)
    view.dispatch(view.state.tr.setCells(view.state.doc.active, entries))
    view.focus()
    onClose()
  }

  return (
    <div className="dialog-backdrop" onClick={onClose}>
      <div
        className="dialog"
        tabIndex={-1}
        autoFocus
        onClick={(e) => e.stopPropagation()}
        onKeyDown={(e) => {
          if (e.key === 'Escape') onClose()
        }}
      >
        <div className="dialog-title">自定义排序</div>
        {keys.map((k, i) => (
          <div className="dialog-row" key={i}>
            <span>{i === 0 ? '主要关键字' : `次要关键字 ${i}`}</span>
            <select
              value={k.col === null ? '' : k.col}
              onChange={(e) => {
                const next = [...keys]
                next[i] = { ...k, col: e.target.value === '' ? null : Number(e.target.value) }
                setKeys(next)
              }}
            >
              {i > 0 && <option value="">（无）</option>}
              {cols.map((c) => (
                <option key={c} value={c}>
                  {colName(c)} 列
                </option>
              ))}
            </select>
            <select
              value={k.asc ? 'asc' : 'desc'}
              onChange={(e) => {
                const next = [...keys]
                next[i] = { ...k, asc: e.target.value === 'asc' }
                setKeys(next)
              }}
            >
              <option value="asc">升序</option>
              <option value="desc">降序</option>
            </select>
          </div>
        ))}
        <label className="dialog-row">
          <input type="checkbox" checked={hasHeader} onChange={(e) => setHasHeader(e.target.checked)} />
          数据包含表头
        </label>
        <div className="dialog-actions">
          <button onClick={apply}>确定</button>
          <button onClick={onClose}>取消</button>
        </div>
      </div>
    </div>
  )
}
