// 自定义排序对话框：最多 3 个关键字（列 + 升/降），含表头勾选。
// 纯受控 UI 态；确定时计算排序条目并 dispatch 一个事务。
import { useState } from 'react'
import { showNotice } from '../app/notice'
import { CellRange, colName } from '@gmi/excel-core'
import { evaluatorFor } from '@gmi/excel-core'
import { computeSortEntries, sortBlockedByMerges } from '@gmi/excel-core'
import type { EditorView } from '@gmi/excel-view'
import { Button } from './ui/Button'
import { Dialog } from './ui/Dialog'
import { Select } from './ui/Select'

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
      showNotice('排序区域包含合并单元格，无法排序')
      return
    }
    const entries = computeSortEntries(sheet, view.state.doc.active, evaluatorFor(view.state.doc), range, active, hasHeader)
    view.dispatch(view.state.tr.setCells(view.state.doc.active, entries))
    view.focus()
    onClose()
  }

  return (
    <Dialog
      title="自定义排序"
      onClose={onClose}
      footer={
        <>
          <Button variant="outline" onClick={onClose}>
            取消
          </Button>
          <Button variant="primary" onClick={apply}>
            确定
          </Button>
        </>
      }
    >
      <div className="flex flex-col gap-2">
        {keys.map((k, i) => (
          <div className="flex items-center gap-2" key={i}>
            <span className="text-sm">{i === 0 ? '主要关键字' : `次要关键字 ${i}`}</span>
            <Select
              value={k.col === null ? '' : k.col}
              options={[
                ...(i > 0 ? [{ value: '', label: '（无）' }] : []),
                ...cols.map((c) => ({ value: c, label: `${colName(c)} 列` })),
              ]}
              onChange={(v) => {
                const next = [...keys]
                next[i] = { ...k, col: v === '' ? null : Number(v) }
                setKeys(next)
              }}
            />
            <Select
              value={k.asc ? 'asc' : 'desc'}
              options={[
                { value: 'asc', label: '升序' },
                { value: 'desc', label: '降序' },
              ]}
              onChange={(v) => {
                const next = [...keys]
                next[i] = { ...k, asc: v === 'asc' }
                setKeys(next)
              }}
            />
          </div>
        ))}
        <label className="flex items-center gap-2 text-sm">
          <input type="checkbox" checked={hasHeader} onChange={(e) => setHasHeader(e.target.checked)} />
          数据包含表头
        </label>
      </div>
    </Dialog>
  )
}
