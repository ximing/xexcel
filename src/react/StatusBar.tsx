// 状态栏：左侧 就绪/编辑中；右侧对选区数字统计（忽略非数字，单格只显示计数）。
import { useEffect, useReducer } from 'react'
import { selectionRange } from '../core/selection'
import type { SheetState } from '../core/state'
import { formatNumber } from '../formula/eval'
import { evaluatorFor } from '../formula/engine'
import { isEditing } from '../view/editbox'
import type { EditorView } from '../view/editorview'
import { useSheetState } from './bridge'

interface Props {
  view: EditorView
}

function collectStats(state: SheetState): { sum: number; count: number; single: boolean } {
  const r = selectionRange(state.selection)
  const ev = evaluatorFor(state.doc)
  const sheet = state.doc.active
  let sum = 0
  let count = 0
  for (let row = r.sr; row <= r.er; row++) {
    for (let col = r.sc; col <= r.ec; col++) {
      const v = ev.get(sheet, row, col)
      if (typeof v === 'number') {
        sum += v
        count++
      }
    }
  }
  return { sum, count, single: r.sr === r.er && r.sc === r.ec }
}

export function StatusBar({ view }: Props) {
  const state = useSheetState(view)
  // 编辑器开/关不产生事务，靠 MutationObserver 监听 view.dom 子节点（.xcell-editor 挂载/移除）刷新
  const [, forceUpdate] = useReducer((x: number) => x + 1, 0)
  useEffect(() => {
    const ob = new MutationObserver(() => forceUpdate())
    ob.observe(view.dom, { childList: true })
    return () => ob.disconnect()
  }, [view])

  const { sum, count, single } = collectStats(state)
  const stats = single
    ? `计数: ${count}`
    : `求和: ${formatNumber(sum)}　平均: ${count ? formatNumber(sum / count) : '-'}　计数: ${count}`

  return (
    <div className="status-bar">
      <span>{isEditing() ? '编辑中' : '就绪'}</span>
      <span>{stats}</span>
    </div>
  )
}
