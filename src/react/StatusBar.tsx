// 状态栏：左侧 就绪/编辑中；右侧对选区数字统计（忽略非数字，单格只显示计数）+ 缩放档位控件。
import { useEffect, useReducer, useState } from 'react'
import { forEachSelectionRange } from '../core/selection'
import type { SheetState } from '../core/state'
import { formatNumber } from '../formula/eval'
import { evaluatorFor } from '../formula/engine'
import { isEditing } from '../view/editbox'
import type { EditorView } from '../view/editorview'
import { zoomKey } from '../view/types'
import { ZOOM_LEVELS, zoomOf } from '../view/zoom'
import { useSheetState } from './bridge'

interface Props {
  view: EditorView
}

function collectStats(state: SheetState): { sum: number; count: number; single: boolean } {
  const ev = evaluatorFor(state.doc)
  const sheet = state.doc.active
  let sum = 0
  let count = 0
  let cellCount = 0
  forEachSelectionRange(state.selection, r => {
    for (let row = r.sr; row <= r.er; row++) {
      for (let col = r.sc; col <= r.ec; col++) {
        cellCount++
        const v = ev.get(sheet, row, col)
        if (typeof v === 'number') {
          sum += v
          count++
        }
      }
    }
  })
  return { sum, count, single: cellCount <= 1 }
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

  // 缩放档位：点击百分比弹出档位菜单；zoom 非文档态（addToHistory:false）
  const zoom = zoomOf(state, state.doc.active)
  const [showZoom, setShowZoom] = useState(false)

  return (
    <div className="status-bar">
      <span>{isEditing() ? '编辑中' : '就绪'}</span>
      <span className="status-right">
        <span>{stats}</span>
        <span className="status-zoom-wrap">
          <button className="status-zoom" onClick={() => setShowZoom(!showZoom)}>
            {Math.round(zoom * 100)}%
          </button>
          {showZoom && (
            <div className="zoom-menu">
              {ZOOM_LEVELS.map((z) => (
                <button
                  key={z}
                  className={'zoom-item' + (z === zoom ? ' active' : '')}
                  onClick={() => {
                    const field = (view.state.getField(zoomKey) as Record<string, number> | null) ?? {}
                    view.dispatch(
                      view.state.tr
                        .setMeta(zoomKey, { ...field, [view.state.doc.active]: z })
                        .setMeta('addToHistory', false),
                    )
                    setShowZoom(false)
                  }}
                >
                  {z * 100}%
                </button>
              ))}
            </div>
          )}
        </span>
      </span>
    </div>
  )
}
