// 状态栏：左侧 就绪/编辑中；右侧对选区数字统计（忽略非数字，单格只显示计数）+ 缩放档位控件。
import { useEffect, useReducer, useSyncExternalStore } from 'react'
import { getNotice, subscribeNotice } from '../app/notice'
import { workbookStorage } from '../app/storage'
import { forEachSelectionRange } from '@gmi/excel-core'
import type { SheetState } from '@gmi/excel-core'
import { formatNumber } from '@gmi/excel-core'
import { evaluatorFor } from '@gmi/excel-core'
import { isEditing } from '../view/editbox'
import type { EditorView } from '../view/editorview'
import { zoomKey } from '../view/types'
import { ZOOM_LEVELS, zoomOf } from '../view/zoom'
import { useSheetState } from './bridge'
import { Button } from './ui/Button'
import { Dropdown } from './ui/Dropdown'
import type { MenuEntry } from './ui/Menu'

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
  const saveStatus = useSyncExternalStore(workbookStorage.subscribeStatus, workbookStorage.getStatus)
  const notice = useSyncExternalStore(subscribeNotice, getNotice)
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

  // 缩放档位：点击百分比向上弹出档位菜单；zoom 非文档态（addToHistory:false）
  const zoom = zoomOf(state, state.doc.active)
  const zoomEntries: MenuEntry[] = ZOOM_LEVELS.map((z) => ({
    id: String(z),
    label: `${z * 100}%`,
    active: z === zoom,
    onSelect: () => {
      const field = (view.state.getField(zoomKey) as Record<string, number> | null) ?? {}
      view.dispatch(
        view.state.tr
          .setMeta(zoomKey, { ...field, [view.state.doc.active]: z })
          .setMeta('addToHistory', false),
      )
    },
  }))

  return (
    <div className="flex h-6 flex-none items-center justify-between border-t border-line bg-surface-2 px-3 text-xs text-ink-2">
      <span className="flex items-center">
        <span>{isEditing() ? '编辑中' : '就绪'}</span>
        {saveStatus.error ? <span className="ml-3 text-danger-deep">自动保存失败</span> : null}
        {notice ? <span className="ml-3 text-primary">{notice}</span> : null}
      </span>
      <span className="flex items-center gap-3">
        <span>{stats}</span>
        <Dropdown
          side="up"
          align="right"
          trigger={(_open, toggle) => (
            <Button variant="ghost" size="sm" onClick={toggle}>
              {Math.round(zoom * 100)}%
            </Button>
          )}
          entries={zoomEntries}
        />
      </span>
    </div>
  )
}
