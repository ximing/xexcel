// Sheet 标签栏：增/删/改名/切换。纯投影组件：读 useSheetState 快照，
// 写操作一律 view.dispatch(tr)。切换表不入 undo 栈并重置选区到 A1。
import { useState } from 'react'
import { nextSheetId, nextSheetName, SheetData } from '../core/model'
import { singleCell } from '../core/selection'
import type { EditorView } from '../view/editorview'
import { useSheetState } from './bridge'

interface Props {
  view: EditorView
}

// 空表判定：usedRange 为单格且该格无内容（usedRange 对空表返回全 0）
function isSheetEmpty(data: SheetData): boolean {
  const r = data.usedRange()
  return r.sr === 0 && r.sc === 0 && r.er === 0 && r.ec === 0 && !data.getCell(0, 0)
}

export function SheetTabBar({ view }: Props) {
  const state = useSheetState(view)
  const [editingId, setEditingId] = useState<string | null>(null)
  const [draft, setDraft] = useState('')

  const switchTo = (id: string): void => {
    if (id === state.doc.active) return
    view.dispatch(
      state.tr.setActiveSheet(id).setSelection(singleCell(0, 0)).setMeta('addToHistory', false),
    )
    view.focus()
  }

  const addSheet = (): void => {
    const id = nextSheetId(state.doc)
    const name = nextSheetName(state.doc)
    const config = { rowCount: state.activeSheet.rowCount, colCount: state.activeSheet.colCount }
    view.dispatch(state.tr.insertSheet(id, name, config).setSelection(singleCell(0, 0)))
    view.focus()
  }

  const removeSheet = (id: string, name: string): void => {
    if (state.doc.order.length <= 1) return
    if (!isSheetEmpty(state.doc.sheet(id))) {
      if (!window.confirm(`确定删除工作表「${name}」？可通过撤销恢复。`)) return
    }
    const tr = state.tr.removeSheet(id)
    if (id === state.doc.active) tr.setSelection(singleCell(0, 0))
    view.dispatch(tr)
    view.focus()
  }

  const startRename = (id: string): void => {
    setEditingId(id)
    setDraft(state.doc.names.get(id) ?? '')
  }

  const commitRename = (id: string): void => {
    setEditingId(null)
    const name = draft.trim()
    if (name === '' || name === state.doc.names.get(id)) return
    const dup = [...state.doc.names.entries()].some(
      ([other, n]) => other !== id && n.toLowerCase() === name.toLowerCase(),
    )
    if (dup) {
      window.alert(`工作表名称重复：${name}`)
      return
    }
    view.dispatch(state.tr.renameSheet(id, name))
    view.focus()
  }

  return (
    <div className="sheet-tab-bar">
      <button className="sheet-tab-add" title="新增工作表" onClick={addSheet}>
        +
      </button>
      {state.doc.order.map((id) => {
        const name = state.doc.names.get(id) ?? id
        const active = id === state.doc.active
        if (editingId === id) {
          return (
            <input
              key={id}
              className="sheet-tab-input"
              value={draft}
              autoFocus
              onChange={(e) => setDraft(e.target.value)}
              onBlur={() => commitRename(id)}
              onKeyDown={(e) => {
                if (e.nativeEvent.isComposing) return
                if (e.key === 'Enter') commitRename(id)
                else if (e.key === 'Escape') setEditingId(null)
              }}
            />
          )
        }
        return (
          <div
            key={id}
            className={active ? 'sheet-tab active' : 'sheet-tab'}
            onClick={() => switchTo(id)}
            onDoubleClick={() => startRename(id)}
          >
            <span>{name}</span>
            {state.doc.order.length > 1 ? (
              <button
                className="sheet-tab-close"
                title="删除工作表"
                onClick={(e) => {
                  e.stopPropagation()
                  removeSheet(id, name)
                }}
              >
                ×
              </button>
            ) : null}
          </div>
        )
      })}
    </div>
  )
}
