// Sheet 标签栏：增/删/改名/切换/拖动排序/右键菜单。纯投影组件：读 useSheetState 快照，
// 写操作一律 view.dispatch(tr)。切换表不入 undo 栈并重置选区到 A1。
// 增/删/改名逻辑在 sheetOps.ts（与右键菜单共用）。
// 拖拽排序：mousedown 记起点 → window mousemove 超 5px 进入拖拽 → 按 clientX 与
// 各 tab 中点算插入位并渲染指示线 → mouseup 派发 moveSheet；Esc 取消；
// 拖拽结束后抑制紧随的 click（避免误触切换表）。
import { Fragment, useEffect, useRef, useState } from 'react'
import { singleCell } from '../core/selection'
import type { EditorView } from '../view/editorview'
import { contextMenuKey, tabRenameKey } from '../view/types'
import { useSheetState } from './bridge'
import { addSheet, removeSheet, renameSheet } from './sheetOps'

interface Props {
  view: EditorView
}

// 一次拖拽会话：active=false 表示已按下但未越阈值（仍视为点击）
interface DragSession {
  id: string
  from: number
  startX: number
  startY: number
  active: boolean
  ins: number // 视觉插入位（0..n，含被拖 tab 自身参与测距）
}

export function SheetTabBar({ view }: Props) {
  const state = useSheetState(view)
  const [editingId, setEditingId] = useState<string | null>(null)
  const [draft, setDraft] = useState('')
  const barRef = useRef<HTMLDivElement>(null)
  const sessionRef = useRef<DragSession | null>(null)
  const [drag, setDrag] = useState<{ id: string; ins: number } | null>(null)
  const suppressClick = useRef(false)

  // window 级监听常驻（会话存 ref，handler 不依赖渲染快照）
  useEffect(() => {
    const calcIns = (clientX: number): number => {
      const bar = barRef.current
      if (!bar) return 0
      const tabs = Array.from(bar.querySelectorAll<HTMLElement>('[data-sheet-id]'))
      let ins = 0
      for (const t of tabs) {
        const r = t.getBoundingClientRect()
        if (clientX > r.left + r.width / 2) ins++
      }
      return ins
    }
    const onMove = (e: MouseEvent): void => {
      const s = sessionRef.current
      if (!s) return
      if (!s.active) {
        if (Math.abs(e.clientX - s.startX) <= 5 && Math.abs(e.clientY - s.startY) <= 5) return
        s.active = true
      }
      s.ins = calcIns(e.clientX)
      setDrag({ id: s.id, ins: s.ins })
    }
    const finish = (cancel: boolean): void => {
      const s = sessionRef.current
      if (!s) return
      sessionRef.current = null
      setDrag(null)
      if (cancel || !s.active) return
      suppressClick.current = true
      // 视觉插入位（含自身）→ 移除自身后的数组下标
      const to = s.ins > s.from ? s.ins - 1 : s.ins
      if (to !== s.from) view.dispatch(view.state.tr.moveSheet(s.id, to))
      view.focus()
    }
    const onUp = (e: MouseEvent): void => {
      // 拖出标签栏区域松开 = 取消（按指针坐标判定落点）
      const bar = barRef.current
      const r = bar?.getBoundingClientRect()
      const outside =
        !r || e.clientX < r.left || e.clientX > r.right || e.clientY < r.top || e.clientY > r.bottom
      finish(outside)
    }
    const onKey = (e: KeyboardEvent): void => {
      if (e.key === 'Escape' && sessionRef.current?.active) finish(true)
    }
    window.addEventListener('mousemove', onMove)
    window.addEventListener('mouseup', onUp)
    window.addEventListener('keydown', onKey)
    return () => {
      window.removeEventListener('mousemove', onMove)
      window.removeEventListener('mouseup', onUp)
      window.removeEventListener('keydown', onKey)
    }
  }, [view])

  const switchTo = (id: string): void => {
    if (suppressClick.current) {
      suppressClick.current = false
      return
    }
    if (id === state.doc.active) return
    view.dispatch(
      state.tr.setActiveSheet(id).setSelection(singleCell(0, 0)).setMeta('addToHistory', false),
    )
    view.focus()
  }

  const startRename = (id: string): void => {
    setEditingId(id)
    setDraft(view.state.doc.names.get(id) ?? '')
  }

  // 右键菜单「重命名」请求：进入改名输入态后清 key（不入 undo）
  const renameReq = state.getField(tabRenameKey) as string | null | undefined
  useEffect(() => {
    if (!renameReq) return
    if (view.state.doc.order.includes(renameReq)) startRename(renameReq)
    view.dispatch(view.state.tr.setMeta(tabRenameKey, null).setMeta('addToHistory', false))
    // eslint-disable-next-line react-hooks/exhaustive-deps
  }, [renameReq, view])

  const commitRename = (id: string): void => {
    setEditingId(null)
    renameSheet(view, id, draft)
  }

  return (
    <div className="sheet-tab-bar" ref={barRef}>
      <button className="sheet-tab-add" title="新增工作表" onClick={() => addSheet(view)}>
        +
      </button>
      {state.doc.order.map((id, index) => {
        const name = state.doc.names.get(id) ?? id
        const active = id === state.doc.active
        const indicator = drag && drag.ins === index ? (
          <div className="sheet-tab-indicator" />
        ) : null
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
          <Fragment key={id}>
            {indicator}
            <div
              data-sheet-id={id}
              className={active ? 'sheet-tab active' : 'sheet-tab'}
              onMouseDown={(e) => {
                // 新按压作废旧抑制标记（防拖拽 mouseup 落在非 tab 元素上时标记残留吞掉下次单击）
                suppressClick.current = false
                if (e.button !== 0) return
                if ((e.target as HTMLElement).closest('.sheet-tab-close')) return
                sessionRef.current = {
                  id,
                  from: index,
                  startX: e.clientX,
                  startY: e.clientY,
                  active: false,
                  ins: index,
                }
              }}
              onClick={() => switchTo(id)}
              onDoubleClick={() => startRename(id)}
              onContextMenu={(e) => {
                e.preventDefault()
                view.dispatch(
                  view.state.tr
                    .setMeta(contextMenuKey, { kind: 'tab', x: e.clientX, y: e.clientY, row: -1, col: -1, sheet: id })
                    .setMeta('addToHistory', false),
                )
              }}
            >
              <span>{name}</span>
              {state.doc.order.length > 1 ? (
                <button
                  className="sheet-tab-close"
                  title="删除工作表"
                  onClick={(e) => {
                    e.stopPropagation()
                    removeSheet(view, id, name)
                  }}
                >
                  ×
                </button>
              ) : null}
            </div>
          </Fragment>
        )
      })}
      {drag && drag.ins === state.doc.order.length ? (
        <div className="sheet-tab-indicator" />
      ) : null}
    </div>
  )
}
