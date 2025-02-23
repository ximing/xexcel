// 右键菜单浮层：订阅 contextMenuKey，按 menuItems 渲染；动作走现有命令/事务。
// 剪切/复制经 execCommand 触发 proxy 的 copy/cut 事件（clipboard 插件接管）；
// 粘贴经 navigator.clipboard.readText 后走 view.someProp('handlePaste')。
import { Fragment, useEffect } from 'react'
import { showNotice } from '../app/notice'
import { selectionRange } from '../core/selection'
import type { EditorView } from '../view/editorview'
import { contextMenuKey, ContextMenuOpen, tabRenameKey } from '../view/types'
import { useSheetState } from './bridge'
import { menuItems } from './menu'
import { addSheet, removeSheet } from './sheetOps'

export function ContextMenu({ view }: { view: EditorView }) {
  const state = useSheetState(view)
  const open = state.getField(contextMenuKey) as ContextMenuOpen | null | undefined

  const close = (): void => {
    view.dispatch(view.state.tr.setMeta(contextMenuKey, null).setMeta('addToHistory', false))
  }

  // Esc 关闭（仅打开期间挂载监听）
  useEffect(() => {
    if (!open) return
    const onKey = (e: KeyboardEvent): void => {
      if (e.key === 'Escape') close()
    }
    window.addEventListener('keydown', onKey)
    return () => window.removeEventListener('keydown', onKey)
    // eslint-disable-next-line react-hooks/exhaustive-deps
  }, [open !== null])

  if (!open) return null
  const items = menuItems(state, open)

  const act = (id: string): void => {
    const st = view.state
    const r = selectionRange(st.selection)
    const sheet = st.activeSheet
    switch (id) {
      case 'cut':
      case 'copy':
        view.focus()
        document.execCommand(id)
        break
      case 'paste':
        navigator.clipboard
          .readText()
          .then((t) => {
            if (t) view.someProp('handlePaste', (p) => p(view, t))
            view.focus()
          })
          .catch(() => showNotice('无法读取剪贴板'))
        break
      case 'insertRows':
        // 数量 = 选区行跨度（非整行选也按跨度，对齐 Excel）
        view.dispatch(st.tr.structure('row', r.sr, r.er - r.sr + 1, 'insert'))
        view.focus()
        break
      case 'insertCols':
        view.dispatch(st.tr.structure('col', r.sc, r.ec - r.sc + 1, 'insert'))
        view.focus()
        break
      case 'deleteRows':
        view.dispatch(st.tr.structure('row', r.sr, r.er - r.sr + 1, 'delete'))
        view.focus()
        break
      case 'deleteCols':
        view.dispatch(st.tr.structure('col', r.sc, r.ec - r.sc + 1, 'delete'))
        view.focus()
        break
      case 'hideRows':
      case 'hideCols': {
        const axis = id === 'hideRows' ? 'row' : 'col'
        const indices: number[] = []
        for (let i = axis === 'row' ? r.sr : r.sc; i <= (axis === 'row' ? r.er : r.ec); i++) indices.push(i)
        view.dispatch(st.tr.setHidden(axis, indices, true))
        view.focus()
        break
      }
      case 'unhide': {
        // 行头菜单只恢复行、列头菜单只恢复列，避免整行/列选区误伤另一轴
        const tr = st.tr
        let touched = false
        if (open.kind !== 'colheader') {
          const rows = sheet.hiddenRows.filter((i) => i >= r.sr && i <= r.er)
          if (rows.length) {
            tr.setHidden('row', rows, false)
            touched = true
          }
        }
        if (open.kind !== 'rowheader') {
          const cols = sheet.hiddenCols.filter((i) => i >= r.sc && i <= r.ec)
          if (cols.length) {
            tr.setHidden('col', cols, false)
            touched = true
          }
        }
        if (touched) view.dispatch(tr)
        view.focus()
        break
      }
      case 'clear':
        view.dispatch(st.tr.clearRange(r))
        view.focus()
        break
      case 'tabAdd':
        addSheet(view)
        break
      case 'tabRemove':
        void removeSheet(view, open.sheet!, st.doc.names.get(open.sheet!) ?? open.sheet!)
        break
      case 'tabRename':
        // 改名输入态由 SheetTabBar 订阅 tabRenameKey 进入
        view.dispatch(st.tr.setMeta(tabRenameKey, open.sheet).setMeta('addToHistory', false))
        break
      case 'tabLeft':
      case 'tabRight': {
        const idx = st.doc.order.indexOf(open.sheet!)
        view.dispatch(st.tr.moveSheet(open.sheet!, id === 'tabLeft' ? idx - 1 : idx + 1))
        view.focus()
        break
      }
    }
    close()
  }

  // 定位 clamp 到窗口内
  const x = Math.max(0, Math.min(open.x, window.innerWidth - 180))
  const y = Math.max(0, Math.min(open.y, window.innerHeight - items.length * 28 - 16))

  return (
    <div
      className="fixed inset-0 z-110"
      onMouseDown={close}
      onMouseUp={(e) => e.stopPropagation()} // 双保险：阻断 window 级 mouseup 穿透到 painter 等插件
      onContextMenu={(e) => {
        e.preventDefault()
        close()
      }}
    >
      <div
        className="fixed flex min-w-40 flex-col rounded-md border border-line-strong bg-surface py-1 shadow-2"
        style={{ left: x, top: y }}
        onMouseDown={(e) => e.stopPropagation()}
        onMouseUp={(e) => e.stopPropagation()}
      >
        {items.map((it) => (
          <Fragment key={it.id}>
            {it.sep ? <div className="my-1 h-px bg-line" /> : null}
            <button
              type="button"
              className={[
                'h-7 px-4 text-left text-xs',
                it.disabled ? 'cursor-default text-ink-disabled' : 'text-ink hover:bg-primary-soft',
              ].join(' ')}
              disabled={it.disabled}
              onClick={() => act(it.id)}
            >
              {it.label}
            </button>
          </Fragment>
        ))}
      </div>
    </div>
  )
}
