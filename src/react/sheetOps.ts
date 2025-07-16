// 工作表增/删/改名共用逻辑（SheetTabBar 与右键菜单共用）。
// 空表跳过删除确认；删 active 表重置选区到 A1；重名 notice 不派发。
import { nextSheetId, nextSheetName, SheetData, SheetId } from '@gmi/excel-core'
import { singleCell } from '@gmi/excel-core'
import { isValidSheetName, SHEET_NAME_MAX_LEN } from '@gmi/excel-core'
import type { EditorView } from '@gmi/excel-view'
import { showNotice } from '../app/notice'
import { askConfirm } from './ui/confirmStore'

// 空表判定：usedRange 为单格且该格无内容（usedRange 对空表返回全 0）
export function isSheetEmpty(data: SheetData): boolean {
  const r = data.usedRange()
  return r.sr === 0 && r.sc === 0 && r.er === 0 && r.ec === 0 && !data.getCell(0, 0)
}

// 用 view.state（始终最新）而非渲染快照：快速连点时 hook 快照可能过期，
// 过期快照会算出重复 id 导致插入变 no-op（e2e 观察 a）
export function addSheet(view: EditorView): void {
  const st = view.state
  const id = nextSheetId(st.doc)
  const name = nextSheetName(st.doc)
  const config = { rowCount: st.activeSheet.rowCount, colCount: st.activeSheet.colCount }
  view.dispatch(st.tr.insertSheet(id, name, config).setSelection(singleCell(0, 0)))
  view.focus()
}

export async function removeSheet(view: EditorView, id: SheetId, name: string): Promise<void> {
  const st = view.state
  if (st.doc.order.length <= 1) return
  if (!isSheetEmpty(st.doc.sheet(id))) {
    const ok = await askConfirm({
      title: '删除工作表',
      body: `确定删除工作表「${name}」？可通过撤销恢复。`,
      danger: true,
      confirmLabel: '删除',
    })
    if (!ok) return
  }
  // 确认框挂起期间可能有其他编辑，事务基于最新 state 重建
  const cur = view.state
  if (cur.doc.order.length <= 1 || !cur.doc.order.includes(id)) return
  const tr = cur.tr.removeSheet(id)
  if (id === cur.doc.active) tr.setSelection(singleCell(0, 0))
  view.dispatch(tr)
  view.focus()
}

export function renameSheet(view: EditorView, id: SheetId, name: string): void {
  const st = view.state
  const trimmed = name.trim()
  if (trimmed === '' || trimmed === st.doc.names.get(id)) return
  // 非法输入拒绝 + 提示（对齐 Excel 弹错），不自动改写
  if (!isValidSheetName(trimmed)) {
    showNotice(`无效的工作表名称：不能包含 * ? : \\ / [ ]，且不超过 ${SHEET_NAME_MAX_LEN} 个字符`)
    return
  }
  const dup = [...st.doc.names.entries()].some(
    ([other, n]) => other !== id && n.toLowerCase() === trimmed.toLowerCase(),
  )
  if (dup) {
    showNotice(`工作表名称重复：${trimmed}`)
    return
  }
  view.dispatch(st.tr.renameSheet(id, trimmed))
  view.focus()
}
