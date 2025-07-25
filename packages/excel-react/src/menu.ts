// 右键菜单项推导（纯逻辑，node 可测）：按 kind + 当前 state 计算可见项与禁用态。
import { selectionRange } from '@gmi/excel-core'
import type { SheetState } from '@gmi/excel-core'
import type { ContextMenuOpen } from '@gmi/excel-view'

export interface MenuItemSpec {
  id: string
  label: string
  disabled: boolean
  sep?: boolean // 该项前画分隔线
}

export function menuItems(state: SheetState, open: ContextMenuOpen): MenuItemSpec[] {
  const r = selectionRange(state.selection)
  const sheet = state.activeSheet
  const hiddenRowsInSel = sheet.hiddenRows.some((i) => i >= r.sr && i <= r.er)
  const hiddenColsInSel = sheet.hiddenCols.some((i) => i >= r.sc && i <= r.ec)
  switch (open.kind) {
    case 'cell':
      return [
        { id: 'cut', label: '剪切', disabled: false },
        { id: 'copy', label: '复制', disabled: false },
        { id: 'paste', label: '粘贴', disabled: false },
        { id: 'insertRows', label: '插入行', disabled: false, sep: true },
        { id: 'insertCols', label: '插入列', disabled: false },
        { id: 'deleteRows', label: '删除行', disabled: r.er - r.sr + 1 >= sheet.rowCount },
        { id: 'deleteCols', label: '删除列', disabled: r.ec - r.sc + 1 >= sheet.colCount },
        { id: 'hideRows', label: '隐藏行', disabled: false, sep: true },
        { id: 'hideCols', label: '隐藏列', disabled: false },
        { id: 'unhide', label: '取消隐藏', disabled: !hiddenRowsInSel && !hiddenColsInSel },
        { id: 'clear', label: '清除内容', disabled: false, sep: true },
      ]
    case 'rowheader':
      return [
        { id: 'insertRows', label: '插入行', disabled: false },
        { id: 'deleteRows', label: '删除行', disabled: r.er - r.sr + 1 >= sheet.rowCount },
        { id: 'hideRows', label: '隐藏行', disabled: false, sep: true },
        { id: 'unhide', label: '取消隐藏', disabled: !hiddenRowsInSel },
      ]
    case 'colheader':
      return [
        { id: 'insertCols', label: '插入列', disabled: false },
        { id: 'deleteCols', label: '删除列', disabled: r.ec - r.sc + 1 >= sheet.colCount },
        { id: 'hideCols', label: '隐藏列', disabled: false, sep: true },
        { id: 'unhide', label: '取消隐藏', disabled: !hiddenColsInSel },
      ]
    case 'tab': {
      const idx = state.doc.order.indexOf(open.sheet!)
      return [
        { id: 'tabAdd', label: '新建工作表', disabled: false },
        { id: 'tabRename', label: '重命名', disabled: false },
        { id: 'tabRemove', label: '删除', disabled: state.doc.order.length <= 1 },
        { id: 'tabLeft', label: '左移', disabled: idx <= 0, sep: true },
        { id: 'tabRight', label: '右移', disabled: idx < 0 || idx >= state.doc.order.length - 1 },
      ]
    }
  }
}
