// 右键菜单项推导（纯逻辑，node 可测）：按 kind + 当前 state 计算可见项与禁用态。
import { selectionRange } from '@xexcel/core'
import type { SheetState } from '@xexcel/core'
import type { ContextMenuOpen } from '@xexcel/view'
import type { Locale } from './i18n'
import { t } from './i18n'

export interface MenuItemSpec {
  id: string
  label: string
  disabled: boolean
  sep?: boolean // 该项前画分隔线
}

export function menuItems(state: SheetState, open: ContextMenuOpen, locale: Locale = 'zh'): MenuItemSpec[] {
  const r = selectionRange(state.selection)
  const sheet = state.activeSheet
  const hiddenRowsInSel = sheet.hiddenRows.some((i) => i >= r.sr && i <= r.er)
  const hiddenColsInSel = sheet.hiddenCols.some((i) => i >= r.sc && i <= r.ec)
  switch (open.kind) {
    case 'cell':
      return [
        { id: 'cut', label: t(locale, 'menu.cut'), disabled: false },
        { id: 'copy', label: t(locale, 'menu.copy'), disabled: false },
        { id: 'paste', label: t(locale, 'menu.paste'), disabled: false },
        { id: 'insertRows', label: t(locale, 'menu.insertRows'), disabled: false, sep: true },
        { id: 'insertCols', label: t(locale, 'menu.insertCols'), disabled: false },
        { id: 'deleteRows', label: t(locale, 'menu.deleteRows'), disabled: r.er - r.sr + 1 >= sheet.rowCount },
        { id: 'deleteCols', label: t(locale, 'menu.deleteCols'), disabled: r.ec - r.sc + 1 >= sheet.colCount },
        { id: 'hideRows', label: t(locale, 'menu.hideRows'), disabled: false, sep: true },
        { id: 'hideCols', label: t(locale, 'menu.hideCols'), disabled: false },
        { id: 'unhide', label: t(locale, 'menu.unhide'), disabled: !hiddenRowsInSel && !hiddenColsInSel },
        { id: 'clear', label: t(locale, 'menu.clear'), disabled: false, sep: true },
      ]
    case 'rowheader':
      return [
        { id: 'insertRows', label: t(locale, 'menu.insertRows'), disabled: false },
        { id: 'deleteRows', label: t(locale, 'menu.deleteRows'), disabled: r.er - r.sr + 1 >= sheet.rowCount },
        { id: 'hideRows', label: t(locale, 'menu.hideRows'), disabled: false, sep: true },
        { id: 'unhide', label: t(locale, 'menu.unhide'), disabled: !hiddenRowsInSel },
      ]
    case 'colheader':
      return [
        { id: 'insertCols', label: t(locale, 'menu.insertCols'), disabled: false },
        { id: 'deleteCols', label: t(locale, 'menu.deleteCols'), disabled: r.ec - r.sc + 1 >= sheet.colCount },
        { id: 'hideCols', label: t(locale, 'menu.hideCols'), disabled: false, sep: true },
        { id: 'unhide', label: t(locale, 'menu.unhide'), disabled: !hiddenColsInSel },
      ]
    case 'tab': {
      const idx = state.doc.order.indexOf(open.sheet!)
      return [
        { id: 'tabAdd', label: t(locale, 'menu.tabAdd'), disabled: false },
        { id: 'tabRename', label: t(locale, 'menu.tabRename'), disabled: false },
        { id: 'tabRemove', label: t(locale, 'menu.tabRemove'), disabled: state.doc.order.length <= 1 },
        { id: 'tabLeft', label: t(locale, 'menu.tabLeft'), disabled: idx <= 0, sep: true },
        { id: 'tabRight', label: t(locale, 'menu.tabRight'), disabled: idx < 0 || idx >= state.doc.order.length - 1 },
      ]
    }
  }
}
