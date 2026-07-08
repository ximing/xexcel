// 工具栏下拉菜单构建：纯函数，disabled/active 规则在此集中，动作由调用方 handlers 注入
import { ArrowDownAZ, ArrowDownZA, SlidersHorizontal } from 'lucide-react'
import type { MenuEntry } from './ui/Menu'
import type { Locale } from './i18n'
import { t } from './i18n'

export type MenuHandlers = Record<string, () => void>

export function buildRowColItems(
  sel: { fullRow: boolean; fullCol: boolean; canUnhide: boolean; canReset: boolean },
  h: MenuHandlers,
  locale: Locale = 'zh',
): MenuEntry[] {
  return [
    { id: 'insertRow', label: t(locale, 'rowcol.insertRow'), disabled: !sel.fullRow, onSelect: h.insertRow },
    { id: 'deleteRow', label: t(locale, 'rowcol.deleteRow'), disabled: !sel.fullRow, onSelect: h.deleteRow },
    { sep: true },
    { id: 'insertCol', label: t(locale, 'rowcol.insertCol'), disabled: !sel.fullCol, onSelect: h.insertCol },
    { id: 'deleteCol', label: t(locale, 'rowcol.deleteCol'), disabled: !sel.fullCol, onSelect: h.deleteCol },
    { sep: true },
    { id: 'hideRow', label: t(locale, 'rowcol.hideRow'), disabled: !sel.fullRow, onSelect: h.hideRow },
    { id: 'hideCol', label: t(locale, 'rowcol.hideCol'), disabled: !sel.fullCol, onSelect: h.hideCol },
    { id: 'unhide', label: t(locale, 'rowcol.unhide'), disabled: !sel.canUnhide, onSelect: h.unhide },
    { sep: true },
    { id: 'resetSize', label: t(locale, 'rowcol.resetSize'), disabled: !sel.canReset, onSelect: h.resetSize },
  ]
}

export function buildFreezeItems(
  frozen: { rows: number; cols: number },
  h: MenuHandlers,
  locale: Locale = 'zh',
): MenuEntry[] {
  const none = frozen.rows === 0 && frozen.cols === 0
  return [
    { id: 'freezeRow', label: t(locale, 'freeze.row'), active: frozen.rows === 1 && frozen.cols === 0, onSelect: h.freezeRow },
    { id: 'freezeCol', label: t(locale, 'freeze.col'), active: frozen.cols === 1 && frozen.rows === 0, onSelect: h.freezeCol },
    { id: 'freezeTo', label: t(locale, 'freeze.to'), onSelect: h.freezeTo },
    { sep: true },
    { id: 'unfreeze', label: t(locale, 'freeze.off'), disabled: none, onSelect: h.unfreeze },
  ]
}

export function buildSortItems(canSortNow: boolean, h: MenuHandlers, locale: Locale = 'zh'): MenuEntry[] {
  return [
    { id: 'asc', label: t(locale, 'sort.asc'), icon: ArrowDownAZ, disabled: !canSortNow, onSelect: h.asc },
    { id: 'desc', label: t(locale, 'sort.desc'), icon: ArrowDownZA, disabled: !canSortNow, onSelect: h.desc },
    { sep: true },
    { id: 'custom', label: t(locale, 'sort.custom'), icon: SlidersHorizontal, disabled: !canSortNow, onSelect: h.custom },
  ]
}

export function buildNumberFormatItems(
  numFmt: string | undefined,
  h: MenuHandlers,
  locale: Locale = 'zh',
): MenuEntry[] {
  return [
    { id: 'thousands', label: t(locale, 'num.thousands'), active: numFmt === '#,##0.00', onSelect: h.thousands },
    { id: 'thousandsInt', label: t(locale, 'num.thousandsInt'), active: numFmt === '#,##0', onSelect: h.thousandsInt },
    { id: 'percent', label: t(locale, 'num.percent'), active: numFmt === '0%', onSelect: h.percent },
    { id: 'currency', label: t(locale, 'num.currency'), active: numFmt === '¥#,##0.00', onSelect: h.currency },
    { id: 'date', label: t(locale, 'num.date'), active: numFmt === 'yyyy-mm-dd', onSelect: h.date },
    { id: 'accounting', label: t(locale, 'num.accounting'), active: numFmt === '¥* #,##0.00', onSelect: h.accounting },
    { sep: true },
    { id: 'decInc', label: t(locale, 'num.decInc'), onSelect: h.decInc },
    { id: 'decDec', label: t(locale, 'num.decDec'), onSelect: h.decDec },
  ]
}
