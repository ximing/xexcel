// 工具栏下拉菜单构建：纯函数，disabled/active 规则在此集中，动作由调用方 handlers 注入
import { ArrowDownAZ, ArrowDownZA, SlidersHorizontal } from 'lucide-react'
import type { MenuEntry } from './ui/Menu'

export type MenuHandlers = Record<string, () => void>

export function buildRowColItems(
  sel: { fullRow: boolean; fullCol: boolean; canUnhide: boolean; canReset: boolean },
  h: MenuHandlers,
): MenuEntry[] {
  return [
    { id: 'insertRow', label: '上方插入行', disabled: !sel.fullRow, onSelect: h.insertRow },
    { id: 'deleteRow', label: '删除选中行', disabled: !sel.fullRow, onSelect: h.deleteRow },
    { sep: true },
    { id: 'insertCol', label: '左侧插入列', disabled: !sel.fullCol, onSelect: h.insertCol },
    { id: 'deleteCol', label: '删除选中列', disabled: !sel.fullCol, onSelect: h.deleteCol },
    { sep: true },
    { id: 'hideRow', label: '隐藏选中行', disabled: !sel.fullRow, onSelect: h.hideRow },
    { id: 'hideCol', label: '隐藏选中列', disabled: !sel.fullCol, onSelect: h.hideCol },
    { id: 'unhide', label: '取消隐藏', disabled: !sel.canUnhide, onSelect: h.unhide },
    { sep: true },
    { id: 'resetSize', label: '重置行高列宽', disabled: !sel.canReset, onSelect: h.resetSize },
  ]
}

export function buildFreezeItems(frozen: { rows: number; cols: number }, h: MenuHandlers): MenuEntry[] {
  const none = frozen.rows === 0 && frozen.cols === 0
  return [
    { id: 'freezeRow', label: '冻结首行', active: frozen.rows === 1 && frozen.cols === 0, onSelect: h.freezeRow },
    { id: 'freezeCol', label: '冻结首列', active: frozen.cols === 1 && frozen.rows === 0, onSelect: h.freezeCol },
    { id: 'freezeTo', label: '冻结到当前选区', onSelect: h.freezeTo },
    { sep: true },
    { id: 'unfreeze', label: '取消冻结', disabled: none, onSelect: h.unfreeze },
  ]
}

export function buildSortItems(canSortNow: boolean, h: MenuHandlers): MenuEntry[] {
  return [
    { id: 'asc', label: '按选区首列升序', icon: ArrowDownAZ, disabled: !canSortNow, onSelect: h.asc },
    { id: 'desc', label: '按选区首列降序', icon: ArrowDownZA, disabled: !canSortNow, onSelect: h.desc },
    { sep: true },
    { id: 'custom', label: '自定义排序…', icon: SlidersHorizontal, disabled: !canSortNow, onSelect: h.custom },
  ]
}

export function buildNumberFormatItems(numFmt: string | undefined, h: MenuHandlers): MenuEntry[] {
  return [
    { id: 'thousands', label: '千分位', active: numFmt === '#,##0.00', onSelect: h.thousands },
    { id: 'percent', label: '百分比', active: numFmt === '0%', onSelect: h.percent },
    { id: 'currency', label: '货币', active: numFmt === '¥#,##0.00', onSelect: h.currency },
    { sep: true },
    { id: 'decInc', label: '增加小数位', onSelect: h.decInc },
    { id: 'decDec', label: '减少小数位', onSelect: h.decDec },
  ]
}
