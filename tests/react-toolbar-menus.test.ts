// toolbarMenus 构建纯函数：disabled/active 映射断言（逐项 find，不钉死数组顺序）
import { describe, expect, it } from 'vitest'
import { buildFreezeItems, buildNumberFormatItems, buildRowColItems, buildSortItems } from '../src/react/toolbarMenus'
import type { MenuEntry } from '../src/react/ui/Menu'

const noop = (): void => {}
const H = {
  insertRow: noop, deleteRow: noop, insertCol: noop, deleteCol: noop,
  hideRow: noop, hideCol: noop, unhide: noop, resetSize: noop,
  freezeRow: noop, freezeCol: noop, freezeTo: noop, unfreeze: noop,
  asc: noop, desc: noop, custom: noop,
  thousands: noop, percent: noop, currency: noop, decInc: noop, decDec: noop,
}

function item(entries: MenuEntry[], id: string): MenuEntry | undefined {
  return entries.find((e) => !('sep' in e) && e.id === id)
}
function disabled(entries: MenuEntry[], id: string): boolean {
  const e = item(entries, id)
  return !!(e && !('sep' in e) && e.disabled)
}
function active(entries: MenuEntry[], id: string): boolean {
  const e = item(entries, id)
  return !!(e && !('sep' in e) && e.active)
}

describe('buildRowColItems', () => {
  it('整行选区：行项可用、列项禁用', () => {
    const m = buildRowColItems({ fullRow: true, fullCol: false, canUnhide: false }, H)
    expect(disabled(m, 'insertRow')).toBe(false)
    expect(disabled(m, 'deleteRow')).toBe(false)
    expect(disabled(m, 'insertCol')).toBe(true)
    expect(disabled(m, 'deleteCol')).toBe(true)
    expect(disabled(m, 'hideRow')).toBe(false)
    expect(disabled(m, 'hideCol')).toBe(true)
    expect(disabled(m, 'unhide')).toBe(true)
    expect(disabled(m, 'resetSize')).toBe(false)
  })
  it('非整行整列：行列结构项全禁用，重置禁用，可取消隐藏', () => {
    const m = buildRowColItems({ fullRow: false, fullCol: false, canUnhide: true }, H)
    expect(disabled(m, 'insertRow')).toBe(true)
    expect(disabled(m, 'insertCol')).toBe(true)
    expect(disabled(m, 'hideRow')).toBe(true)
    expect(disabled(m, 'hideCol')).toBe(true)
    expect(disabled(m, 'resetSize')).toBe(true)
    expect(disabled(m, 'unhide')).toBe(false)
  })
})

describe('buildFreezeItems', () => {
  it('冻结首行激活；无冻结时取消禁用', () => {
    const items = buildFreezeItems({ rows: 1, cols: 0 }, H)
    expect(active(items, 'freezeRow')).toBe(true)
    expect(disabled(items, 'unfreeze')).toBe(false)
    const none = buildFreezeItems({ rows: 0, cols: 0 }, H)
    expect(active(none, 'freezeRow')).toBe(false)
    expect(disabled(none, 'unfreeze')).toBe(true)
  })
  it('冻结首列激活', () => {
    const items = buildFreezeItems({ rows: 0, cols: 1 }, H)
    expect(active(items, 'freezeCol')).toBe(true)
  })
})

describe('buildSortItems', () => {
  it('不可排序时三项全禁用', () => {
    const items = buildSortItems(false, H)
    expect(items.every((e) => 'sep' in e || e.disabled)).toBe(true)
  })
  it('可排序时三项可用', () => {
    const items = buildSortItems(true, H)
    expect(disabled(items, 'asc')).toBe(false)
    expect(disabled(items, 'desc')).toBe(false)
    expect(disabled(items, 'custom')).toBe(false)
  })
})

describe('buildNumberFormatItems', () => {
  it('当前 numFmt 项 active', () => {
    const items = buildNumberFormatItems('0%', H)
    expect(active(items, 'percent')).toBe(true)
    expect(active(items, 'thousands')).toBe(false)
  })
  it('无 numFmt 时无 active 项', () => {
    const items = buildNumberFormatItems(undefined, H)
    expect(items.every((e) => 'sep' in e || !e.active)).toBe(true)
  })
})
