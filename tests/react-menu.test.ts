import { describe, expect, it } from 'vitest'
import { SheetData, Workbook } from '@gmi/excel-core'
import { rangeSelection } from '@gmi/excel-core'
import type { Selection } from '@gmi/excel-core'
import { SheetState } from '@gmi/excel-core'
import { menuItems } from '../src/react/menu'
import type { ContextMenuOpen } from '../src/view/types'

const mk = (doc?: Workbook, selection?: Selection): SheetState =>
  SheetState.create({ doc: doc ?? Workbook.create({ rowCount: 10, colCount: 10 }), selection })

const cellOpen: ContextMenuOpen = { kind: 'cell', x: 0, y: 0, row: 0, col: 0 }
const tabOpen = (sheet: string): ContextMenuOpen => ({ kind: 'tab', x: 0, y: 0, row: -1, col: -1, sheet })

const byId = (items: ReturnType<typeof menuItems>, id: string) => items.find((i) => i.id === id)!

describe('menuItems cell', () => {
  it('全集且顺序固定', () => {
    const items = menuItems(mk(), cellOpen)
    expect(items.map((i) => i.id)).toEqual([
      'cut', 'copy', 'paste',
      'insertRows', 'insertCols', 'deleteRows', 'deleteCols',
      'hideRows', 'hideCols', 'unhide',
      'clear',
    ])
  })

  it('选区内无隐藏行列时 unhide 禁用；有隐藏行/列时启用', () => {
    expect(byId(menuItems(mk(), cellOpen), 'unhide').disabled).toBe(true)
    const wb = Workbook.create({ rowCount: 10, colCount: 10 })
    const docRows = wb.setSheet('s1', wb.activeSheet.setHidden('row', [2], true))
    const sel = rangeSelection({ sr: 1, sc: 0, er: 3, ec: 0 })
    expect(byId(menuItems(mk(docRows, sel), cellOpen), 'unhide').disabled).toBe(false)
    const docCols = wb.setSheet('s1', wb.activeSheet.setHidden('col', [1], true))
    const selCols = rangeSelection({ sr: 0, sc: 0, er: 0, ec: 2 })
    expect(byId(menuItems(mk(docCols, selCols), cellOpen), 'unhide').disabled).toBe(false)
  })

  it('全表选中时 deleteRows/deleteCols 禁用；部分选区可用', () => {
    const all = rangeSelection({ sr: 0, sc: 0, er: 9, ec: 9 })
    const items = menuItems(mk(undefined, all), cellOpen)
    expect(byId(items, 'deleteRows').disabled).toBe(true)
    expect(byId(items, 'deleteCols').disabled).toBe(true)
    const partial = menuItems(mk(), cellOpen)
    expect(byId(partial, 'deleteRows').disabled).toBe(false)
    expect(byId(partial, 'deleteCols').disabled).toBe(false)
  })
})

describe('menuItems tab', () => {
  const threeSheets = (): Workbook => {
    const cfg = { rowCount: 10, colCount: 10 }
    return Workbook.create(cfg)
      .addSheet('s2', SheetData.create(cfg))
      .addSheet('s3', SheetData.create(cfg))
  }

  it('首项左移禁用，末项右移禁用，中间项均可', () => {
    const doc = threeSheets()
    expect(byId(menuItems(mk(doc), tabOpen('s1')), 'tabLeft').disabled).toBe(true)
    expect(byId(menuItems(mk(doc), tabOpen('s1')), 'tabRight').disabled).toBe(false)
    expect(byId(menuItems(mk(doc), tabOpen('s3')), 'tabRight').disabled).toBe(true)
    expect(byId(menuItems(mk(doc), tabOpen('s3')), 'tabLeft').disabled).toBe(false)
    const mid = menuItems(mk(doc), tabOpen('s2'))
    expect(byId(mid, 'tabLeft').disabled).toBe(false)
    expect(byId(mid, 'tabRight').disabled).toBe(false)
  })

  it('单表时删除禁用', () => {
    expect(byId(menuItems(mk(), tabOpen('s1')), 'tabRemove').disabled).toBe(true)
    expect(byId(menuItems(mk(threeSheets()), tabOpen('s1')), 'tabRemove').disabled).toBe(false)
  })
})
