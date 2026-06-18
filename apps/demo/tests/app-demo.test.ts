import { describe, expect, it } from 'vitest'
import { undoDepth } from '@xexcel/core'
import { Workbook } from '@xexcel/core'
import { createStateFromWorkbook } from '@xexcel/react'
import { createDemoState } from '../src/demo'

describe('createStateFromWorkbook', () => {
  it('以给定 workbook 建 state，history 等插件就位', () => {
    const wb = Workbook.create({ rowCount: 5, colCount: 5 }).renameSheet('s1', '存档表')
    const s = createStateFromWorkbook(wb)
    expect(s.doc.names.get('s1')).toBe('存档表')
    // history 插件已注册：一个 setCell 事务后 undoDepth=1
    const s2 = s.applyTransaction(s.tr.setCell(0, 0, 'x')).state
    expect(undoDepth(s2)).toBe(1)
  })
})

describe('createDemoState', () => {
  it('三张展示表，销售表 A1 为产品', () => {
    const s = createDemoState()
    const names = s.doc.order.map((id) => s.doc.names.get(id))
    expect(names).toEqual(['销售', '样式', '试用'])
    expect(s.doc.activeSheet.getCell(0, 0)?.raw).toBe('产品')
    expect(s.doc.activeSheet.getCell(1, 3)?.raw).toBe('=B2*C2')
    expect(s.doc.activeSheet.filter).toBeTruthy()
    expect(s.doc.activeSheet.condFormats.length).toBeGreaterThan(0)
  })
})
