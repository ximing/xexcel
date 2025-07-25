import { describe, expect, it } from 'vitest'
import { undoDepth } from '@gmi/excel-core'
import { Workbook } from '@gmi/excel-core'
import { createStateFromWorkbook } from '@gmi/excel-react'

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
