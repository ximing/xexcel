import { describe, expect, it } from 'vitest'
import { SheetState } from '../src/core/state'
import { Workbook } from '../src/core/model'
import { formatPainterKey } from '../src/view/types'
import { metaField } from '../src/plugins/uistate'

// metaField 经 tr meta 透传：激活/锁定/解除（显式注册 metaField，getField 默认 null）
const mkState = () =>
  SheetState.create({
    doc: Workbook.create({ rowCount: 10, colCount: 10 }),
    plugins: [metaField(formatPainterKey, null)],
  })

describe('formatPainter state field', () => {
  it('默认 null；setMeta 设置；不带 meta 的事务保持', () => {
    let s = mkState()
    expect(s.getField(formatPainterKey)).toBeNull()
    const fp = { style: { bold: true }, locked: false }
    s = s.applyTransaction(s.tr.setMeta(formatPainterKey, fp).setMeta('addToHistory', false)).state
    expect(s.getField(formatPainterKey)).toEqual(fp)
    s = s.applyTransaction(s.tr.setCell(0, 0, 'x')).state
    expect(s.getField(formatPainterKey)).toEqual(fp)
  })
})

describe('格式刷应用', () => {
  it('setCellStyles 整体替换目标 style（含 border/numFmt），非合并', () => {
    let s = mkState()
    s = s.applyTransaction(
      s.tr.setCell(0, 0, 'a', { bold: true, numFmt: '0%' }).setCell(1, 1, 'b', { italic: true, bg: '#fff' }),
    ).state
    const src = s.activeSheet.getCell(0, 0)!.style!
    s = s.applyTransaction(s.tr.setCellStyles([{ row: 1, col: 1, style: { ...src } }])).state
    expect(s.activeSheet.getCell(1, 1)!.style).toEqual({ bold: true, numFmt: '0%' })
    // undo 恢复
    s = s.applyTransaction(s.tr.setCellStyles([{ row: 1, col: 1, style: { italic: true, bg: '#fff' } }])).state
    expect(s.activeSheet.getCell(1, 1)!.style).toEqual({ italic: true, bg: '#fff' })
  })
})
